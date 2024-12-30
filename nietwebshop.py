from datetime import datetime, timedelta
import pandas as pd
import win32com.client
from typing import Dict, Tuple
import os

# Define column mappings
# Update column mappings to include alternatives
COLUMN_NAMES = {
    'ordernr': ['Ordernummer', 'OrderNummer'],
    'orderdate': ['Orderdatum'],
    'customer': ['Klant'],
    'mailed': ['Gemaild', 'gemaild'],
    'delivery_date': ['Leverdatum'],
    'expected_date': ['Verwachte leverdatum'],
    'country': ['Land', 'Land/site'],
    'extra_note': ['extra opmerking', 'Extra info', 'Extra opmerking']  # Optional column
}

# Define signature
SIGNATURE = ("Met vriendelijke groet,\n\n"
            "LampenTotaal\n"
            "Dorpsstraat 2a  |  4043 KK  Opheusden\n"
            "T. 0488.750930  |  www.LampenTotaal.nl\n"
            "IBAN NL65RABO0116060549")

def get_delivery_timing(expected_date: datetime) -> Tuple[bool, int, str]:
    """Return if delay is short-term, week number, and timing in week"""
    today = datetime.now()
    diff = (expected_date - today).days
    weekday = expected_date.weekday()
    
    timing = "in de loop van"
    if weekday <= 1:  # Monday-Tuesday
        timing = "begin"
    elif weekday >= 4:  # Friday-Sunday
        timing = "eind"
    
    return diff <= 7, expected_date.isocalendar()[1], timing

def generate_nl_short_delay(timing: str) -> str:
    return f"""Goedemiddag,

Hartelijk dank voor uw bestelling.

Hierbij meer informatie met betrekking tot de uitlevering van de bestelling die u bij ons hebt geplaatst.

Wij verwachten {timing} volgende week uw pakket te gaan ontvangen van onze leverancier, uiteraard gaan wij ons best doen om uw pakket verder direct te gaan versturen naar uw postadres.
Zodra wij uw pakket hebben verzonden ontvangt u een track en trace code per mail. Hiermee kunt u het pakket volgen.

Wij hopen u voldoende te hebben geïnformeerd. Mocht u nog vragen hebben, mail of bel gerust.

{SIGNATURE}"""

def generate_nl_long_delay(week_number: int) -> str:
    return f"""Goedemiddag,

Hartelijk dank voor uw bestelling.

Hierbij meer informatie met betrekking tot de uitlevering van de bestelling die u bij ons hebt geplaatst.
Helaas hebben wij van de leverancier vernomen dat het artikel wat u besteld heeft momenteel niet op voorraad is. Wij verwachten uw bestelling in week {week_number} binnen te krijgen.

Wij hopen u voldoende te hebben geïnformeerd. Mocht u nog vragen hebben, mail of bel gerust.

Excuses voor het ongemak!

Indien er wellicht een aanpassing komt in de bovenstaande levertijd hoort u dat uiteraard van ons.

{SIGNATURE}"""

def generate_fr_short_delay(timing: str) -> str:
    return f"""Bonjour,

Nous vous remercions de votre commande.

Voici plus d'informations concernant la livraison de votre commande.

Nous prévoyons de recevoir votre colis de notre fournisseur {timing} de la semaine prochaine. Nous ferons de notre mieux pour expédier votre colis immédiatement à votre adresse postale.
Dès que nous aurons expédié votre colis, vous recevrez un code de suivi par e-mail vous permettant de suivre le colis.

Nous espérons vous avoir suffisamment informé. Si vous avez des questions, n'hésitez pas à nous contacter.

Cordialement,
{SIGNATURE}"""

def generate_de_short_delay(timing_de: str) -> str:
    timing_map = {"begin": "Anfang", "eind": "Ende", "in de loop van": "im Laufe"}
    timing = timing_map[timing_de]
    
    return f"""Guten Tag,

Vielen Dank für Ihre Bestellung.

Hiermit möchten wir Sie über den Status Ihrer Bestellung informieren.

Wir erwarten, dass wir Ihr Paket {timing} nächster Woche von unserem Lieferanten erhalten. Selbstverständlich werden wir uns bemühen, Ihr Paket dann umgehend an Ihre Postadresse zu versenden.
Sobald wir Ihr Paket versendet haben, erhalten Sie eine Track & Trace Nummer per E-Mail.

Wir hoffen, Sie ausreichend informiert zu haben. Bei Fragen können Sie uns gerne kontaktieren.

Mit freundlichen Grüßen,
{SIGNATURE}"""

def create_outlook_mail(to_address: str, subject: str, body: str):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = subject
    mail.HTMLBody = body.replace('\n', '<br>')
    return mail

def get_actual_column_name(df: pd.DataFrame, possible_names: list) -> str:
    """Get the actual column name from possible alternatives"""
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def parse_date(date_str) -> datetime:
    """Parse date string in multiple formats"""
    if isinstance(date_str, datetime):
        return date_str
        
    formats = ['%d-%m-%Y', '%Y-%m-%d', '%d-%m-%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S']
    for fmt in formats:
        try:
            return datetime.strptime(str(date_str).split()[0], fmt)
        except ValueError:
            continue
    raise ValueError(f"Unable to parse date: {date_str}")

def generate_email(row: pd.Series) -> Tuple[str, str]:
    if row['country'] == 'Bol.com':
        return "", ""
    
    try:
        expected_date = parse_date(row['expected_date'])
        is_short_delay, week_number, timing = get_delivery_timing(expected_date)
        subject = f"Update bestelling {row['ordernr']}"
        
        if row['country'] == 'NL':
            body = generate_nl_short_delay(timing) if is_short_delay else generate_nl_long_delay(week_number)
        elif row['country'] == 'F':
            body = generate_fr_short_delay(timing) if is_short_delay else generate_fr_long_delay(week_number)
        elif row['country'] == 'D':
            body = generate_de_short_delay(timing) if is_short_delay else generate_de_long_delay(week_number)
        else:
            return "", ""
        
        # Add extra note if present and not empty
        if pd.notna(row['extra_note']) and str(row['extra_note']).strip():
            body = body.replace("Met vriendelijke groet,", 
                              f"Extra opmerking: {row['extra_note']}\n\nMet vriendelijke groet,")
            
        return subject, body
    except Exception as e:
        print(f"Error processing row: {e}")
        print(f"Row data: {row}")
        return "", ""

def process_orders(file_path: str):
    try:
        # Read all columns from Excel with date parsing
        df = pd.read_excel(file_path, parse_dates=[COLUMN_NAMES['expected_date'][0]])
        
        # Create column mapping dictionary
        column_mapping = {}
        missing_required = []
        
        for key, possible_names in COLUMN_NAMES.items():
            actual_name = get_actual_column_name(df, possible_names)
            if actual_name:
                column_mapping[key] = actual_name
            elif key != 'extra_note':  # extra_note is optional
                missing_required.append(possible_names[0])
        
        if missing_required:
            raise ValueError(f"Missing required columns: {', '.join(missing_required)}")
        
        # Add empty extra_note column if missing
        if 'extra_note' not in column_mapping:
            df['extra_note'] = ''
            column_mapping['extra_note'] = 'extra_note'
        
        for index, row in df.iterrows():
            if pd.notna(row[column_mapping['mailed']]):
                continue
                
            # Convert row to use our column mapping
            mapped_row = pd.Series({key: row[column_mapping[key]] for key in column_mapping})
            
            subject, body = generate_email(mapped_row)
            if body:
                mail = create_outlook_mail("", subject, body)
                mail.Display()
                
    except Exception as e:
        print(f"Error processing orders: {str(e)}")
        print(f"Available columns in file: {', '.join(df.columns)}")
        return

if __name__ == "__main__":
    process_orders("test.xlsx")
