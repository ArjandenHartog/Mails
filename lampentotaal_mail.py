import sys
import subprocess
import importlib.util
import os

# Add PyInstaller specific code for determining base path
def get_base_path():
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return base_path

def check_package(package_name):
    """Check if a package is installed and importable"""
    try:
        # Try to actually import the package instead of just checking if it exists
        __import__(package_name.split('.')[0])
        return True
    except ImportError:
        return False

def install_packages():
    """Install required packages only if they're not already installed"""
    required_packages = ['pandas', 'win32com']  # Changed from pywin32 to win32com
    
    try:
        need_install = False
        for package in required_packages:
            if not check_package(package):
                print(f"Package {package} needs to be installed")
                need_install = True
                if package == 'win32com':
                    package = 'pywin32'  # Install pywin32 when win32com is needed
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        
        if need_install:
            print("Required packages installed. Please restart the application.")
            sys.exit(0)
            
    except Exception as e:
        print(f"Error installing packages: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    install_packages()

# Standaard imports
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import pandas as pd
import win32com.client
import os

def detect_outlook_version():
    """Detecteert welke Outlook versie wordt gebruikt"""
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        # Als we hier komen is Outlook beschikbaar
        return "classic"  # Standaard classic gebruiken voor betere compatibiliteit
    except:
        return None

def check_outlook():
    """Controleert of Outlook draait en toegankelijk is"""
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        outlook.GetNamespace("MAPI")  # Test de verbinding
        return True
    except:
        return False

def get_delivery_timing(exp_date):
    """Bepaal de juiste leveringstiming op basis van de verwachte datum"""
    today = datetime.now()
    this_week = today.isocalendar()[1]
    delivery_week = exp_date.isocalendar()[1]
    days_until = (exp_date - today).days
    
    if days_until < 0:
        return "deze week"  # Als datum in verleden ligt
    elif days_until <= 2:
        return "deze week"
    elif days_until <= 5:
        return "eind deze week"
    elif days_until <= 7:
        return "begin volgende week"
    elif days_until <= 10:
        return "in de loop van volgende week"
    elif days_until <= 14:
        return "eind volgende week"
    else:
        return f"in week {delivery_week}"

def parse_date(date_str):
    # ...existing code from nietwebshop.py...
    pass

def get_niet_webshop_template(country, is_short_delay, week_nr, timing):
    # Merged short/long delay logic from nietwebshop.py
    # Support for NL/DE/FR with short/long delay
    # ...existing code...
    pass

def get_nml_template(row):
    # Merged 2049 (unknown) / 2099 (not available) logic from NML.py
    # ...existing code...
    pass

def create_outlook_mail(to_address, subject, body):
    # ...existing code from nietwebshop.py...
    pass

def get_greeting(name, language='NL'):
    """Get time-based greeting without name in specified language"""
    hour = datetime.now().hour
    
    if language.upper() in ['NL', 'BE']:
        if 5 <= hour < 12:
            greeting = "Goedemorgen"
        elif 12 <= hour < 17:
            greeting = "Goedemiddag"
        else:
            greeting = "Goedenavond"
    elif language.upper() in ['F', 'FR']:
        if 5 <= hour < 12:
            greeting = "Bonjour"
        elif 12 <= hour < 17:
            greeting = "Bon après-midi"
        else:
            greeting = "Bonsoir"
    elif language.upper() in ['D', 'DE']:
        if 5 <= hour < 12:
            greeting = "Guten Morgen"
        elif 12 <= hour < 17:
            greeting = "Guten Tag"
        else:
            greeting = "Guten Abend"
            
    return greeting  # Return without comma

class MailApp:
    def __init__(self):
        # Get base path for resources
        self.base_path = get_base_path()
        
        # Window setup
        self.window = tk.Tk()
        self.window.title("LampenTotaal Mail Verwerker")
        self.window.geometry("800x600")  # Larger window
        self.window.configure(bg='#f5f6fa')
        
        # Center window
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width/2) - (800/2)
        y = (screen_height/2) - (600/2)
        self.window.geometry(f'+{int(x)}+{int(y)}')
        
        # Enhanced style configuration
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main button style
        style.configure(
            'Main.TButton', 
            padding=15,
            font=('Helvetica', 12, 'bold'),
            background='#4a90e2',
            foreground='white'
        )
        
        # Template button style
        style.configure(
            'Template.TButton',
            padding=10,
            font=('Helvetica', 10),
            background='#2ecc71'
        )
        
        self.signature = """<br><br>Met vriendelijke groet,<br><br>
LampenTotaal<br>
Dorpsstraat 2a  |  4043 KK  Opheusden  |  T. 0488.750930<br>
IBAN NL65RABO0116060549  |  www.LampenTotaal.nl"""
        
        self.create_widgets()

    def verify_outlook_connection(self):
        """Verify Outlook connection before processing"""
        if not check_outlook():
            messagebox.showerror(
                "Outlook niet gevonden",
                "Start Microsoft Outlook en probeer het opnieuw.\n\n"
                "Als Outlook al draait, herstart deze dan."
            )
            return False
        return True

    def create_widgets(self):
        # Main container with padding
        container = tk.Frame(self.window, bg='#f5f6fa', padx=40, pady=30)
        container.pack(fill='both', expand=True)
        
        # Header section
        header = tk.Frame(container, bg='#f5f6fa')
        header.pack(fill='x', pady=(0, 30))
        
        title = tk.Label(
            header,
            text="LampenTotaal Mail Verwerker",
            font=('Helvetica', 24, 'bold'),
            bg='#f5f6fa',
            fg='#2c3e50'
        )
        title.pack()
        
        subtitle = tk.Label(
            header,
            text="Selecteer een optie om mail templates te genereren",
            font=('Helvetica', 12),
            bg='#f5f6fa',
            fg='#7f8c8d'
        )
        subtitle.pack(pady=(5, 0))
        
        # Main buttons section
        main_buttons = tk.Frame(container, bg='#f5f6fa')
        main_buttons.pack(fill='x', pady=20)
        
        # Create two columns for main buttons
        left_col = tk.Frame(main_buttons, bg='#f5f6fa')
        left_col.pack(side='left', expand=True, padx=10)
        
        right_col = tk.Frame(main_buttons, bg='#f5f6fa')
        right_col.pack(side='right', expand=True, padx=10)
        
        # Main action buttons with icons and descriptions
        self.create_action_button(
            left_col,
            "📦 Niet Webshop Orders",
            "Verwerk orders van niet-webshop bestellingen",
            lambda: self.process_file('niet_webshop')
        )
        
        self.create_action_button(
            right_col,
            "🚚 NML + NNB Orders",  # Aangepaste tekst
            "Verwerk orders van NML en NNB",  # Duidelijkere beschrijving
            lambda: self.process_file('nml')
        )
        
        # Template section
        template_section = tk.Frame(container, bg='#f5f6fa')
        template_section.pack(fill='x', pady=30)
        
        template_header = tk.Label(
            template_section,
            text="Excel Templates",
            font=('Helvetica', 14, 'bold'),
            bg='#f5f6fa',
            fg='#2c3e50'
        )
        template_header.pack(pady=(0, 15))
        
        # Template buttons in a row
        template_buttons = tk.Frame(template_section, bg='#f5f6fa')
        template_buttons.pack()
        
        for template_type in ['niet_webshop', 'nml']:
            btn = ttk.Button(
                template_buttons,
                text=f"📄 {template_type.title()} Template",
                style='Template.TButton',
                command=lambda t=template_type: self.create_template(t)
            )
            btn.pack(side='left', padx=10)
        
        # Status section
        self.status = tk.Label(
            container,
            text="✅ Gereed om orders te verwerken",
            font=('Helvetica', 11),
            bg='#f5f6fa',
            fg='#27ae60'
        )
        self.status.pack(pady=20)

    def create_action_button(self, parent, text, description, command):
        frame = tk.Frame(parent, bg='#f5f6fa')
        frame.pack(pady=10, padx=20)
        
        btn = ttk.Button(
            frame,
            text=text,
            style='Main.TButton',
            command=command
        )
        btn.pack(fill='x', ipady=15)
        
        desc = tk.Label(
            frame,
            text=description,
            font=('Helvetica', 10),
            bg='#f5f6fa',
            fg='#7f8c8d',
            wraplength=250
        )
        desc.pack(pady=(5, 0))

    def process_file(self, order_type):
        if not self.verify_outlook_connection():
            return
            
        file_path = filedialog.askopenfilename(
            title=f"Selecteer Excel bestand voor {order_type}",
            filetypes=[("Excel bestanden", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
            
        try:
            self.status.config(text="Bezig met verwerken...", fg='#e67e22')
            self.window.update()
            
            if order_type == 'niet_webshop':
                self.process_niet_webshop(file_path)
            else:
                self.process_nml(file_path)
            
            self.status.config(
                text=f"✅ Succesvol verwerkt: {os.path.basename(file_path)}", 
                fg='#27ae60'
            )
            
        except Exception as e:
            self.status.config(text="❌ Fout bij verwerken", fg='#c0392b')
            messagebox.showerror("Fout", str(e))

    def process_niet_webshop(self, file_path):
        df = pd.read_excel(file_path)
        required_columns = ['Ordernummer', 'Klant', 'Verwachte leverdatum', 'Land', 'Gemaild']
        
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Kolommen niet gevonden: {', '.join(missing)}")

        processed = 0
        skipped_mailed = 0
        skipped_bol = 0
        mails = []  # Store mail objects to display later

        for _, row in df.iterrows():
            if pd.notna(row.get('Gemaild')) and str(row.get('Gemaild')).strip():
                skipped_mailed += 1
                continue
                
            if str(row['Land']) == 'Bol.com':
                skipped_bol += 1
                continue
                
            mail = self.create_mail(row, display=False)
            if mail:
                mails.append(mail)
                processed += 1

        # Display all mails at once
        for mail in mails:
            mail.Display()

        # Update status with detailed counts
        status_text = f"✅ Verwerkt: {processed} orders"
        if skipped_mailed > 0:
            status_text += f"\n(Overgeslagen: {skipped_mailed} reeds gemailden)"
        if skipped_bol > 0:
            status_text += f"\n(Overgeslagen: {skipped_bol} Bol.com orders)"
        
        self.status.config(text=status_text, fg='#27ae60')

    def process_nml(self, file_path):
        df = pd.read_excel(file_path)
        required_columns = ['Ordernummer', 'Klant', 'Verwachte leverdatum', 'Land/site', 'Extra info']
        
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Kolommen niet gevonden: {', '.join(missing)}")

        processed = 0
        skipped_bol = 0
        skipped_mailed = 0
        mails = []  # Store mail objects to display later

        for _, row in df.iterrows():
            # Skip if already mailed
            if pd.notna(row.get('gemaild')) and str(row.get('gemaild')).strip():
                skipped_mailed += 1
                continue
                
            # Skip Bol.com orders
            if str(row['Land/site']) == 'Bol.com':
                skipped_bol += 1
                continue
                
            # Create mail but don't display yet
            mail = self.create_nml_mail(row, display=False)
            if mail:
                mails.append(mail)
                processed += 1
        
        # Display all mails at once
        for mail in mails:
            mail.Display()
        
        # Update status with detailed information
        status_text = f"✅ Verwerkt: {processed} orders"
        if skipped_bol > 0:
            status_text += f"\n(Overgeslagen: {skipped_bol} Bol.com orders)"
        if skipped_mailed > 0:
            status_text += f"\n(Overgeslagen: {skipped_mailed} reeds gemailden)"
        
        self.status.config(text=status_text, fg='#27ae60')

    def create_mail(self, row, display=True):
        if str(row['Land']) == 'Bol.com':
            return None
        
        try:
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # Simpelweg een nieuwe mail maken
            
            exp_date = pd.to_datetime(row['Verwachte leverdatum'])
            is_short_delay, week_nr = self.get_delivery_info(exp_date)
            delivery_timing = get_delivery_timing(exp_date)
            
            mail.Subject = f"Update bestelling {row['Ordernummer']}"
            
            # Get template first
            body = self.get_mail_template(str(row['Land']).upper(), is_short_delay, week_nr, delivery_timing)
            
            # Get greeting without customer name and replace it in the template
            greeting = get_greeting(str(row['Land']).upper())
            body = body.replace("Goedemiddag,", f"{greeting} {row['Klant']},")
            body = body.replace("Guten Tag,", f"{greeting} {row['Klant']},")
            body = body.replace("Bonjour,", f"{greeting} {row['Klant']},")
            
            if pd.notna(row.get('Extra opmerking')):
                body = f"{body}\n\nExtra opmerking: {row['Extra opmerking']}"
            
            # Convert to HTML and add signature
            mail.HTMLBody = body.replace('\n', '<br>') + self.signature
            
            if display:
                mail.Display()
            
            return mail
        except Exception as e:
            messagebox.showerror("Outlook Fout", f"Fout bij maken mail: {str(e)}")
            return None

    def create_nml_mail(self, row, display=True):
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)  # Simpelweg een nieuwe mail maken

            mail.Subject = f"Update bestelling {row['Ordernummer']}"
            
            # Get template
            body = self.get_nml_template(row)
            
            # Get greeting without customer name
            greeting = get_greeting(str(row['Land/site']).upper())
            
            # Replace all possible greetings with time-based greeting + name (only once)
            replacements = {
                f"Geachte {row['Klant']}": f"{greeting} {row['Klant']},",
                f"Cher/Chère {row['Klant']}": f"{greeting} {row['Klant']},",
                f"Sehr geehrte(r) {row['Klant']}": f"{greeting} {row['Klant']},"
            }
            
            for old, new in replacements.items():
                body = body.replace(old, new)
            
            # Convert to HTML and add signature
            mail.HTMLBody = body.replace('\n', '<br>') + self.signature
            
            if display:
                mail.Display()
            
            return mail
        except Exception as e:
            messagebox.showerror("Outlook Fout", f"Fout bij maken mail: {str(e)}")
            return None

    def get_delivery_info(self, exp_date):
        today = datetime.now()
        diff = (exp_date - today).days
        week_nr = exp_date.isocalendar()[1]
        return diff <= 7, week_nr

    def get_mail_template(self, country, is_short_delay, week_nr, timing):
        if country in ['NL', 'BE']:  # Handle both NL and BE
            if is_short_delay:
                return f"""Goedemiddag,

Hartelijk dank voor uw bestelling.

Hierbij meer informatie met betrekking tot de uitlevering van de bestelling die u bij ons hebt geplaatst.

Wij verwachten {timing} uw pakket te gaan ontvangen van onze leverancier, uiteraard gaan wij ons best doen om uw pakket verder direct te gaan versturen naar uw postadres.
Zodra wij uw pakket hebben verzonden ontvangt u een track en trace code per mail. Hiermee kunt u het pakket volgen."""
            else:
                return f"""Goedemiddag,

Hartelijk dank voor uw bestelling.

Hierbij meer informatie met betrekking tot de uitlevering van de bestelling die u bij ons hebt geplaatst.
Helaas hebben wij van de leverancier vernomen dat het artikel wat u besteld heeft momenteel niet op voorraad is. Wij verwachten uw bestelling in week {week_nr} binnen te krijgen.

Wij hopen u voldoende te hebben geïnformeerd. Mocht u nog vragen hebben, mail of bel gerust.

Excuses voor het ongemak!"""

        elif country in ['D', 'DE']:  # Handle both D and DE
            timing_de = {
                "deze week": "diese Woche",
                "eind deze week": "Ende dieser Woche",
                "begin volgende week": "Anfang nächster Woche",
                "in de loop van volgende week": "im Laufe der nächsten Woche",
                "eind volgende week": "Ende nächster Woche"
            }.get(timing, f"in Kalenderwoche {week_nr}")
            
            if is_short_delay:
                return f"""Guten Tag,

Vielen Dank für Ihre Bestellung.

Hiermit möchten wir Sie über den Status Ihrer Bestellung informieren.

Wir erwarten, dass wir Ihr Paket {timing_de} von unserem Lieferanten erhalten werden. Selbstverständlich werden wir uns bemühen, Ihr Paket dann umgehend an Ihre Postadresse zu versenden.
Sobald wir Ihr Paket versendet haben, erhalten Sie eine Track & Trace Nummer per E-Mail.

Wir hoffen, Sie ausreichend informiert zu haben. Bei Fragen können Sie uns gerne kontaktieren."""
            else:
                return f"""Guten Tag,

Vielen Dank für Ihre Bestellung.

Hiermit möchten wir Sie über den Status Ihrer Bestellung informieren.
Leider müssen wir Ihnen mitteilen, dass der von Ihnen bestellte Artikel derzeit nicht vorrätig ist. Wir erwarten die Lieferung in Kalenderwoche {week_nr}.

Wir hoffen, Sie ausreichend informiert zu haben. Bei Fragen können Sie uns gerne kontaktieren.

Wir entschuldigen uns für die Unannehmlichkeiten.

Falls sich an der oben genannten Lieferzeit etwas ändern sollte, werden wir Sie selbstverständlich umgehend informieren."""

        elif country in ['F', 'FR']:  # Handle both F and FR
            timing_fr = {
                "deze week": "cette semaine",
                "eind deze week": "en fin de semaine",
                "begin volgende week": "en début de semaine prochaine",
                "in de loop van volgende week": "au cours de la semaine prochaine",
                "eind volgende week": "en fin de semaine prochaine"
            }.get(timing, f"dans la semaine {week_nr}")
            
            if is_short_delay:
                return f"""Bonjour,

Nous vous remercions de votre commande.

Voici plus d'informations concernant la livraison de votre commande.

Nous prévoyons de recevoir votre colis {timing_fr} de notre fournisseur. Bien entendu, nous ferons de notre mieux pour expédier votre colis directement à votre adresse postale.
Dès que nous aurons expédié votre colis, vous recevrez un code de suivi par e-mail vous permettant de suivre le colis.

Nous espérons vous avoir suffisamment informé. Si vous avez des questions, n'hésitez pas à nous contacter."""
            else:
                return f"""Bonjour,

Nous vous remercions de votre commande.

Voici plus d'informations concernant la livraison de votre commande.
Malheureusement, nous devons vous informer que l'article que vous avez commandé n'est actuellement pas en stock. Nous prévoyons de recevoir votre commande dans la semaine {week_nr}.

Nous espérons vous avoir suffisamment informé. Si vous avez des questions, n'hésitez pas à nous contacter.

Nous nous excusons pour les désagréments.

Si le délai de livraison mentionné ci-dessus devait changer, nous vous en informerons immédiatement."""

    def get_nml_template(self, row):
        klantnaam = row['Klant']
        ordernummer = row['Ordernummer']
        land = row['Land/site']
        leverdatum = row['Verwachte leverdatum']
        extra_info = row.get('Extra info', '')

        nl_be_base = (f"Geachte {klantnaam},\n\n"
                     f"Betreft: Status update bestelling {ordernummer}\n\n")

        fr_base = (f"Cher/Chère {klantnaam},\n\n"
                  f"Objet : Mise à jour de votre commande {ordernummer}\n\n")

        de_base = (f"Sehr geehrte(r) {klantnaam},\n\n"
                  f"Betreff: Update zu Ihrer Bestellung {ordernummer}\n\n")

        # Check special dates and country
        # Convert timestamp to string for comparison
        if str(leverdatum).startswith("2049"):
            if str(land).upper() in ["NL", "BE"]:
                body = (f"{nl_be_base}"
                       f"Wij hebben een update over uw bestelling. Op dit moment is er helaas vertraging bij onze leverancier "
                       f"waardoor we geen exacte leverdatum kunnen geven.\n\n"
                       f"Wij staan in nauw contact met de leverancier en zodra wij meer informatie hebben over de leverdatum, "
                       f"informeren wij u direct per e-mail.\n\n"
                       f"Wij begrijpen dat dit voor u vervelend kan zijn. Mocht u naar aanleiding van deze informatie "
                       f"vragen hebben of uw bestelling willen wijzigen, dan horen wij dat graag.\n\n"
                       f"Wij danken u voor uw begrip.\n\n")
            elif str(land).upper() == "FR":
                body = (f"{fr_base}"
                       f"Nous vous contactons au sujet de votre commande. Actuellement, il y a un retard chez notre fournisseur "
                       f"et nous ne pouvons pas vous donner une date de livraison exacte.\n\n"
                       f"Nous sommes en contact étroit avec le fournisseur et dès que nous aurons plus d'informations sur la date "
                       f"de livraison, nous vous en informerons immédiatement par e-mail.\n\n"
                       f"Nous comprenons que cela puisse être gênant pour vous. Si vous avez des questions ou "
                       f"si vous souhaitez modifier votre commande, n'hésitez pas à nous contacter.\n\n"
                       f"Nous vous remercions de votre compréhension.\n\n")
            elif str(land).upper() == "DE":
                body = (f"{de_base}"
                       f"Wir möchten Sie über den Status Ihrer Bestellung informieren. Derzeit gibt es leider Verzögerungen "
                       f"bei unserem Lieferanten, wodurch wir kein genaues Lieferdatum nennen können.\n\n"
                       f"Wir stehen in engem Kontakt mit dem Lieferanten und werden Sie umgehend per E-Mail informieren, "
                       f"sobald wir weitere Informationen zum Lieferdatum haben.\n\n"
                       f"Wir verstehen, dass dies für Sie unangenehm sein kann. Sollten Sie aufgrund dieser Information "
                       f"Fragen haben oder Ihre Bestellung ändern möchten, lassen Sie es uns bitte wissen.\n\n"
                       f"Wir danken Ihnen für Ihr Verständnis.\n\n")

        elif str(leverdatum).startswith("2099"):
            alternatief_deel = f"\nSpeciaal voor u hebben wij het volgende alternatief geselecteerd:\n{extra_info}\n\n" if extra_info else "\n"
            alternatief_fr = f"\nNous avons sélectionné spécialement pour vous l'alternative suivante :\n{extra_info}\n\n" if extra_info else "\n"
            alternatief_de = f"\nSpeziell für Sie haben wir folgende Alternative ausgewählt:\n{extra_info}\n\n" if extra_info else "\n"

            if str(land).upper() in ["NL", "BE"]:
                body = (f"{nl_be_base}"
                       f"Wij hebben helaas minder goed nieuws over uw bestelling. De fabrikant heeft ons geïnformeerd dat het "
                       f"door u bestelde artikel niet meer geproduceerd wordt en daardoor definitief niet meer leverbaar is.{alternatief_deel}"
                       f"Graag horen wij van u of:\n"
                       f"- u interesse heeft in het voorgestelde alternatief\n"
                       f"- u liever een ander artikel wilt uitzoeken\n"
                       f"- u de bestelling wilt annuleren\n\n"
                       f"U kunt reageren op deze e-mail of telefonisch contact met ons opnemen.\n\n"
                       f"Wij bieden u onze welgemeende excuses aan voor het ongemak.\n\n")
            elif str(land).upper() == "FR":
                body = (f"{fr_base}"
                       f"Nous avons malheureusement de mauvaises nouvelles concernant votre commande. Le fabricant nous a informés que "
                       f"l'article que vous avez commandé n'est plus produit et ne sera donc plus disponible.{alternatief_fr}"
                       f"Nous aimerions savoir si :\n"
                       f"- vous êtes intéressé par l'alternative proposée\n"
                       f"- vous préférez choisir un autre article\n"
                       f"- vous souhaitez annuler la commande\n\n"
                       f"Vous pouvez répondre à cet e-mail ou nous contacter par téléphone.\n\n"
                       f"Nous vous présentons nos excuses pour le désagrément.\n\n")
            elif str(land).upper() == "DE":
                body = (f"{de_base}"
                       f"Wir haben leider schlechte Nachrichten zu Ihrer Bestellung. Der Hersteller hat uns informiert, dass der von Ihnen "
                       f"bestellte Artikel nicht mehr produziert wird und daher endgültig nicht mehr lieferbar ist.{alternatief_de}"
                       f"Wir möchten von Ihnen wissen, ob:\n"
                       f"- Sie an der vorgeschlagenen Alternative interessiert sind\n"
                       f"- Sie lieber einen anderen Artikel auswählen möchten\n"
                       f"- Sie die Bestellung stornieren möchten\n\n"
                       f"Sie können auf diese E-Mail antworten oder uns telefonisch kontaktieren.\n\n"
                       f"Wir entschuldigen uns aufrichtig für die Unannehmlichkeiten.\n\n")

        else:
            # Normal delivery templates
            return self.get_mail_template(land, False, leverdatum.isocalendar()[1])

        return body

    def create_template(self, template_type):
        try:
            # Gebruik AppData/Local voor templates in plaats van Documents
            default_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'LampenTotaal', 'Backorders')
            
            # Maak directory aan als deze niet bestaat
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
            
            # Bepaal de standaard bestandsnaam
            default_filename = f"template_{template_type}.xlsx"
            default_path = os.path.join(default_dir, default_filename)
            
            # Open bestandsdialoog met standaard directory
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename,
                initialdir=default_dir,
                title="Sla template op als"
            )
            
            if not file_path:  # Als gebruiker annuleert
                return
            
            # Controleer of het pad geldig is
            directory = os.path.dirname(file_path)
            if not os.path.exists(directory):
                os.makedirs(directory)
                
            # Maak DataFrame met juiste kolommen zoals eerder
            if template_type == 'niet_webshop':
                df = pd.DataFrame(columns=[
                    'Ordernummer', 'Orderdatum', 'Klant', 'Gemaild', 
                    'Leverdatum', 'Verwachte leverdatum', 'Land', 'extra opmerking'
                ])
                example_data = {
                    'Ordernummer': ['12345'],
                    'Orderdatum': ['2024-01-01'],
                    'Klant': ['Voorbeeld Klant'],
                    'Gemaild': [''],
                    'Leverdatum': [''],
                    'Verwachte leverdatum': ['2024-02-01'],
                    'Land': ['NL'],
                    'extra opmerking': ['Optionele opmerking']
                }
            else:  # NML template
                df = pd.DataFrame(columns=[
                    'Ordernummer', 'Orderdatum', 'Klant', 'gemaild',
                    'Leverdatum', 'Verwachte leverdatum', 'Land/site', 'Extra info'
                ])
                example_data = {
                    'Ordernummer': ['12345'],
                    'Orderdatum': ['2024-01-01'],
                    'Klant': ['Voorbeeld Klant'],
                    'gemaild': [''],
                    'Leverdatum': [''],
                    'Verwachte leverdatum': ['2049-01-01'],
                    'Land/site': ['NL'],
                    'Extra info': ['Alternatief product suggestie']
                }

            # Voeg voorbeeldrij toe
            example_df = pd.DataFrame(example_data)
            df = pd.concat([df, example_df], ignore_index=True)

            # Sla het bestand op met extra controles
            try:
                df.to_excel(file_path, index=False)
                
                # Controleer of het bestand succesvol is aangemaakt
                if os.path.exists(file_path):
                    os.startfile(file_path)
                    messagebox.showinfo(
                        "Template aangemaakt",
                        f"Template is opgeslagen en geopend:\n{file_path}"
                    )
                else:
                    raise FileNotFoundError(f"Kon het bestand niet aanmaken op locatie: {file_path}")
                    
            except PermissionError:
                messagebox.showerror(
                    "Fout",
                    "Kon het bestand niet opslaan. Mogelijk is het geopend in een ander programma."
                )
            except Exception as e:
                messagebox.showerror(
                    "Fout",
                    f"Fout bij opslaan van het bestand:\n{str(e)}"
                )
                
        except Exception as e:
            messagebox.showerror(
                "Fout",
                f"Onverwachte fout bij aanmaken template:\n{str(e)}\n\nLocatie: {default_dir}"
            )

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = MailApp()
    app.run()
