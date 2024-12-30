import pandas as pd
import win32com.client as win32

def detect_outlook_version():
    """Detecteert of New Outlook of Classic Outlook wordt gebruikt"""
    try:
        outlook = win32.Dispatch('outlook.application')
        version = outlook.Version
        # New Outlook heeft versienummer dat begint met "16.0.15"
        return "new" if version.startswith("16.0.15") else "classic"
    except:
        return "classic"  # Fallback naar classic bij twijfel

# Functie om de tekst van de e-mail op te stellen
def genereer_email_tekst(row):
    klantnaam = row['Klant']
    ordernummer = row['Ordernummer']
    land = row['Land/site']
    leverdatum = row['Verwachte leverdatum']
    extra_info = row.get('Extra info', '')

    handtekening = ("Met vriendelijke groet,\n\n"
                    "LampenTotaal\n"
                    "Dorpsstraat 2a  |  4043 KK  Opheusden\n"
                    "T. 0488.750930  |  www.LampenTotaal.nl\n"
                    "IBAN NL65RABO0116060549")

    nl_be_base = (f"Geachte {klantnaam},\n\n"
                  f"Betreft: Status update bestelling {ordernummer}\n\n")

    if str(leverdatum).startswith("2049"):  # Onbekende levertijd
        body = (f"{nl_be_base}"
                f"Wij hebben een update over uw bestelling. Op dit moment is er helaas vertraging bij onze leverancier "
                f"waardoor we geen exacte leverdatum kunnen geven.\n\n"
                f"Wij staan in nauw contact met de leverancier en zodra wij meer informatie hebben over de leverdatum, "
                f"informeren wij u direct per e-mail.\n\n"
                f"Wij begrijpen dat dit voor u vervelend kan zijn. Mocht u naar aanleiding van deze informatie "
                f"vragen hebben of uw bestelling willen wijzigen, dan horen wij dat graag.\n\n"
                f"Wij danken u voor uw begrip.\n\n")

    elif str(leverdatum).startswith("2099"):  # Niet meer leverbaar
        alternatief_deel = f"\nSpeciaal voor u hebben wij het volgende alternatief geselecteerd:\n{extra_info}\n\n" if extra_info else "\n"
        body = (f"{nl_be_base}"
                f"Wij hebben helaas minder goed nieuws over uw bestelling. De fabrikant heeft ons geïnformeerd dat het "
                f"door u bestelde artikel niet meer geproduceerd wordt en daardoor definitief niet meer leverbaar is.{alternatief_deel}"
                f"Graag horen wij van u of:\n"
                f"- u interesse heeft in het voorgestelde alternatief\n"
                f"- u liever een ander artikel wilt uitzoeken\n"
                f"- u de bestelling wilt annuleren\n\n"
                f"U kunt reageren op deze e-mail of telefonisch contact met ons opnemen.\n\n"
                f"Wij bieden u onze welgemeende excuses aan voor het ongemak.\n\n")

    else:
        body = (f"{nl_be_base}"
                f"Uw bestelling wordt volgens planning geleverd op {leverdatum}.\n\n"
                f"Heeft u hierover nog vragen? Neem dan gerust contact met ons op.\n\n")

    if land.upper() in ["NL", "BE"]:
        return body + handtekening
    elif land.upper() == "FR":
        fr_base = (f"Cher/Chère {klantnaam},\n\n"
                  f"Objet : Mise à jour de votre commande {ordernummer}\n\n")

        if str(leverdatum).startswith("2049"):  # Onbekende levertijd FR
            body = (f"{fr_base}"
                    f"Nous vous contactons au sujet de votre commande. Actuellement, il y a un retard chez notre fournisseur "
                    f"et nous ne pouvons pas vous donner une date de livraison exacte.\n\n"
                    f"Nous sommes en contact étroit avec le fournisseur et dès que nous aurons plus d'informations sur la date "
                    f"de livraison, nous vous en informerons immédiatement par e-mail.\n\n"
                    f"Nous comprenons que cela puisse être gênant pour vous. Si vous avez des questions ou "
                    f"si vous souhaitez modifier votre commande, n'hésitez pas à nous contacter.\n\n"
                    f"Nous vous remercions de votre compréhension.\n\n")

        elif str(leverdatum).startswith("2099"):  # Niet meer leverbaar FR
            alternatief_fr = f"\nNous avons sélectionné spécialement pour vous l'alternative suivante :\n{extra_info}\n\n" if extra_info else "\n"
            body = (f"{fr_base}"
                    f"Nous avons malheureusement une mauvaise nouvelle concernant votre commande. Le fabricant nous a informés "
                    f"que l'article que vous avez commandé n'est plus en production et ne sera donc plus disponible.{alternatief_fr}"
                    f"Nous aimerions savoir si :\n"
                    f"- vous êtes intéressé(e) par l'alternative proposée\n"
                    f"- vous préférez choisir un autre article\n"
                    f"- vous souhaitez annuler la commande\n\n"
                    f"Vous pouvez répondre à cet e-mail ou nous contacter par téléphone.\n\n"
                    f"Nous vous prions de nous excuser pour ce désagrément.\n\n")
        else:
            body = (f"{fr_base}"
                    f"Votre commande sera livrée selon le planning le {leverdatum}.\n\n"
                    f"Si vous avez des questions, n'hésitez pas à nous contacter.\n\n")

        return body + handtekening

    elif land.upper() == "DE":
        de_base = (f"Sehr geehrte(r) {klantnaam},\n\n"
                  f"Betreff: Update zu Ihrer Bestellung {ordernummer}\n\n")

        if str(leverdatum).startswith("2049"):  # Onbekende levertijd DE
            body = (f"{de_base}"
                    f"Wir möchten Sie über den Status Ihrer Bestellung informieren. Derzeit gibt es leider Verzögerungen "
                    f"bei unserem Lieferanten, wodurch wir kein genaues Lieferdatum nennen können.\n\n"
                    f"Wir stehen in engem Kontakt mit dem Lieferanten und werden Sie umgehend per E-Mail informieren, "
                    f"sobald wir weitere Informationen zum Lieferdatum haben.\n\n"
                    f"Wir verstehen, dass dies für Sie unangenehm sein kann. Sollten Sie aufgrund dieser Information "
                    f"Fragen haben oder Ihre Bestellung ändern möchten, lassen Sie es uns bitte wissen.\n\n"
                    f"Wir danken Ihnen für Ihr Verständnis.\n\n")

        elif str(leverdatum).startswith("2099"):  # Niet meer leverbaar DE
            alternatief_de = f"\nSpeziell für Sie haben wir folgende Alternative ausgewählt:\n{extra_info}\n\n" if extra_info else "\n"
            body = (f"{de_base}"
                    f"Leider haben wir keine guten Nachrichten zu Ihrer Bestellung. Der Hersteller hat uns informiert, "
                    f"dass der von Ihnen bestellte Artikel nicht mehr produziert wird und daher endgültig nicht mehr "
                    f"lieferbar ist.{alternatief_de}"
                    f"Bitte teilen Sie uns mit, ob:\n"
                    f"- Sie Interesse an der vorgeschlagenen Alternative haben\n"
                    f"- Sie lieber einen anderen Artikel aussuchen möchten\n"
                    f"- Sie die Bestellung stornieren möchten\n\n"
                    f"Sie können auf diese E-Mail antworten oder uns telefonisch kontaktieren.\n\n"
                    f"Wir entschuldigen uns aufrichtig für die Unannehmlichkeiten.\n\n")
        else:
            body = (f"{de_base}"
                    f"Ihre Bestellung wird planmäßig am {leverdatum} geliefert.\n\n"
                    f"Haben Sie dazu noch Fragen? Kontaktieren Sie uns gerne.\n\n")

        return body + handtekening

    else:
        return None

# Functie om een e-mail te maken in Outlook
def maak_email(row):
    outlook = win32.Dispatch('outlook.application')
    outlook_version = detect_outlook_version()
    
    if outlook_version == "new":
        # New Outlook methode
        namespace = outlook.GetNamespace("MAPI")
        mail = namespace.CreateItem(0)
    else:
        # Classic Outlook methode
        mail = outlook.CreateItem(0)

    mail.To = row['gemaild'] if row['gemaild'] else input(f"Voer e-mailadres in voor {row['Klant']}: ")
    mail.Subject = f"Update over uw bestelling: {row['Ordernummer']}"
    mail.Body = genereer_email_tekst(row)

    if mail.Body:
        try:
            if outlook_version == "new":
                mail.Display(False)  # False om pop-up waarschuwingen te voorkomen
            else:
                mail.Display()
        except:
            print(f"Waarschuwing: Kon e-mail niet weergeven voor {row['Klant']}")

# Inlezen van het Excel-bestand
bestandspad = "test.xlsx"  # Changed from "/test.xlsx" to use relative path
# or use absolute path with proper Windows formatting:
# bestandspad = r"C:\Users\arjan_h\Documents\GitHub\Mails\test.xlsx"
kolomnamen = ['Ordernummer', 'Orderdatum', 'Klant', 'gemaild', 'Leverdatum', 'Verwachte leverdatum', 'Land/site', 'Extra info']
data = pd.read_excel(bestandspad, usecols=kolomnamen)

# Itereren over de rijen in het Excel-bestand
for index, row in data.iterrows():
    print(f"Maak e-mail voor {row['Klant']}...")
    maak_email(row)
