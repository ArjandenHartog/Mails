import os
import sys
import subprocess
import urllib.request
import time

def print_status(message, success=True):
    """Print een statusbericht met emoji"""
    emoji = "✅" if success else "❌"
    print(f"{emoji} {message}")

def install_requirements():
    """Installeer benodigde packages"""
    try:
        packages = ['pandas', 'pywin32', 'requests']
        for package in packages:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        return True
    except Exception as e:
        print_status(f"Fout bij installeren packages: {str(e)}", False)
        return False

def download_files():
    """Download bestanden van GitHub"""
    base_url = "https://raw.githubusercontent.com/ArjandenHartog/Mails/main/"
    files = {
        'lampentotaal_mail.py': 'Main application file',
        'requirements.txt': 'Requirements file'
    }
    
    try:
        # Maak Backorders directory in Documents
        docs_path = os.path.expanduser("~/Documents/Backorders")
        if not os.path.exists(docs_path):
            os.makedirs(docs_path)
        
        # Download elk bestand
        for filename, description in files.items():
            print(f"Downloading {description}...")
            url = base_url + filename
            dest_path = os.path.join(docs_path, filename)
            urllib.request.urlretrieve(url, dest_path)
            print_status(f"Gedownload: {filename}")
            
        return True
    except Exception as e:
        print_status(f"Fout bij downloaden: {str(e)}", False)
        return False

def main():
    print("LampenTotaal Mail Verwerker - Installer")
    print("======================================")
    
    if not install_requirements():
        input("\nFout bij installatie. Druk op Enter om af te sluiten...")
        return
    
    if not download_files():
        input("\nFout bij downloaden. Druk op Enter om af te sluiten...")
        return
    
    print("\n✨ Installatie succesvol afgerond!")
    print("\nHet programma wordt nu gestart...")
    time.sleep(2)
    
    # Start het hoofdprogramma
    try:
        docs_path = os.path.expanduser("~/Documents/Backorders")
        main_script = os.path.join(docs_path, "lampentotaal_mail.py")
        subprocess.Popen([sys.executable, main_script])
    except Exception as e:
        print_status(f"Fout bij starten programma: {str(e)}", False)
        input("\nDruk op Enter om af te sluiten...")

if __name__ == "__main__":
    main()
