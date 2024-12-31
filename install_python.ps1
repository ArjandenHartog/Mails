# Controleer of Python is geïnstalleerd
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "Python is niet geïnstalleerd. Download en installeer Python..."
    # Download Python installer (versie 3.x) van de officiële site
    $pythonInstaller = "https://www.python.org/ftp/python/3.10.9/python-3.10.9-amd64.exe"
    $installerPath = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest -Uri $pythonInstaller -OutFile $installerPath
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Write-Host "Python is geïnstalleerd."
    
    # Wacht even om zeker te zijn dat Python beschikbaar is in PATH
    Start-Sleep -Seconds 10
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# Controleer opnieuw of Python nu beschikbaar is
$python = Get-Command python -ErrorAction SilentlyContinue
if ($python) {
    Write-Host "Controleer of pip geïnstalleerd is..."
    try {
        $pipVersion = python -m pip --version
        Write-Host "pip versie: $pipVersion"
    } catch {
        Write-Host "pip is niet geïnstalleerd. Installeren..."
        python -m ensurepip --upgrade
    }

    Write-Host "Upgraden van pip naar laatste versie..."
    python -m pip install --upgrade pip

    Write-Host "Installeren van de benodigde packages..."
    # Array met alle benodigde packages
    $packages = @(
        'pandas',
        'pywin32',
        'openpyxl',  # Voor Excel ondersteuning
        'tkinter'    # Voor de GUI (meestal standaard aanwezig maar voor de zekerheid)
    )

    foreach ($pkg in $packages) {
        Write-Host "Installeren van $pkg..."
        try {
            python -m pip install --upgrade $pkg
            Write-Host "$pkg succesvol geïnstalleerd" -ForegroundColor Green
        } catch {
            Write-Host "Fout bij installeren van $pkg: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "`nAlle benodigde pakketten zijn geïnstalleerd." -ForegroundColor Green
    Write-Host "Het programma is klaar voor gebruik!"
} else {
    Write-Host "Er is een fout opgetreden tijdens het installeren van Python." -ForegroundColor Red
    Write-Host "Probeer Python handmatig te installeren vanaf python.org"
}

# Wacht op gebruiker input voordat het venster sluit
Write-Host "`nDruk op een toets om dit venster te sluiten..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
