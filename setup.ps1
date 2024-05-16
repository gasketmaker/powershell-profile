# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator"
    break
}

# Function to test internet connectivity
function Test-InternetConnection {
    try {
        $testConnection = Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Internet connection is required but not available. Please check your connection."
        return $false
    }
}

# Check for internet connectivity before proceeding
if (-not (Test-InternetConnection)) {
    break
}

# Profile creation or update
if (!(Test-Path -Path $PROFILE -PathType Leaf)) {
    try {
        # Detect Version of PowerShell & Create Profile directories if they do not exist.
        $profilePath = ""
        if ($PSVersionTable.PSEdition -eq "Core") { 
            $profilePath = "$env:userprofile\Documents\Powershell"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:userprofile\Documents\WindowsPowerShell"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        Invoke-RestMethod https://github.com/gasketmaker/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] has been created."
        Write-Host "If you want to add any persistent components, please do so at [$profilePath\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to create or update the profile. Error: $_"
    }
}
else {
    try {
        Get-Item -Path $PROFILE | Move-Item -Destination "oldprofile.ps1" -Force
        Invoke-RestMethod https://github.com/gasketmaker/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "The profile @ [$PROFILE] has been created and old profile removed."
        Write-Host "Please back up any persistent components of your old profile to [$HOME\Documents\PowerShell\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to backup and update the profile. Error: $_"
    }
}

# OMP Install
try {
    winget install -e --accept-source-agreements --accept-package-agreements JanDeDobbeleer.OhMyPosh
}
catch {
    Write-Error "Failed to install Oh My Posh. Error: $_"
}

function Install-Font {
    param (
        [string]$fontName,
        [string]$fontUrl
    )

    $localZip = "$env:TEMP\$fontName.zip"
    $extractionPath = "$env:TEMP\$fontName"

    # Download the font zip file
    Write-Host "Downloading $fontName..."
    Invoke-WebRequest -Uri $fontUrl -OutFile $localZip

    # Extract the font
    Write-Host "Installing $fontName..."
    Expand-Archive -Path $localZip -DestinationPath $extractionPath -Force

    # Copy font files to the Windows Fonts directory silently and register them
    $fontsDir = "C:\Windows\Fonts"
    Get-ChildItem -Path $extractionPath -Recurse -Filter "*.ttf" | ForEach-Object {
        $fontFile = $_
        $fontFilePath = Join-Path -Path $fontsDir -ChildPath $fontFile.Name

        if (-not (Test-Path $fontFilePath)) {
            Write-Host "Installing font: $fontFile.Name"
            $fontFile | Copy-Item -Destination $fontFilePath

            # Add font to registry
            $keyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            $fontNameValue = [System.IO.Path]::GetFileNameWithoutExtension($fontFile.Name) + " (TrueType)"
            Set-ItemProperty -Path $keyPath -Name $fontNameValue -Value $fontFile.Name
        } else {
            Write-Host "Font already installed: $fontFile.Name, skipping..."
        }
    }

    # Clean up
    Remove-Item -Path $extractionPath -Recurse -Force
    Remove-Item -Path $localZip -Force
}

# Prepare list of installed fonts
Add-Type -AssemblyName System.Drawing
$fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families
$installedFonts = [System.Collections.Generic.List[string]]::new()
foreach ($family in $fontFamilies) {
    $installedFonts.Add($family.Name)
}

# Font URLs and installation checks
$fontUrls = @{
    "FiraCode" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/FiraCode.zip"
    "JetBrainsMono" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/JetBrainsMono.zip"
    "SourceCodePro" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/SourceCodePro.zip"
    "RobotoMono" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/RobotoMono.zip"
    "Meslo" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/Meslo.zip"
    "CaskaydiaCove NF" = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/CascadiaCode.zip"
}

# Download and install each font if not already installed
foreach ($font in $fontUrls.GetEnumerator()) {
    Install-Font -fontName $font.Key -fontUrl $font.Value -installedFonts $installedFonts
}

# Load the System.Drawing assembly to access font families
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
$fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name

# List of all fonts to check
$requiredFonts = @("CaskaydiaCove NF", "FiraCode NF", "JetBrains Mono", "SourceCodePro NF", "RobotoMono NF", "Meslo NF")

# Check if all required fonts are installed
$allFontsInstalled = $requiredFonts | ForEach-Object { $fontFamilies -contains $_ } | ForEach-Object { $_ -eq $true }

# Final check and message to the user
if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and $allFontsInstalled) {
    Write-Host "Setup completed successfully. Please restart your PowerShell session to apply changes."
} else {
    Write-Warning "Setup completed with errors. Please check the error messages above."
}

# Choco install
try {
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}
catch {
    Write-Error "Failed to install Chocolatey. Error: $_"
}

# Terminal Icons Install
try {
    Install-Module -Name Terminal-Icons -Repository PSGallery -Force
}
catch {
    Write-Error "Failed to install Terminal Icons module. Error: $_"
}
# zoxide Install
try {
    winget install -e --id ajeetdsouza.zoxide
    Write-Host "zoxide installed successfully."
}
catch {
    Write-Error "Failed to install zoxide. Error: $_"
}
