<#
.SYNOPSIS
    OPTI-SOURCE - SCRIPT DE MAINTENANCE CENTRALISÉ
.DESCRIPTION
    Script regroupant les modes de maintenance pour l'école de la Source.
#>

# ==============================
# VÉRIFICATION ADMIN GLOBALE
# ==============================
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "ERREUR : Ce script doit être exécuté avec les droits Administrateur."
    Read-Host "Appuyez sur Entrée pour fermer le script..."
    exit
}

# ==============================
# FONCTIONS GLOBALES (COMMUNES)
# ==============================
function Show-Status($message, $success) {
    if ($success) { Write-Host "$message - OK" -ForegroundColor Green }
    else { Write-Host "$message - ERREUR / IGNORÉ" -ForegroundColor Red }
}

function Get-TargetProfilePath {
    <# 
    Recherche avancée du profil utilisateur. 
    Priorité 1 : Le compte "cours" (même avec un suffixe de domaine comme cours.ecole)
    Priorité 2 : L'utilisateur qui a ouvert la session Windows (pas l'admin caché)
    #>
    $coursProfile = Get-ChildItem -Path "C:\Users" -Filter "cours*" -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($coursProfile) { return $coursProfile.FullName }
    
    $interactiveUser = (Get-CimInstance Win32_ComputerSystem).UserName
    if ($interactiveUser) { 
        $uName = $interactiveUser.Split('\')[-1]
        $activeProfile = Get-ChildItem -Path "C:\Users" -Filter "$uName*" -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($activeProfile) { return $activeProfile.FullName }
    }
    
    return $env:USERPROFILE
}

function Get-DesktopPath($profilePath) {
    # Scan de tous les chemins possibles du bureau (Local ou OneDrive institutionnel, FR ou EN)
    $possiblePaths = @(
        "$profilePath\Desktop",
        "$profilePath\Bureau",
        "$profilePath\OneDrive - Institut et Haute Ecole de la Santé La Source\Bureau",
        "$profilePath\OneDrive - Institut et Haute Ecole de la Santé La Source\Desktop",
        "$profilePath\OneDrive - Ecole la Source\Bureau",
        "$profilePath\OneDrive - Ecole la Source\Desktop",
        "$profilePath\OneDrive\Desktop",
        "$profilePath\OneDrive\Bureau"
    )
    foreach ($p in $possiblePaths) {
        if (Test-Path $p) { return $p }
    }
    throw "Dossier Bureau introuvable dans $profilePath"
}

function Restart-Explorer {
    try {
        Write-Host "  -> Redémarrage de l'explorateur Windows..." -ForegroundColor Yellow
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) { Start-Process explorer -ErrorAction Stop }
    } catch {
        Write-Warning "  -> Impossible de redémarrer l'explorateur : $_"
    }
}

function Invoke-BrowserCleaning {
    Write-Host "Nettoyage ciblé des navigateurs (Conservation Préférences & Versions)..." -ForegroundColor Cyan
    try {
        $processes = @("chrome", "msedge", "msedgewebview2")
        foreach ($proc in $processes) {
            Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 3 

        $targetProfile = Get-TargetProfilePath

        $profiles = @(
            "$targetProfile\AppData\Local\Microsoft\Edge\User Data\Default",
            "$targetProfile\AppData\Local\Google\Chrome\User Data\Default"
        )
        
        $itemsToRemove = @(
            "History", "History-journal", "Visited Links",
            "Cookies", "Cookies-journal", "Network\Cookies", "Network\Cookies-journal",
            "Login Data", "Login Data-journal", "Web Data", "Web Data-journal",
            "Sessions", "Session Storage", "Cache", "Code Cache", "GPUCache",
            "Service Worker\CacheStorage", "Service Worker\ScriptCache"
        )

        foreach ($profile in $profiles) {
            if (Test-Path $profile) {
                Write-Host "  -> Nettoyage : $profile"
                foreach ($item in $itemsToRemove) {
                    $targetPath = Join-Path $profile $item
                    if (Test-Path $targetPath) {
                        Remove-Item -Path $targetPath -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        return $true
    } catch {
        Write-Warning "  -> Erreur nettoyage navigateurs : $_"
        return $false
    }
}

function Invoke-DeepSystemCleaning {
    param([bool]$IncludeDownloads = $true)
    
    Write-Host "Nettoyage approfondi du stockage (Corbeille, Temp, Cache, etc.)..." -ForegroundColor Cyan
    $ok = $true

    try {
        # 1. Corbeille (Nettoyage standard + Force-Delete absolu sur C:)
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path "C:\`$Recycle.Bin" -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

        $targetProfile = Get-TargetProfilePath
        
        # 2. Téléchargements (Prend en compte le dossier "Downloads" ou "Téléchargements")
        if ($IncludeDownloads) {
            $downPath = "$targetProfile\Downloads"
            if (-not (Test-Path $downPath)) { $downPath = "$targetProfile\Téléchargements" }
            if (Test-Path $downPath) {
                Remove-Item "$downPath\*" -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        # 3. Fichiers Temporaires Windows ET Utilisateur cible
        @("C:\Windows\Temp\*", "$targetProfile\AppData\Local\Temp\*", "$env:TEMP\*") | ForEach-Object {
            Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue
        }

        # 4. Cache Windows Update
        Stop-Service wuauserv,DoSvc -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        Start-Service wuauserv,DoSvc -ErrorAction SilentlyContinue

        # 5. Cache des miniatures (Thumbnails)
        Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

        # 6. Cache DirectX
        if (Test-Path "$env:LOCALAPPDATA\D3DSCache") {
            Remove-Item "$env:LOCALAPPDATA\D3DSCache\*" -Recurse -Force -ErrorAction SilentlyContinue
        }

        # 7. Rapports d'erreurs Windows (WER)
        @("C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*", "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive\*", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue\*") | ForEach-Object {
            Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue
        }

        # 8. Fichiers Internet Temporaires globaux
        if (Test-Path "$env:LOCALAPPDATA\Microsoft\Windows\INetCache") {
            Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Warning "  -> Erreur critique nettoyage système : $_"
        $ok = $false
    }
    
    return $ok
}

# ==============================
# MODULE 1 : PC COURS
# ==============================
function Run-ModeCours {
    Write-Host "`n=== EXÉCUTION : 1. PC COURS ===" -ForegroundColor Cyan
    
    $desktopReset = {
        try {
            $targetProfile = Get-TargetProfilePath
            $desktopPath = Get-DesktopPath -profilePath $targetProfile
            Write-Host "  -> Profil ciblé : $targetProfile" -ForegroundColor Yellow
            Write-Host "  -> Bureau trouvé : $desktopPath" -ForegroundColor Gray
            
            # Nettoyage exclusif du bureau cible et du bureau public
            Get-ChildItem -Path @($desktopPath, [Environment]::GetFolderPath("CommonDesktopDirectory")) -Exclude "desktop.ini" -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            
            $wshShell = New-Object -ComObject WScript.Shell
            
            # L'ordre ici définit l'ordre d'empilement sur le bureau de haut en bas
            $apps = @(
                @{ Name="Microsoft Edge"; Path="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" },
                @{ Name="Google Chrome"; Path="C:\Program Files\Google\Chrome\Application\chrome.exe" },
                @{ Name="Word"; Path="C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" },
                @{ Name="Excel"; Path="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" },
                @{ Name="PowerPoint"; Path="C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" },
                @{ Name="ClickShare"; Path="$targetProfile\AppData\Local\ClickShare\ClickShare.exe" },
                @{ Name="Microsoft Teams"; Path="$targetProfile\AppData\Local\Microsoft\Teams\current\Teams.exe" }
            )
            
            foreach ($app in $apps) { 
                if (Test-Path $app.Path) { 
                    try {
                        $lnk = $wshShell.CreateShortcut("$desktopPath\$($app.Name).lnk")
                        $lnk.TargetPath = $app.Path
                        $lnk.Save() 
                        # Le petit délai magique (250ms) pour forcer Windows à les ranger proprement
                        Start-Sleep -Milliseconds 250
                    } catch { Write-Warning "  -> Échec création raccourci : $($app.Name)" }
                } 
            }
            return $true
        } catch {
            Write-Warning "  -> Exception Bureau : $_"
            return $false
        }
    }

    Show-Status "Nettoyage Stockage (Corbeille/Temp/Down)" (Invoke-DeepSystemCleaning -IncludeDownloads $true)
    Show-Status "Nettoyage Navigateurs (Chirurgical)" (Invoke-BrowserCleaning)
    Show-Status "Bureau Standard (Complet ordonné)" (&$desktopReset)
    Restart-Explorer
    try { Start-Process "ms-settings:windowsupdate" -ErrorAction SilentlyContinue } catch {}
    Write-Host "`nTERMINÉ. Mode PC Cours appliqué." -ForegroundColor Green
}

# ==============================
# MODULE 2 : PC VALIDATION
# ==============================
function Run-ModeValidation {
    Write-Host "`n=== EXÉCUTION : 2. PC VALIDATION ===" -ForegroundColor Cyan
    
    $desktopReset = {
        try {
            $targetProfile = Get-TargetProfilePath
            $desktopPath = Get-DesktopPath -profilePath $targetProfile
            Write-Host "  -> Profil ciblé : $targetProfile" -ForegroundColor Yellow
            Write-Host "  -> Bureau trouvé : $desktopPath" -ForegroundColor Gray
            
            Get-ChildItem -Path @($desktopPath, [Environment]::GetFolderPath("CommonDesktopDirectory")) -Exclude "desktop.ini" -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            
            $wshShell = New-Object -ComObject WScript.Shell
            $apps = @(
                @{ Name="Microsoft Edge"; Path="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" },
                @{ Name="Google Chrome"; Path="C:\Program Files\Google\Chrome\Application\chrome.exe" },
                @{ Name="Word"; Path="C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" }
            )
            foreach ($app in $apps) { 
                if (Test-Path $app.Path) { 
                    try {
                        $lnk = $wshShell.CreateShortcut("$desktopPath\$($app.Name).lnk")
                        $lnk.TargetPath = $app.Path
                        $lnk.Save() 
                        Start-Sleep -Milliseconds 250
                    } catch { Write-Warning "  -> Échec création raccourci : $($app.Name)" }
                } 
            }
            return $true
        } catch {
            Write-Warning "  -> Exception Bureau : $_"
            return $false
        }
    }

    $moodleSetup = {
        try {
            $startupFolder = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
            if (-not (Test-Path $startupFolder)) { New-Item -Path $startupFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null }
            $content = "[InternetShortcut]`nURL=microsoft-edge:https://moodle.ecolelasource.ch/login/index.php`nIconIndex=0`nIconFile=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            Set-Content -Path "$startupFolder\Moodle La Source.url" -Value $content -Force -ErrorAction Stop
            return $true
        } catch {
            Write-Warning "  -> Erreur raccourci Moodle : $_"
            return $false
        }
    }

    Show-Status "Nettoyage Stockage (Corbeille/Temp/Down)" (Invoke-DeepSystemCleaning -IncludeDownloads $true)
    Show-Status "Nettoyage Navigateurs (Historique/Cache)" (Invoke-BrowserCleaning)
    Show-Status "Bureau (Restreint)" (&$desktopReset)
    Show-Status "Moodle Démarrage" (&$moodleSetup)
    Restart-Explorer
    try { Start-Process "ms-settings:windowsupdate" -ErrorAction SilentlyContinue } catch {}
    Write-Host "`nTERMINÉ. Mode PC Validation appliqué." -ForegroundColor Green
}

# ==============================
# MODULE 3 : PC NOUVEAU COLLABORATEUR
# ==============================
function Run-ModeSetup {
    Write-Host "`n=== EXÉCUTION : 3. PC NOUVEAU COLLABORATEUR ===" -ForegroundColor Cyan
    
    $checkBitLocker = {
        try {
            $vol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
            if ($vol.ProtectionStatus -eq 'On') { Write-Host "  -> Chiffré ($($vol.EncryptionPercentage)%)" -ForegroundColor Green; return $true }
            return $false
        } catch { 
            Write-Warning "  -> Erreur BitLocker (Module manquant ou statut illisible) : $_"
            return $false 
        }
    }

    $configTaskbar = {
        try {
            $reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            Set-ItemProperty -Path $reg -Name "TaskbarAl" -Value 1 -ErrorAction Stop
            Set-ItemProperty -Path $reg -Name "ShowTaskViewButton" -Value 0 -ErrorAction Stop
            $search = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
            if (-not (Test-Path $search)) { New-Item -Path $search -Force -ErrorAction Stop | Out-Null }
            Set-ItemProperty -Path $search -Name "SearchboxTaskbarMode" -Value 4 -ErrorAction Stop
            return $true
        } catch {
            Write-Warning "  -> Erreur Registre Barre des tâches : $_"
            return $false
        }
    }

    $configScaling = {
        try {
            $reg = "HKCU:\Control Panel\Desktop"
            Set-ItemProperty -Path $reg -Name "Win8DpiScaling" -Value 1 -ErrorAction Stop
            Set-ItemProperty -Path $reg -Name "LogPixels" -Value 120 -ErrorAction Stop
            return $true
        } catch {
            Write-Warning "  -> Erreur Registre Mise à l'échelle : $_"
            return $false
        }
    }

    $configOffice = {
        try {
            foreach ($v in @("16.0", "15.0", "14.0")) {
                $reg = "HKCU:\Software\Microsoft\Office\$v\Common\Find\OpenDocuments"
                if (Test-Path $reg) { Set-ItemProperty -Path $reg -Name "ODFDefault" -Value 0 -ErrorAction Stop }
            }
            return $true
        } catch {
            Write-Warning "  -> Erreur Registre Office : $_"
            return $false
        }
    }

    $resetOutlook = {
        try {
            Start-Process -FilePath "outlook.exe" -ArgumentList "/resetfoldernames" -ErrorAction Stop
            return $true
        } catch {
            Write-Warning "  -> Erreur lancement Outlook : $_"
            return $false
        }
    }

    Show-Status "Statut BitLocker (C:)" (&$checkBitLocker)
    Show-Status "Barre des tâches (Centrée, Zone Recherche)" (&$configTaskbar)
    Show-Status "Mise à l'échelle (125%)" (&$configScaling)
    Show-Status "Format Office (Open XML par défaut)" (&$configOffice)
    Show-Status "Réinitialisation noms dossiers Outlook" (&$resetOutlook)
    Restart-Explorer
    Write-Host "`nTERMINÉ. Mode PC Nouveau collaborateur appliqué." -ForegroundColor Green
}

# ==============================
# MODULE 4.1 : PC LENTEUR/REPARATION (RAPIDE)
# ==============================
function Run-ModeRepair {
    Write-Host "`n--- EXÉCUTION : RÉPARATION RAPIDE ---" -ForegroundColor Cyan
    $logFile = "$env:USERPROFILE\Desktop\Rapport-Reparation-$(Get-Date -Format 'yyyyMMdd-HHmm').txt"
    try { Start-Transcript -Path $logFile -Force -ErrorAction Stop | Out-Null } catch {}

    $checkDisk = {
        try {
            $disks = Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName, HealthStatus
            $ok = $true
            foreach ($disk in $disks) {
                if ($disk.HealthStatus -eq 'Healthy') { Write-Host "  -> $($disk.FriendlyName) : SAIN" -ForegroundColor Green }
                else { Write-Warning "  -> $($disk.FriendlyName) : $($disk.HealthStatus)"; $ok = $false }
            }
            return $ok
        } catch {
            Write-Warning "  -> Erreur lecture SMART : $_"
            return $false
        }
    }

    $resetGpu = {
        try {
            Write-Host "L'écran va clignoter noir..." -ForegroundColor Yellow
            $gpus = Get-PnpDevice -Class Display -Status OK -ErrorAction Stop
            foreach ($gpu in $gpus) {
                Disable-PnpDevice -InstanceId $gpu.InstanceId -Confirm:$false -ErrorAction Stop
                Enable-PnpDevice -InstanceId $gpu.InstanceId -Confirm:$false -ErrorAction Stop
            }
            return $true
        } catch {
            Write-Warning "  -> Erreur réinitialisation GPU : $_"
            return $false
        }
    }

    $runSfc = {
        try {
            Write-Host "Lancement de SFC (Progression affichée ci-dessous)..." -ForegroundColor Yellow
            $sysNative = "$env:windir\Sysnative\sfc.exe"
            $sfcPath = if (Test-Path $sysNative) { $sysNative } else { "sfc.exe" }
            $sfcProc = Start-Process -FilePath $sfcPath -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru -ErrorAction Stop
            return ($sfcProc.ExitCode -eq 0)
        } catch {
            Write-Warning "  -> Erreur exécution SFC : $_"
            return $false
        }
    }

    Show-Status "Vérification Disques" (&$checkDisk)
    Show-Status "Nettoyage Stockage (Corbeille/Temp/Down)" (Invoke-DeepSystemCleaning -IncludeDownloads $true)
    Show-Status "Reset Graphique" (&$resetGpu)
    Show-Status "Réparation Système (SFC uniquement)" (&$runSfc)
    
    try { Start-Process "ms-settings:windowsupdate" -ErrorAction SilentlyContinue } catch {}
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Host "`nTERMINÉ. Rapport créé sur le bureau." -ForegroundColor Green
}

# ==============================
# MODULE 4.2 : PC RÉPARATION COMPLÈTE (LENT)
# ==============================
function Run-ModeRepairComplete {
    Write-Host "`n--- EXÉCUTION : RÉPARATION COMPLÈTE ---" -ForegroundColor Cyan
    $logFile = "$env:USERPROFILE\Desktop\Rapport-Reparation-Complete-$(Get-Date -Format 'yyyyMMdd-HHmm').txt"
    try { Start-Transcript -Path $logFile -Force -ErrorAction Stop | Out-Null } catch {}

    $checkDisk = {
        try {
            $disks = Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName, HealthStatus
            $ok = $true
            foreach ($disk in $disks) {
                if ($disk.HealthStatus -eq 'Healthy') { Write-Host "  -> $($disk.FriendlyName) : SAIN" -ForegroundColor Green }
                else { Write-Warning "  -> $($disk.FriendlyName) : $($disk.HealthStatus)"; $ok = $false }
            }
            return $ok
        } catch {
            Write-Warning "  -> Erreur lecture SMART : $_"
            return $false
        }
    }

    $resetGpu = {
        try {
            Write-Host "L'écran va clignoter noir..." -ForegroundColor Yellow
            $gpus = Get-PnpDevice -Class Display -Status OK -ErrorAction Stop
            foreach ($gpu in $gpus) {
                Disable-PnpDevice -InstanceId $gpu.InstanceId -Confirm:$false -ErrorAction Stop
                Enable-PnpDevice -InstanceId $gpu.InstanceId -Confirm:$false -ErrorAction Stop
            }
            return $true
        } catch {
            Write-Warning "  -> Erreur réinitialisation GPU : $_"
            return $false
        }
    }

    $runDism = {
        try {
            Write-Host "Lancement de DISM (Réparation de l'image locale via Internet, patientez)..." -ForegroundColor Yellow
            $sysNative = "$env:windir\Sysnative\dism.exe"
            $dismPath = if (Test-Path $sysNative) { $sysNative } else { "dism.exe" }
            $dismProc = Start-Process -FilePath $dismPath -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow -PassThru -ErrorAction Stop
            return ($dismProc.ExitCode -eq 0)
        } catch {
            Write-Warning "  -> Erreur exécution DISM : $_"
            return $false
        }
    }

    $runSfc = {
        try {
            Write-Host "Lancement de SFC (Progression affichée ci-dessous)..." -ForegroundColor Yellow
            $sysNative = "$env:windir\Sysnative\sfc.exe"
            $sfcPath = if (Test-Path $sysNative) { $sysNative } else { "sfc.exe" }
            $sfcProc = Start-Process -FilePath $sfcPath -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru -ErrorAction Stop
            return ($sfcProc.ExitCode -eq 0)
        } catch {
            Write-Warning "  -> Erreur exécution SFC : $_"
            return $false
        }
    }

    Show-Status "Vérification Disques" (&$checkDisk)
    Show-Status "Nettoyage Stockage (Corbeille/Temp/Down)" (Invoke-DeepSystemCleaning -IncludeDownloads $true)
    Show-Status "Reset Graphique" (&$resetGpu)
    Show-Status "Restauration Image Système (DISM)" (&$runDism)
    Show-Status "Réparation Fichiers Système (SFC)" (&$runSfc)
    
    try { Start-Process "ms-settings:windowsupdate" -ErrorAction SilentlyContinue } catch {}
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Host "`nTERMINÉ. Rapport complet créé sur le bureau." -ForegroundColor Green
}

# ==============================
# MENU 10 : INFORMATIONS
# ==============================
function Show-Informations {
    Write-Host "`n=== INFORMATIONS SUR LES SCRIPTS ===" -ForegroundColor Cyan
    Write-Host "1. PC Cours :" -ForegroundColor Green
    Write-Host "   - Nettoyage complet du système (Téléchargements, Cache divers, Corbeille)."
    Write-Host "   - Nettoyage des navigateurs en conservant les préférences/pages d'accueil."
    Write-Host "   - Remet les raccourcis bureau complets (Edge, Chrome, Word, Excel, PowerPoint, ClickShare)."
    Write-Host ""
    Write-Host "2. PC Validation :" -ForegroundColor Yellow
    Write-Host "   - Nettoyage complet du système (Téléchargements, Cache divers, Corbeille)."
    Write-Host "   - Nettoyage des navigateurs (Historique, cache, sessions, mdp)."
    Write-Host "   - Remet des raccourcis bureau restreints (Edge, Chrome, Word)."
    Write-Host "   - Configure Moodle pour s'ouvrir au démarrage de la session."
    Write-Host ""
    Write-Host "3. PC Nouveau collaborateur :" -ForegroundColor Magenta
    Write-Host "   - Vérifie le statut de chiffrement BitLocker."
    Write-Host "   - Configure l'interface (Barre des tâches centrée, Recherche)."
    Write-Host "   - Règle la mise à l'échelle d'affichage à 125%."
    Write-Host "   - Configure Office (Format Open XML par défaut) et réinitialise Outlook."
    Write-Host ""
    Write-Host "4. PC Lenteur/Reparation :" -ForegroundColor Red
    Write-Host "   -> Ouvre un sous-menu avec deux options :"
    Write-Host "      - Rapide : Dépannage express (5-10 min) avec SFC."
    Write-Host "      - Complet : Ajoute DISM pour retélécharger l'image saine (15-30 min)."
}

# ==============================
# MENU PRINCIPAL INTERACTIF
# ==============================
$menuOpen = $true

while ($menuOpen) {
    Clear-Host
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "                   OPTI-SOURCE                    " -ForegroundColor White -BackgroundColor Blue
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host " Version : 1.0" -ForegroundColor DarkGray
    Write-Host " Auteur  : Dylan Martins Fernandes" -ForegroundColor DarkGray
    Write-Host " Date    : $(Get-Date -Format 'dd.MM.yyyy HH:mm')" -ForegroundColor DarkGray
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host " ATTENTION : Les actions de ce script sont" -ForegroundColor Red
    Write-Host " irréversibles. L'auteur décline toute" -ForegroundColor Red
    Write-Host " responsabilité en cas d'erreur de choix." -ForegroundColor Red
    Write-Host " Consultez [10] Informations en cas de doute." -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] PC Cours" -ForegroundColor Green
    Write-Host "  [2] PC Validation" -ForegroundColor Yellow
    Write-Host "  [3] PC Nouveau collaborateur" -ForegroundColor Magenta
    Write-Host "  [4] PC Lenteur/Reparation" -ForegroundColor Red
    Write-Host ""
    Write-Host "  [10] Informations" -ForegroundColor Blue
    Write-Host ""
    Write-Host "  [Q] Quitter" -ForegroundColor Gray
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $choice = Read-Host "Sélectionnez une option"
    $actionTerminee = $false
    
    switch ($choice) {
        '1' { Run-ModeCours; $actionTerminee = $true }
        '2' { Run-ModeValidation; $actionTerminee = $true }
        '3' { Run-ModeSetup; $actionTerminee = $true }
        '4' { 
            Write-Host "`n=== SOUS-MENU : PC LENTEUR/REPARATION ===" -ForegroundColor Cyan
            Write-Host "  [1] Rapide (SFC uniquement)" -ForegroundColor Red
            Write-Host "  [2] Complet (DISM + SFC)" -ForegroundColor DarkRed
            Write-Host "  [M] ou [Entrée] Retour au menu principal" -ForegroundColor Gray
            Write-Host ""
            $repChoice = Read-Host "Choisissez le type de réparation"
            
            if ($repChoice -eq '1') { Run-ModeRepair; $actionTerminee = $true }
            elseif ($repChoice -eq '2') { Run-ModeRepairComplete; $actionTerminee = $true }
            else { 
                $actionTerminee = $false 
            }
        }
        '10' { 
            Show-Informations
            Write-Host ""
            Read-Host "Appuyez sur Entrée pour revenir au menu principal..."
            $actionTerminee = $false 
        }
        'q' { 
            $menuOpen = $false 
        }
        Default { 
            Write-Warning "Choix invalide." 
            Start-Sleep -Seconds 1 
        }
    }
    
    if ($actionTerminee -and $menuOpen) {
        Write-Host "`n==================================================" -ForegroundColor Cyan
        Write-Host "  L'action est terminée." -ForegroundColor Cyan
        Write-Host "  [M] ou [Entrée] : Revenir au menu principal" -ForegroundColor Green
        Write-Host "  [Q]             : Quitter le script" -ForegroundColor Yellow
        Write-Host "==================================================" -ForegroundColor Cyan
        
        $endChoice = Read-Host "Votre choix"
        if ($endChoice -match "^[qQ]$") {
            $menuOpen = $false
        }
    }
}