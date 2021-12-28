Function Install-O365([String] $SiteCode = "Generic"){
    <#Download the installer from https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117
    Extract the Setup.exe and upload it somewhere
    It helps to keep it up to date
    https://docs.microsoft.com/en-us/officeupdates/odt-release-history
    #>
    $O365Setupexe = "http://URLHERE/365Setup.exe"
    Write-Host "Downloading MS Office"
        Enable-SSL
        New-Item -ItemType Directory -Force -Path "C:\IT\O365"
        (New-Object System.Net.WebClient).DownloadFile($O365Setupexe, 'C:\IT\O365\setup.exe')
    Write-Host "Downloading MS Office config files"
        $O365ConfigSource = "https://raw.githubusercontent.com/shiftnerd/PSInstall365/main/configs/O365Business-x64.xml"
        $O365ConfigDest = "C:\IT\O365\" + $SiteCode + "_O365_Config.xml"
        (New-Object System.Net.WebClient).DownloadFile($O365ConfigSource, $O365ConfigDest)
    Write-Host "Installing Office"
        & C:\IT\O365\setup.exe /configure $O365ConfigDest | Wait-Process
    Write-Host "Placing Shortcuts"
        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Outlook.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Excel.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Word.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
}
