#------------- Programmkopf ---------------#
# Autor: Marticc Igor
# Version: 1.0.0
# Datum: 20.06.2022
# Titel: Dateienordnungsprogramm
# Beschreibung:
# Dateien von einem Modul werden in BBBaden-Ordner verschoben
#------------------------------------------#
$logfile = "log.txt"

if (Test-Path -path $logfile) {
    Remove-Item -path $logfile -Force
    New-Item -path $logfile -type file -force
}

#-----------Change this part to your needs-----------#
$TargetPath = "C:\Schule\BBBaden\Informatik" # Where you want to create the folder
$DownloadsPath = "C:\Users\Igor Martic\Dropbox\PC\Downloads" # Where you want to pull the files from
#----------------------------------------------------# 
$Sorting = $true


While ($Sorting -eq $true) {
    $Deletable = New-Object System.Collections.ArrayList

    $ifDelete = $true
    $ifContinue = $true
    $NumberInput = $true

    #Checks if the input Number is a valid Number
    while ($NumberInput -eq $true) {
        try {
            Write-Host -ForegroundColor Cyan "Welches Modul moechten Sie sortieren"
            $folderNumber = Read-Host  
            If (($folderNumber.length -ne 3) -or ($folderNumber -match "[a-z]") ) {
                throw
            }
            $NumberInput = $false
        }
        catch {
            Write-Host -ForegroundColor Red "Ungueltige Eingabe" 
        }
    }
    Write-Output "-----------------------------------------------------------------------------------" >> $logfile
    Write-Output "Modul: $folderNumber `n" >> $logfile


    $FolderPath = "$TargetPath\M$folderNumber"
    
    #Creates Folder if not already present
    if (!(Test-Path -Path "C:\Schule\BBBaden\Informatik\M$folderNumber")) {
        Write-Host -ForegroundColor Yellow "Ein Ordner fuer dieses Modul ist nicht vorhanden, ein neues wird erstellt."
        New-Item -path "$TargetPath" -type directory -name "M$folderNumber" 
        foreach ($dir in ("Aufgaben", "Powerpoints", "Loesungen", "Zips")) {
            
            New-Item -path "$FolderPath" -type directory -name "$dir"
        }
        Write-Output "Ordner $FolderPath erstellt" >> $logfile
    } 
    else {
        foreach ($dir in ("Aufgaben", "Powerpoints", "Loesungen", "Zips")) {
            if (!(Test-Path "$FolderPath\$dir")) {
                New-Item -path "$FolderPath" -type directory -name "$dir"
                Write-Output "Ordner $dir wurde in $FolderPath erstellt" >> $logfile
            }
        }
    }

    Write-Output "`n" >> $logfile

    #Searches for files containing the input Number in their name
    $Verschieben = Get-ChildItem -path "$DownloadsPath" -Filter "*$folderNumber*"

    Write-Output "Verschobene Dateien: `n" >> $logfile
    #Moves the files to the correct folder
    foreach ($item in $Verschieben) {
        if (($item.extension -in ".docx") -and (!(Test-Path -Path "$FolderPath\Aufgaben\$item")) -and (!(Test-Path -Path "$FolderPath\Loesungen\$item"))) {
            if ($item.fullname -like "*_L.docx") {
                Move-Item -path $item.fullname -destination "$FolderPath\Loesungen" 
            }
            else {
                Move-Item -path $item.fullname -destination "$FolderPath\Aufgaben" 
            }

            Write-Output "$item" >> $logfile
        }
        elseif (($item.extension -in ".pptx") -and (!(Test-Path -Path "$FolderPath\Powerpoints\$item"))) {
            Move-Item -path $item.fullname -destination "$FolderPath\Powerpoints"
            Write-Output "$item" >> $logfile
        }
        elseif (($item.extension -in ".xlsx") -or ($item.extension -in ".pdf") -and (!(Test-Path -Path "$FolderPath\$item"))) {
            Move-Item -path $item.fullname -destination "$FolderPath"
            Write-Output "$item" >> $logfile
        }
        elseif (($item.extension -in ".fs") -and (!(Test-Path -Path "$FolderPath\FsFiles\$item"))) {
            if (!(Test-Path "$FolderPath\FsFiles")) {
                New-Item -path "$FolderPath" -type directory -name "FsFiles"
            }
            Move-Item -path $item.fullname -destination "$FolderPath\FsFiles"
            Write-Output "$item" >> $logfile
        }
        elseif (($item.extension -in ".zip") -or ($item.extension -in ".7z")) {
            $filename = $item.Basename;
            if (!(Test-Path -Path "$FolderPath\Zips\$filename")) {
                New-Item -path "$FolderPath\Zips" -type directory -name "$filename" 
                Expand-Archive -path $item.fullname -DestinationPath "$FolderPath\Zips\$filename"
                Write-Output "$item" >> $logfile
            }
        }
        elseif (($item.extension -in ".docx") -or ($item.Extension -in ".pptx") -or ($item.Extension -in ".fs") -or ($item.Extension -in ".xlsx") -or ($item.Extension -in ".pdf")) { 
            Write-Host -ForegroundColor Yellow "$item existiert bereits in $FolderPath"
            $Deletable += $item
        }
    }

    Write-Output "`n" >> $logfile

    # Gives option to delete files from downloads-folder that are already in the correct folder
    if ($Deletable.count -gt 0) {
        while ($ifDelete -eq $true) {
            Write-Host -ForegroundColor Cyan "Wollen Sie die doppelten Dateien aus Ihrem Downloads Ordner loeschen? (j/n)"
            $DeleteOrKeep = Read-Host 
            if ($DeleteOrKeep -eq "j") {
                $ifDelete = $false
                Write-Output "Geloeschte Dateien: `n" >> $logfile
                foreach ($item in $Deletable) {
                    Remove-Item -path $item.fullname
                    write-output "$item " >> $logfile
                }
                Write-Host -ForegroundColor Yellow "Die doppelten Dateien wurden geloescht"
            }
            elseif ($DeleteOrKeep -eq "n") {
                $ifDelete = $false
            }
            else {
                Write-Output "Ungueltige Eingabe"
            }
        }
    }

    #Asks if the user wants to continue and checks if valid input
    while ($ifContinue -eq $true) {
        Write-Host -ForegroundColor Cyan "Wollen Sie noch ein Modul sortieren? (j/n)"
        $Answer = Read-Host 
        if ($Answer -eq "j") {
            $ifContinue = $false
        }
        elseif ($Answer -eq "n") {
            $ifcontinue = $false
            $Sorting = $false
        }
        else {
            Write-Output "Ungueltige Eingabe"
        }
    }
}
