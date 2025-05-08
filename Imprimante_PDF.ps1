<########################################################################################################################################################################

                                                        IMPRESSIONS VIRTUELLES MULTIPLES FICHIERS PDF (FORMAT A4)
                                                        ------------------------------------------------                                                             
# REQUIS:
- Logiciel Adobe Acrobat Reader DC
- Programme par defaut pour les PDF 
########################################################################################################################################################################>
Clear-Host
# LANCEMENT DU SCRIPT EN TANT QU'ADMIN
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
exit $LASTEXITCODE
}
#
Clear-Host
#
If (Test-Path -Path "$($env:ProgramFiles)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"){
  
                    Write-Host "
---------------------------------------------------------------------------------------------------
                                LANCEMENT DU SCRIPT
---------------------------------------------------------------------------------------------------
                               " -ForegroundColor Green
Start-Sleep -Seconds 3
Clear-Host  
    
    
    Write-Host "LE LOGICIEL ACROBAT READER EST INSTALLE SUR LE POSTE" -ForegroundColor Green
    Start-Sleep -Seconds 2
    Clear-Host

    # VERIFICATION DU DOSSIER SOURCE
    If((Test-Path -Path "$home\Desktop\PDF_A_IMPRIMER") -eq $false){

        Write-Host "Le dossier Source n'est pas present pour le script !" -ForegroundColor Red
        Start-Sleep -Seconds 3
        Clear-Host
        Write-Host "Creation du dossier 'PDF_A_IMPRIMER' pour le depot" -ForegroundColor Green
        New-Item -Path "$home\Desktop\" -ItemType Directory -Name "PDF_A_IMPRIMER" -Force
        Start-Sleep -Seconds 3
        Write-Host "Vous pouvez desormais deposer des fichiers PDF dans le dossier 'PDF_A_IMPRIMER' et relancer le script." -ForegroundColor Green
        Start-Sleep -Seconds 3
    }

    # VERIFICATION DU DOSSIER DESTINATION
    If((Test-Path -Path "$home\Desktop\PDF_IMPRIMER") -eq $false){

        Write-Host "Le dossier de destination n'est pas present pour le script !" -ForegroundColor Red
        Start-Sleep -Seconds 3
        Clear-Host
        Write-Host "Creation du dossier 'PDF_IMPRIMER' pour l'impression" -ForegroundColor Green
        New-Item -Path "$home\Desktop\" -ItemType Directory -Name "PDF_IMPRIMER" -Force
        Start-Sleep -Seconds 3  
        Clear-Host
        Write-Host "Le dossier de destination 'PDF_IMPRIMER' est present sur le bureau." -ForegroundColor Green
        Write-Host "Vous pouvez desormais deposer des fichiers PDF dans le dossier 'PDF_A_IMPRIMER' et relancer le script." -ForegroundColor Green
        Start-Sleep -Seconds 3
    }

    # VERIFICATION FICHIERS OU DOSSIER PRESENT DANS LE DOSSIER SOURCE/DESTINATION
    If ((Get-ChildItem -Path "$Home\Desktop\PDF_IMPRIMER").count -ge "1"){
        Write-Host "AVERTISSEMENT" -ForegroundColor White -BackgroundColor Red 
        Write-Host "! ATTENTION DES FICHIERS SONT PRESENTS DANS LE DOSSIER DESTINATION 'PDF_IMPRIMER' SUR LE BUREAU !" -ForegroundColor Red
        Start-Sleep -Seconds 5
        Write-Host "! NOUS VOUS CONSEILLONS DE LES DEPLACER CAR CELA PEUT POSER PROBLEME POUR LE SCRIPT !" -ForegroundColor Red  
        Start-Sleep -Seconds 7
        Clear-Host
        Remove-Item -Path "$home\Desktop\PDF_IMPRIMER\*" -Force -Recurse -Verbose -Confirm:$true
        Clear-Host
        Write-Host "SI VOUS AVEZ REFUSE OU ACCEPTER, LE SCRIPT VA CONTINUER." -ForegroundColor Red
        Write-Host "SI UN DOCUMENT DETIENT LE MEME NOM QUE DANS LE DOSSIER DE DESTINATION, IL SERA ECRASE." -ForegroundColor Red
        Start-Sleep -Seconds 7      
    }

# EXECUTION DU SCRIPT    
# PDF PRESENT DANS LE DOSSIER SOURCE
    If ((Get-ChildItem -Path "$Home\Desktop\PDF_A_IMPRIMER" -Filter "*.pdf").count -ge "1"){

    $PDF_Count = (Get-ChildItem -Path "$Home\Desktop\PDF_A_IMPRIMER" -Filter "*.pdf").count
    Start-Sleep -Seconds 4
    Clear-Host
    Write-Host "$PDF_Count FICHIERS PDF VONT ETRE IMPRIMES"
    Start-Sleep -Seconds 4
    Clear-Host
    #
            # PREPARATION/NETTOYAGE DU PC AVANT LANCEMENT
            If((Get-Process -Name "AcroRd32" -ErrorAction SilentlyContinue).Count -ge "1"){Stop-Process -name AcroRd32}
            Start-Sleep -Seconds 2
            If(Test-path -Path "$env:temp\PDFResultFile.pdf"){Remove-Item -Path "$env:temp\PDFResultFile.pdf"}
            Write-Host "REDEMARRAGE SERVICE SPOOLER" -ForegroundColor Red
            Start-Sleep -Seconds 3 
            Stop-Service -Name Spooler -Force
            Clear-Host
            Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*" -Recurse -Force -Verbose
            Clear-Host
            Start-Service -Name Spooler
            Write-Host "REDEMARRAGE SERVICE SPOOLER OK" -ForegroundColor Green 
            Start-Sleep -Seconds 3
            Clear-Host
    
                # Installation Imprimante
                if((Get-Printer).Name -match "IMPRIMANTE_VIRTUELLE"){
                    Write-Host "
        ---------------------------------------------------------------------------------------------------
                                    L'IMPRIMANTE EST DEJA INSTALLEE
        ---------------------------------------------------------------------------------------------------
                                " -ForegroundColor Green
                Start-Sleep -Seconds 3
                Clear-host
                }

                Else{
            
                    Write-Host "
        ---------------------------------------------------------------------------------------------------
                                    L'IMPRIMANTE N'EST PAS INSTALLEE
        ---------------------------------------------------------------------------------------------------
                                " -ForegroundColor Red
                Start-Sleep -Seconds 3
                Clear-host

                # Definition du nom de l'imprimante
                $printerName = 'IMPRIMANTE_VIRTUELLE'
                # Chemin par defaut (Enregistrement des PDF apres Impression)
                $PDFFilePath = "$env:temp\PDFResultFile.pdf"

                # Verification du port de sortie
                $port = Get-PrinterPort -Name $PDFFilePath -ErrorAction SilentlyContinue
                if ($null -eq $port){
    
                    # Création du port de sortie sur "Microsoft Print to PDF"
                    Add-PrinterPort -Name $PDFFilePath 
                }

                # Ajout de l'imprimante
                Add-Printer -DriverName "Microsoft Print to PDF" -Name $printerName -PortName $PDFFilePath 
        
                # Imprimante par defaut
                $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='IMPRIMANTE_VIRTUELLE'"
                Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
                Start-Sleep -Seconds 3
                Clear-Host
                Write-Host "L'imprimante 'IMPRIMANTE_VIRTUELLE' est installe et mis pas defaut" -ForegroundColor Green
                Start-Sleep -Seconds 3
                Clear-Host
                }
        #
        ####################################################################################
        # TRAITEMENT PDF
        #
        $files = Get-ChildItem "$Home\Desktop\PDF_A_IMPRIMER\" | Sort-Object

        foreach ($file in $files){
    
            Start-process -FilePath $file.fullName -Verb Print -WindowStyle Minimized 

                Start-Sleep -Seconds 2
        
            Write-Host "
        ---------------------------------------------------------------------------------------------------
                        LANCEMENT DE L'IMPRESSION DU DOCUMENT '$($file.Name)'
            " -ForegroundColor Red

                #Verification Spooler si en attente d'impression (valeur 0 JobsSpooling) - Chargement
                Write-Host "EN ATTENTE IMPRESSION CHARGEMENT" -ForegroundColor Red
                Start-Sleep -Seconds 2
                do
                {
                $printerName = "IMPRIMANTE_VIRTUELLE"
                $PrintJobsSpooling = (Get-CimInstance -Class Win32_PerfFormattedData_Spooler_PrintQueue | 
                Where-Object -Property Name -eq $printerName | 
                Select-Object JobsSpooling).JobsSpooling
                }
                    Until($PrintJobsSpooling -eq 0)
        
                Write-Host "EN ATTENTE IMPRESSION CHARGEMENT OK" -ForegroundColor Green
                Start-Sleep -Seconds 1
        
                Write-Host "KILL PROCESS ADOBE" -ForegroundColor Red
                # Kill Process Adobe
                if((Get-Process -Name "AcroRd32" -ErrorAction SilentlyContinue).Count -ge "1"){Stop-Process -name AcroRd32}
                Start-Sleep -Seconds 2
                Write-Host "KILL PROCESS ADOBE OK" -ForegroundColor Green
                Start-Sleep -Seconds 1

                #Verification Spooler si impression terminee (valeur 0 jobs) 
                Write-Host "IMPRESSION EN COURS" -ForegroundColor Red
                Start-Sleep -Seconds 1
                do
            {
                $printerName = "IMPRIMANTE_VIRTUELLE"
                $PrintJobs = (Get-CimInstance -Class Win32_PerfFormattedData_Spooler_PrintQueue | 
                Where-Object -Property Name -eq $printerName | 
                Select-Object Jobs).Jobs
            }
                Until ($PrintJobs -eq 0)
    
            Write-Host "
                                IMPRESSION TERMINEE
            " -ForegroundColor Green  
               Start-Sleep -Seconds 1


            if((Get-Process -Name "AcroRd32" -ErrorAction SilentlyContinue).Count -eq "0"){
                $file_PDF = $file.Name
                Rename-Item -Path "$Home\AppData\Local\Temp\PDFResultFile.pdf" -NewName "$Home\AppData\Local\Temp\$file_PDF" -Verbose
                #
                Move-Item -Path "$Home\AppData\Local\Temp\$file_PDF" -Destination "$home\Desktop\PDF_IMPRIMER" -Force -Verbose
            Write-Host "
        ---------------------------------------------------------------------------------------------------"
                Stop-Process -Name "AcroRd32" -ErrorAction Ignore
                }
    
            }
        #
        If(Test-path -Path "$env:temp\PDFResultFile.pdf"){
        Remove-Item -Path "$env:temp\PDFResultFile.pdf"}
        #
        Clear-Host
        Write-Host "FIN DU SCRIPT :)" -ForegroundColor Green
        Start-Sleep -Seconds 2  
        }
}
Else{
Write-Host "LE LOGICIEL ADOBE READER DC N'EST PAS INSTALLE SUR LE POSTE." -ForegroundColor Red
Write-Host "VEUILLEZ INSTALLER LE LOGICIEL SUR LE POSTE PUIS RELANCER." -ForegroundColor Red
Start-Sleep -Seconds 5
Pause
}
########################################################################################################################################################################