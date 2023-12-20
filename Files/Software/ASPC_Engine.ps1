#ASPC engine. 


#-------Define Variables (outside Engine Config File)-------



#-------Functions-------

Function updateChecker {
    $Error.Clear()
        Write-Output "Running update check"
        Write-Output "Checking to see if its the correct day to run update check"
        $updateRunDate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
        $todaysDate = Get-Date -Hour 0 -Minute 0 -Second 0
        if ($updateRunDate.day -eq $todaysDate.day) {
        
        $getVersionsWeb = Invoke-WebRequest -URI "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/versions.json" | ConvertFrom-Json
        $engVersionWeb = $getVersionsWeb.Eng_Ver
        $confVersionWeb = $getVersionsWeb.Conf_Ver

        $getVersionsLocal = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Version\versions.json" | ConvertFrom-Json
            $engVersionLocal = $getVersionsLocal.Eng_Ver
            $confVersionLocal = $getVersionsLocal.Conf_Ver

        if ($engVersionWeb -gt $engVersionLocal) {
               $EngupdateAvail = "1"  
            } else {
               $EngupdateAvail ="0"
            }
        if ($confVersionWeb -gt $confVersionLocal) {
                $ConfupdateAvail = "1"
            } else {
                $ConfupdateAvail = "0"
            }
        if ($EngupdateAvail -eq "1" -OR $ConfupdateAvail -eq "1") {
            $Global:updateMessage = "There is an update available for ASPC. Please use the Configurator to run the update."
            } else { 
            $Global:updateMessage = ""
    }
    Write-Output $Global:updateMessage
    Write-Output $Error
	PullEngineConfigs
    } Else {
    Write-Output "No Update available"
    Write-Output $Error
    PullEngineConfigs
    }
}

#Start by pulling Engine configs from JSON file
Function PullEngineConfigs {
    Write-output "Opening Engine config file to get variables"
    $Error.Clear()
    try {
        $Global:engConfJson =  Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\Eng_Config.json" | ConvertFrom-Json -Verbose
        Write-Output "Config file opened and read"
        Write-Output $Error
        SetEngVariables
        

} catch {
    Write-Output "There was an issue importing the Engine Variables, please investigate"
    Write-Output $Error
    Write-Output "Stopping Engine"
    EngineStop
    }      
}  

#Set Engine configs based off the read in JSON file in the PullEngineConfigs function
Function SetEngVariables {
    $Error.Clear()
    Write-Output "Setting Engine variables"
        $Global:DefaultEmail = $Global:engConfJson.DefaultEmail
        $Global:XcludeExpired = $Global:engConfJson.XcludeExpired
        $Global:ExpireWithin = $Global:engConfJson.ExpireWithin
        $Global:ReportName = $Global:engConfJson.ReportName
    Write-Output "Variables set to..."
        $Global:DefaultEmail
        $Global:XcludeExpired
        $Global:ExpireWithin
        $Global:ReportName
    Write-Output "Moving to Main Engine"
    Write-Output $Error
    Engine_LoadCreds_Login
}

#Engine start. First function pulls credentials from files and then logs into Azure
Function Engine_LoadCreds_Login {
    $Error.Clear()
    Write-Output "Starting Main Engine Components"
    Write-Output "Creating Credential"

        $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"

        $EncThumb = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\spThumb.txt" | ConvertTo-SecureString -Key $ASPCKeyData
        $PlainThumb = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($EncThumb))

        $EncAppID = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\spAppID.txt" | ConvertTo-SecureString -Key $ASPCKeyData
        $PlainAppID = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($EncAppID))

        $EncTennant = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\spTennantID.txt" | ConvertTo-SecureString -Key $ASPCKeyData
        $PlainTennant = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($EncTennant))

    
    Write-Output "Credential created"
    Write-Output "Trying to Connect to Azure"
        try {
            Connect-AzureAD -TenantId $PlainTennant -ApplicationId $PlainAppID -CertificateThumbprint $PlainThumb -Verbose
            Write-Output "Connected to Azure completed"
            Write-Output "Moving to Get ServicePrincipals"
            Write-Output $Error
            Engine_GetSP_KeysNCerts
        } catch {
            Write-Output "There was an issue connecting to Azure. Stopping Engine."
            Write-Output $Error
            EngineStop
    }
}

#Pull the Service Principals, keys and cert expirations
Function Engine_GetSP_KeysNCerts {
    $Error.Clear()
    Write-Output "Start pulling in Azure SP Certificates and Keys"
    $Global:SPKeysCerts = @()
        Get-AzureADApplication -All $true | ForEach-Object{  
            $app = $_
            $owner = Get-AzureADApplicationOwner -ObjectId $_.ObjectID -Top 1
            (Get-AzureADApplication -ObjectId $_.ObjectID).PasswordCredentials | 
            ForEach-Object{ 
                $Global:SPKeysCerts += [PSCustomObject] @{
                        CredentialType = "PasswordCredentials"
                        DisplayName = $app.DisplayName; 
                        ExpiryDate = $_.EndDate;
                        StartDate = $_.StartDate;
                        KeyID = $_.KeyId;
                        Type = 'NA';
                        Usage = 'NA';
                        Owners = $owner.UserPrincipalName;
                    }
                } 
            (Get-AzureADApplication -ObjectId $_.ObjectID).KeyCredentials | 
            ForEach-Object{ 
                $Global:SPKeysCerts += [PSCustomObject] @{
                        CredentialType = "KeyCredentials"                                        
                        DisplayName = $app.DisplayName; 
                        ExpiryDate = $_.EndDate;
                        StartDate = $_.StartDate;
                        KeyID = $_.KeyId;
                        Type = $_.Type;
                        Usage = $_.Usage;
                        Owners = $owner.UserPrincipalName;
            }
        } 
                             
    } 
    Write-Output $Error
    Engine_Parse_Expiry_Soon 
}

#Parse out anything that is expiring within the next "X" amount of days (defined by user in configurator,default 30. Can only be between 15 and 90)
#OR anything already expired
Function Engine_Parse_Expiry_Soon {
    $Error.Clear()
    Write-Output "Started Engine_Parse_Expiry_soon"
    $Global:ExpireKeysCertsSoon =@()
    $Global:ExpiryDate = (Get-Date).AddDays($Global:ExpireWithin)
    $Global:TodaysDate = Get-Date

    foreach($SPKC in $Global:SPKeysCerts) {
        $Global:keyend = $SPKC.ExpiryDate
        
            if ($SPKC.ExpiryDate -le $Global:ExpiryDate -AND $SPKC.ExpiryDate -gt $Global:TodaysDate) {
            $ExSoonDaysBetween = New-Timespan -Start $Global:TodaysDate -End $Global:keyend
            $ExSoonDaysBetweenDays = $ExSoonDaysBetween.Days
                $SPKC | Add-Member -NotePropertyName DaysToExpire -NotePropertyValue "   $ExSoonDaysBetweenDays" -Force
                $Global:ExpireKeysCertsSoon += @($SPKC)
    
            }
      }
      Write-Output $Error
      Engine_Parse_Expiry_Already    
}

Function Engine_Parse_Expiry_Already {
    $Error.Clear()
    Write-Output "Started Engine_Parse_Expiry_Already"
    $Global:ExpireKeysCertsAlready =@()
    $Global:TodaysDate = Get-Date

    foreach($SPKCAll in $Global:SPKeysCerts) {
        
            if ($SPKCAll.ExpiryDate -lt $Global:TodaysDate) {
                $ExAlreadyDaysBetween = New-Timespan -Start $SPKCAll.ExpiryDate -End $Global:TodaysDate
                $ExAlreadyDaysBetweenDays = $ExAlreadyDaysBetween.Days
                    $SPKCAll | Add-Member -NotePropertyName ExpiredDaysAgo -NotePropertyValue "Expired -$ExAlreadyDaysBetweenDays days ago" -Force
                    $Global:ExpireKeysCertsAlready += @($SPKCAll)
    
            }
    }
    Write-Output $Error
    Check_JSON_Notes_Expiring
}

Function Check_JSON_Notes_Expiring {
        $Error.Clear()
        Write-Output "Started Check_JSON_Notes_Expiring"
        $ListNoteFiles = (Get-Childitem "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\").BaseName
        foreach($AlrtDef in $Global:ExpireKeysCertsSoon) {

        
            $CompareNames = Compare-Object -ReferenceObject $AlrtDef.DisplayName -DifferenceObject $ListNoteFiles -IncludeEqual
                foreach($CompName in $CompareNames) {
                    If($CompName.SideIndicator -eq "==") {
                        $SPFinName = $CompName.InputObject
                        $SPFinNameInfo = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPFinName.json" | ConvertFrom-Json -Verbose
                        $SPNotes = $SPFinNameInfo.SPNotes
						$InternalNotes = $SPFinNameInfo.IntOwner
						$ExternalNotes = $SPFinNameInfo.ExtOwner
                            $AlrtDef | Add-Member -NotePropertyName Notes -NotePropertyValue $SPNotes -Force
							$AlrtDef | Add-Member -NotePropertyName Internal -NotePropertyValue $InternalNotes -Force
							$AlrtDef | Add-Member -NotePropertyName External -NotePropertyValue $ExternalNotes -Force
                            $Global:ExpireKeysCertsSoon += @($AlrtDef)
                    }
              }
       }
       $GetUniqueEntries = $Global:ExpireKeysCertsSoon | Sort-Object DisplayName | Get-Unique -AsString
       foreach($AlrtDefXmail in $GetUniqueEntries) {
                $ListNoteFilesXmail = (Get-Childitem "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\").BaseName
                $CompareNamesXmail = Compare-Object -ReferenceObject $AlrtDefXmail.DisplayName -DifferenceObject $ListNoteFilesXmail -IncludeEqual
                foreach($CompNameXmail in $CompareNamesXmail) {
                    If($CompNameXmail.SideIndicator -eq "==" -AND $AlrtDefXmail.DaystoExpire -le '15') {
                        $SPFinNameXmail = $CompNameXmail.InputObject
                        $SPFinNameInfoXmail = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPFinNameXmail.json" | ConvertFrom-Json -Verbose
                        $Global:ExternalMailAddr = $SPFinNameInfoXmail.ExternalAlertMailAddr
                        $Global:ExternalMailNote = $SPFinNameInfoXmail.ExternalAlertMailNote
                        Mail_Handler_Secondary
                }
            }
       }
       Write-Output $Error
       Check_JSON_Notes_Expired
}

Function Check_JSON_Notes_Expired {
        $Error.Clear()
        Write-Output "Check_JSON_Notes_Expired"
        $ListNoteFilesEX = (Get-Childitem "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\").BaseName
        foreach($AlrtDefEX in $Global:ExpireKeysCertsAlready) {

        
            $CompareNamesEX = Compare-Object -ReferenceObject $AlrtDefEX.DisplayName -DifferenceObject $ListNoteFilesEX -IncludeEqual
                foreach($CompNameEX in $CompareNamesEX) {
                    If($CompNameEX.SideIndicator -eq "==") {
                        $SPFinNameEX = $CompNameEX.InputObject
                        $SPFinNameInfoEX = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPFinNameEX.json" | ConvertFrom-Json -Verbose
                        $SPNotesEX = $SPFinNameInfoEX.SPNotes
                            $AlrtDefEX | Add-Member -NotePropertyName Notes -NotePropertyValue $SPNotesEX -Force
                            $Global:ExpireKeysCertsAlready += @($AlrtDefEX)
                    }
              }
       }
       Write-Output $Error
       Expire_On_Report
}

Function Expire_On_Report {
    $Error.Clear()
    Write-Output "Made it to the decision tree for report Type"
    if ($Global:XcludeExpired -eq 'YES') {
        Write-output "was true - only running the primary Report"
            Write-Output $Error
            Primary_Report
            } else {
            Write-Output "was false - running primary Report WITH Expired"
            Write-Output $Error
            Primary_Report_With_Expired
    }
}
Function Primary_Report {
    Write-Output "running primary report"
    
    $KeysThatExpireSoon = $Global:ExpireKeysCertsSoon | Sort-Object DisplayName -Unique
    $KeysThatExpireSoon | Select-Object DisplayName,DaysToExpire,StartDate,ExpiryDate,KeyID,Notes,Internal,External
  
  foreach($keys in $KeysThatExpireSoon){
$DN = $Keys.DisplayName
$DTE = $Keys.DaysToExpire
$SD = $Keys.StartDate
$EXD = $Keys.ExpiryDate
$KID = $Keys.KeyID
$NOTES = $Keys.Notes
$IntOwn = $Keys.Internal
$ExtOwn = $Keys.External

$dataRow = "
</tr>
<td>$DN</td>
<td><center>$DTE</center></td>
<td><center>$SD</center></td>
<td><center>$EXD</center></td>
<td>$KID</td>
<td>$NOTES</td>
<td>$IntOwn</td>
<td>$ExtOwn</td>
</tr>
"

$dataforreport += $datarow

}           
          
$Global:report = "<html>
<style>
{font-family: Arial; font-size: 13pt;}
TABLE{border: 1px solid black; border-collapse: collapse; font-size:10pt;}
TH{border: 1px solid black; background: #71aac7; padding: 5px; color: #000000;}
TD{border: 1px solid black; padding: 5px; }
</style>
<h3>$Global:updateMessage</h3>
<h2>Azure Service Principal Key/Certificate Expiry Report</h2>
<table>
<tr>
<th>Display Name</th>
<th>Days To Expire</th>
<th>Start Date</th>
<th>Expiry Date</th>
<th>Key ID</th>
<th>Notes</th>
<th>Internal Owners</th>
<th>External Owners</th>
</tr>
$dataforreport
</table>
"

Mail_Handler_Primary
} 
   

Function Primary_Report_With_Expired{
    Write-Output "Running primary with expired"

    $KeysThatExpireSoon = $Global:ExpireKeysCertsSoon | Sort-Object DisplayName -Unique
    $KeysThatExpireSoon | Select-Object DisplayName,DaysToExpire,StartDate,ExpiryDate,KeyID,Notes,Internal,External

    $KeysThatExpired = $Global:ExpireKeysCertsAlready | Sort-Object ExpiredDaysAgo -Unique
    $KeysThatExpired | Select-Object DisplayName,ExpiredDaysAgo,ExpiryDate,KeyID,Notes,Internal,External
  
  foreach($keys in $KeysThatExpireSoon)
{
$DN = $Keys.DisplayName
$DTE = $Keys.DaysToExpire
$SD = $Keys.StartDate
$EXD = $Keys.ExpiryDate
$KID = $Keys.KeyID
$NOTES = $Keys.Notes
$IntOwn = $Keys.Internal
$ExtOwn = $Keys.External

$dataRow = "
</tr>
<td>$DN</td>
<td><center>$DTE</center></td>
<td><center>$SD</center></td>
<td><center>$EXD</center></td>
<td>$KID</td>
<td>$NOTES</td>
<td>$IntOwn</td>
<td>$ExtOwn</td>
</tr>
"

$dataforreport += $datarow

}      
   
foreach($keysEX in $KeysThatExpired)
{
$DNX = $KeysEX.DisplayName
$XDAX = $KeysEX.ExpiredDaysAgo
$EXDX = $KeysEX.ExpiryDate
$KIDX = $KeysEX.KeyID
$NOTESX = $KeysEX.Notes
$IntOwn = $Keys.Internal
$ExtOwn = $Keys.External

$dataRowEX = "
</tr>
<td>$DNX</td>
<td><center>$XDAX</center></td>
<td><center>$EXDX</center></td>
<td>$KIDX</td>
<td>$NOTESX</td>
<td>$IntOwn</td>
<td>$ExtOwn</td>
</tr>
"

$dataforreportEX += $datarowEX

}      
   
        
$Global:report = "<html>
<style>
{font-family: Arial; font-size: 13pt;}
TABLE{border: 1px solid black; border-collapse: collapse; font-size:10pt;}
TH{border: 1px solid black; background: #71aac7; padding: 5px; color: #000000;}
TD{border: 1px solid black; padding: 5px; }
</style>
<h3>$Global:updateMessage</h3>
<h2>Azure Service Principal Key/Certificate Expiry Report</h2>
<table>
<tr>
<th>Display Name</th>
<th>Days To Expire</th>
<th>Start Date</th>
<th>Expiry Date</th>
<th>Key ID</th>
<th>Notes</th>
<th>Internal Owners</th>
<th>External Owners</th>
</tr>
$dataforreport
</table>

<h2>Azure Service Principals with already Expired Keys/Certificates</h2>
<table>
<tr>
<th>Display Name</th>
<th>Expired Days Ago</th>
<th>Expiry Date</th>
<th>Key ID</th>
<th>Notes</th>
<th>Internal Owners</th>
<th>External Owners</th>
</tr>
$dataforreportEX
</table>
<tr>  
" 

Mail_Handler_Primary
}   
        
        
Function Mail_Handler_Primary { 
        $Error.Clear()

            $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
                $encPass = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\emailCred.txt" | ConvertTo-SecureString -Key $ASPCKeyData

                $encmailuser = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\emailFrom.txt" | ConvertTo-SecureString -Key $ASPCKeyData
                $Plainmailuser = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($encmailuser))

                $credential = New-Object System.Management.Automation.PSCredential ($Plainmailuser,$encPass)
                
                $Global:DefaultEmail = $Global:DefaultEmail.Split(",")
        
                
                foreach ($DefEm in $Global:DefaultEmail) {
        
                $mailParams =@{
                    SmtpServer          = 'smtp.office365.com'
                    UseSsl              = $true
                    Credential          = $credential
                    From                = $Plainmailuser
                    To                  = $DefEm
                    Subject             = $Global:ReportName
                    Body                = $Global:report
        }
    Send-MailMessage @mailParams -BodyAsHtml
    }        
    Write-Output $Error
    EngineCleanup
}


Function Mail_Handler_Secondary {
          $Error.Clear()

            $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
                $encPass = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\emailCred.txt" | ConvertTo-SecureString -Key $ASPCKeyData

                $encmailuser = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Creds\emailFrom.txt" | ConvertTo-SecureString -Key $ASPCKeyData
                $Plainmailuser = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($encmailuser))

                $credential = New-Object System.Management.Automation.PSCredential ($Plainmailuser,$encPass)
        
                $Global:ExternalMailAddr = $Global:ExternalMailAddr.Split(",")
                
                foreach ($DefEmX in $Global:ExternalMailAddr) {    
					if (!$DefEmX) {
						Write-output "No external email address"
					} else {
					
                    $mailParamsXmail =@{
                        SmtpServer          = 'smtp.office365.com'
                        UseSsl              = $true
                        Credential          = $credential
                        From                = $Plainmailuser
                        To                  = $DefEmX
                        Subject             = 'CSS - Expiring Service Principal Information'
                        Body                = $Global:ExternalMailNote
                    }
        
                Send-MailMessage @mailParamsXmail
                Write-Output $Error
			}
        }
}

Function EngineCleanup {
    Write-Host "Engine Cleanup Routine Started"
    Write-Host "Deleteing logs old logs.."
    
    Write-Host "Deleting any AlertDef Logs older than 90 days..."
    Get-ChildItem "C:\Program Files\Enigma-Tek\ASPC\Logs\AlertDefs"| where-Object LastWriteTime -LT (Get-Date).AddDays(-90) | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    
    Write-Host "Deleting any Engine Logs older than 60 days..."
    Get-ChildItem "C:\Program Files\Enigma-Tek\ASPC\Logs\Engine" | where-Object LastWriteTime -LT (Get-Date).AddDays(-60) | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    
    Write-Host "Deleting any Update Logs older than 180 days..."
    Get-ChildItem "C:\Program Files\Enigma-Tek\ASPC\Logs\Updater" | where-Object LastWriteTime -LT (Get-Date).AddDays(-180) | Remove-Item -Recurse -Force -Confirm:$false -Verbose

    EngineStop
}

Function EngineStop {
    $ScriptStopTime = Get-Date -Format HH.mm.ss
    Write-Output "Stopping Engine and Logging at $ScriptStopTime"
    Stop-Transcript
    exit
}

#Script

#Start Engine Transcription / log
    $TranscriptDateTime = Get-Date -Format HH.mm-MM.dd.yyyy
    Start-Transcript -Path "C:\Program Files\Enigma-Tek\ASPC\Logs\Engine\ASPC_Engine_RunLog_$TranscriptDateTime.log" -Verbose
    
    $ScriptStartTime = Get-Date -Format HH.mm.ss
    Write-Output "ASPC Engine Started at $ScriptStartTime"

    Write-Output "Calling first function: updateChecker"
    updateChecker
