

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="600" Title="(ASPC) Azure Service Principal Checker Installer/Uninstaller" Height="400" Name="Window_Main" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
 <Window.Icon>
        <DrawingImage/>
    </Window.Icon>
<Grid Margin="0,2,0,-2">
 




<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Welcome}" Margin="10,10,0,10" Width="575" Height="70" Name="Welcome_Text_Block" Background="#4a90e2" Foreground="#ffffff" Padding="0,10,0,0"/>

<Button Content="Install" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="16,153,0,0" Foreground="#194d4a" BorderBrush="#194d4a" Height="30" FontSize="15" FontWeight="DemiBold" Background="#ffffff" Name="Install_Button"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="66" Width="578" TextWrapping="Wrap" Margin="8.5,286.984375,0,0" Foreground="{Binding Output_textColor}" Text="{Binding Output}" Name="lku8k2wt4toca"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Output" Margin="10,268,0,0"/>

<Button Content="Uninstall" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="490,240,0,0" Foreground="#194d4a" BorderBrush="#194d4a" Height="30" FontSize="15" FontWeight="DemiBold" Background="#ffffff" Name="Uninstal_BTN"/>
<CheckBox HorizontalAlignment="Left" VerticalAlignment="Top" Content="I understand that I am uninstalling ASPC" Margin="299,206,0,0" Name="uninstallAgree_chkBox"/>
<CheckBox HorizontalAlignment="Left" VerticalAlignment="Top" Content="Keep your Configuration files" Margin="299,239,0,0" Name="keepConfigs_chkBox"/>
 


<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="2" Width="555" Margin="18.8438,196.391,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Install - This will setup the File Structure, download all relevant software and setup the shortcut for the Configurator tool. The shortcuts will be set to &quot;Run as Admin&quot;. All docs are on the Github Wiki. Once the install is complete, close the installer and then Click the shortcut on your desktop to get started." Margin="8,89,0,0" Name="Install_txt" Width="585" Height="66" FontSize="11"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Uninstall - Click the box to agree to uninstall. If you would like to keep your config files, check that box(Config Files will be archived in case of re-install)." Margin="18.5,206,0,0" Name="Uninstall_txt" Width="273" Height="63" FontSize="11"/>
</Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#region Welcome_Message
function welcome_message {
    Async {
        $welcomeMessage = "   Welcome to the ASPC Installer. Please read the EULA on the Github Wiki,` 
        Installing is an automatic agreement to the EULA posted there.`
        Installer must be run as Administrator. This installer is used to Install or Uninstall ASPC"
        $State.Welcome = $welcomeMessage
    }
}
#endregion 
#region Buttons
function install_button {
    Async {
        $outputText = "Installation will start shortly..."
        $State.Output = $outputText
        softInstall
    }
}

function uninstall_button {
        Async {
        $outputText = "Uninstall will start shortly..."
        $State.Output_textColor = "#000000"
        $State.Output = $outputText
        uninstallASPCChk
    }
}
#endregion 
#region Checkboxs

function uninstall_Checked {
      Async {
       $State.uninstallAgree = "1"
    }
}

function uninstall_Unchecked {
      Async {
       $State.uninstallAgree = "0"
    }
}

function keepConfigs_Checked {
      Async {
       $State.keepConfigFiles = "1"
    }
}

function keepConfigs_Unchecked {
      Async {
       $State.keepConfigFiles = "0"
    }
}
#endregion 
#region Installer
function softInstall {
        Start-Sleep -Seconds 2
        
        $installTest = test-path "C:\Program Files\Enigma-Tek\ASPC\Configs\Install\install.log" -PathType Leaf
        if ($installTest -eq $True) {
            $outputText = "It looks like we found an install file. ASPC may be installed. Please see if there is an Enigma-Tek folder under Program Files (C:\Program Files\Enigma-Tek\ASPC) and see what's inside. If you want to make sure its a clean install, re-name the ASPC folder to ASPC.old and then re-run the installer."
            $State.Output_textColor = "#D0021B"
            $State.Output = $outputText
        } else {
            $State.Output_textColor = "#000000"
            $installTimeStart = Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"
            create_folders
        }
}

function create_folders {
        $outputText = "Creating Folder structure."
        $State.Output = $outputText
        $Error.Clear()
           	New-Item -Path "C:\Program Files\" -Name "Enigma-Tek" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\" -Name "ASPC" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\" -Name "Configs" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\" -Name "AlertDefs" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\" -Name "Engine" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\" -Name "Install" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\" -Name "Configurator" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\" -Name "Version" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\" -Name "Creds" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\" -Name "Software" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\" -Name "Logs" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Logs\" -Name "AlertDefs" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Logs\" -Name "Engine" -ItemType "directory"
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Logs\" -Name "Updater" -ItemType "directory"
        $folderCreateErrors = Write-output $Error
            
        $outputText = $outputText + "`nFolder structure created."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        download_files
}
#Download the Engine, Configurator Updater and Reports
function download_files {
        $outputText = "Downloading Engine and Configurator Files."
        $State.Output = $outputText
        $Error.Clear()
            if (Test-Path -Path "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Engine.exe" -PathType Leaf){

            } else {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/ASPC_Engine.ps1" -OutFile "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Engine.ps1"
            }
            
             if (Test-Path -Path "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Configurator.exe" -PathType Leaf){

            } else {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/ASPC_Configurator.ps1" -OutFile "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Configurator.ps1"
            }
            
            if (Test-Path -Path "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Updater.ps1" -PathType Leaf){

            } else {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/ASPC_Updater.ps1" -OutFile "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Updater.exe"
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/versions.json" -OutFile "C:\Program Files\Enigma-Tek\ASPC\Configs\Version\versions.json"
            }
            
            
        $downloadFilesErrors = Write-output $Error
        $outputText = $outputText + "`nEngine, Configurator and Updater Files downloaded."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        write_shortcuts
}

function getVersions {
    
    
}
#Creates a desktop shortcut for all users. Sets the icon and sets the shortcut to "Run as Administrator".
function write_shortcuts {
        $Error.Clear()
        $outputText = "Creating Desktop Shortcuts."
        $State.Output = $outputText
         $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\ASPC_Configurator.lnk")
        $Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.Arguments = "-noexit -ExecutionPolicy Bypass -File ""C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Configurator.ps1"""
        $Shortcut.RelativePath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $shortcut.IconLocation = "shell32.dll,21"
        $Shortcut.Save()
        
        $bytes = [System.IO.File]::ReadAllBytes("C:\Users\Public\Desktop\ASPC_Configurator.lnk")
        $bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
        [System.IO.File]::WriteAllBytes("C:\Users\Public\Desktop\ASPC_Configurator.lnk", $bytes)
        
        $shortcutsCreated = Write-output $Error
        $outputText = $outputText + "`nShortcuts Created."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        write_default_eng_configs
}

#Writes the default engine config file
function write_default_eng_configs {
        $Error.Clear()
        $outputText = "Writing default Engine config file."
        $State.Output = $outputText
            $ExpireWithin = "30"
            $XcludeExpired = "NO"
            $DefaultEmail = "distro@yourcompany.com"
            $ReportName = "ASPC Key / Certificate Expiring Report"

        $engConfig = @{
        ExpireWithin = $ExpireWithin
        XcludeExpired = $XcludeExpired
        DefaultEmail = $DefaultEmail
        ReportName = $ReportName
        
        } 
        $engConfig | ConvertTo-Json | Set-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\Eng_Config.json"
        $outputText = $outputText + "`nEngine config file written."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        create_enc_key
}

function create_enc_key {
        $EncryptionKeyBytes = New-Object Byte[] 32
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($EncryptionKeyBytes)
        $EncryptionKeyBytes | Out-File "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
        write_install_log
}

#Writes information to a log file
function write_install_log {
        $outputText = "Creating install log file."
        $installTimeStop = Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"
        $UserName = $env:UserName
        $ComputerName = $env:computername
            
            $InstallLogFile = "C:\Program Files\Enigma-Tek\ASPC\Configs\Install\install.log"
            $InstallLogOutput = "","","Install Date and Time",$installTimeStart,"","Install Stop Time",$installTimeStop,"","Computer Name:",$ComputerName,"","Installed By:",$UserName,"","Folder_Creation_Errors",$folderCreateErrors,"","Download_File_Errors",$downloadFilesErrors,"","Create_Shortcut_Errors",$shortcutsCreated,
            $InstallLogOutput  | Out-File -FilePath $InstallLogFile -append
        $outputText = $outputText + "`nInstall log file written."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        $outputText = "Install complete." 
        $outputText = $outputText + "`nPlease close the Installer with the X in the upper right hand corner. You will find the Configurator Shortcut on your desktop. Click that to get started."
        $State.Output = $outputText
}
#endregion 
#region Uninstaller

function uninstallASPCChk {
   Start-Sleep -Seconds 2
   
        $uninstallBoxCheck = $State.uninstallAgree
        If ($uninstallBoxCheck -eq "0") {
            $outputText = "Please check the box that you understand that you are removing all ASPC software, shortcuts, configs, etc(unless keeping configs, check that box then as well)"
            $State.Output_textColor = "#D0021B"
            $State.Output = $outputText
            
        } elseif ($uninstallBoxCheck -eq "1") {
            $outputText = "Starting Uninstall."
            $State.Output_textColor = "#000000"
            $State.Output = $outputText
            Start-Sleep -Seconds 2
            StartUninstall
        }
}

function StartUninstall {
        $outputText = $outputText + "`nChecking to if you want to keep configs."
        $State.Output = $outputText
            Start-Sleep -Seconds 3
            $KeepConfigs = $State.keepConfigFiles
            
        If ($KeepConfigs -eq "0") {
            $outputText = $outputText + "`nLooks like you dont want to keep anything."
            $State.Output = $outputText
            Start-Sleep -Seconds 3
            
            $outputText = "Deleting all Files and Folders."
            $State.Output = $outputText
                $RemovePath = "C:\Program Files\Enigma-Tek\ASPC"
                Get-ChildItem -Path $RemovePath -Recurse | Remove-Item -force -recurse
                Remove-Item $RemovePath -Force 
                Start-Sleep -Seconds 3
            
            $outputText = $outputText + "`nDeleting shortcuts."
            $State.Output = $outputText
                $LinkDelete = "C:\Users\Public\Desktop\ASPC_Configurator.lnk"
                Remove-Item $LinkDelete -Force
                Unregister-ScheduledTask -TaskName ASPC_Engine_Run -Confirm:$false
                Start-Sleep -Seconds 3
             $outputText = "Uninstaller Complete. ASPC has been removed"
             $outputText = $outputText + "`nPlease close the Installer with the X in the upper right hand corner."
             $State.Output = $outputText
            
        } elseif ($KeepConfigs -eq "1") {
            $outputText = "Keep Configurations checked."
            $State.Output = $outputText
                Start-Sleep -Seconds 3
                
            $outputText = "Copying Config Files to folder C:\Program Files\Enigma-Tek\ASPC\ConfigBackup"
            $State.Output = $outputText
            
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\" -Name "ConfigBackup" -ItemType "directory"
            Copy-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs" "C:\Program Files\Enigma-Tek\ASPC\ConfigBackup" -Recurse
            Copy-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\Configurator" "C:\Program Files\Enigma-Tek\ASPC\ConfigBackup" -Recurse
            New-Item -Path "C:\Program Files\Enigma-Tek\ASPC\ConfigBackup" -Name "Engine" -ItemType "directory"
            Copy-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\Eng_Config.json" "C:\Program Files\Enigma-Tek\ASPC\ConfigBackup\Engine" -Recurse
            
             $outputText = "Configuration files backed up." 
             $State.Output += $outputText
                Start-Sleep -Seconds 3
                
             $outputText = "Removing Executables, Folders and Logs." 
             $State.Output = $outputText
                Start-Sleep -Seconds 3
                
            Remove-Item "C:\Program Files\Enigma-Tek\ASPC\Configs" -Recurse -Force
            Remove-Item "C:\Program Files\Enigma-Tek\ASPC\Creds" -Recurse -Force
            Remove-Item "C:\Program Files\Enigma-Tek\ASPC\Logs" -Recurse -Force
            Remove-Item "C:\Program Files\Enigma-Tek\ASPC\Software" -Recurse -Force
            
                   $outputText = $outputText + "`nDeleting shortcuts."
            $State.Output = $outputText
                $LinkDelete = "C:\Users\Public\Desktop\ASPC_Configurator.lnk"
                Remove-Item $LinkDelete -Force
                
                Unregister-ScheduledTask -TaskName ASPC_Engine_Run -Confirm:$false
                Start-Sleep -Seconds 3
             $outputText = "Uninstaller Complete. ASPC has been removed and your Configuration files were saved."
             $outputText = $outputText + "`nPlease close the Installer with the X in the upper right hand corner."
             $State.Output = $outputText
        }
}
#endregion 


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$Welcome_Text_Block.Add_Loaded({welcome_message $this $_})
$Install_Button.Add_Click({install_button $this $_})
$Uninstal_BTN.Add_Click({uninstall_button $this $_})
$uninstallAgree_chkBox.Add_Checked({uninstall_Checked $this $_})
$uninstallAgree_chkBox.Add_Unchecked({uninstall_Unchecked $this $_})
$keepConfigs_chkBox.Add_Checked({keepConfigs_Checked $this $_})
$keepConfigs_chkBox.Add_Unchecked({keepConfigs_Unchecked $this $_})

$State = [PSCustomObject]@{}


Function Set-Binding {
    Param($Target,$Property,$Index,$Name,$UpdateSourceTrigger)
 
    $Binding = New-Object System.Windows.Data.Binding
    $Binding.Path = "["+$Index+"]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    if($UpdateSourceTrigger -ne $null){$Binding.UpdateSourceTrigger = $UpdateSourceTrigger}


    [void]$Target.SetBinding($Property,$Binding)
}

function FillDataContext($props){

    For ($i=0; $i -lt $props.Length; $i++) {
   
   $prop = $props[$i]
   $DataContext.Add($DataObject."$prop")
   
    $getter = [scriptblock]::Create("Write-Output `$DataContext['$i'] -noenumerate")
    $setter = [scriptblock]::Create("param(`$val) return `$DataContext['$i']=`$val")
    $State | Add-Member -Name $prop -MemberType ScriptProperty -Value  $getter -SecondValue $setter
               
       }
   }



$DataObject =  ConvertFrom-Json @"

{
    "Welcome" : "",
    "Output" : "",
    "Output_textColor" : "",
    "uninstallAgree" : "0",
    "keepConfigFiles" : "0"
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("Welcome","Output","Output_textColor","uninstallAgree","keepConfigFiles") 

$Window.DataContext = $DataContext
Set-Binding -Target $Welcome_Text_Block -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 0 -Name "Welcome"  
Set-Binding -Target $lku8k2wt4toca -Property $([System.Windows.Controls.TextBox]::ForegroundProperty) -Index 2 -Name "Output_textColor"  
Set-Binding -Target $lku8k2wt4toca -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 1 -Name "Output"  




$Global:SyncHash = [HashTable]::Synchronized(@{})
$SyncHash.Window = $Window
$Jobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$initialSessionState = [initialsessionstate]::CreateDefault()

Function Start-RunspaceTask
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,Position=0)][ScriptBlock]$ScriptBlock,
          [Parameter(Mandatory=$True,Position=1)][PSObject[]]$ProxyVars)
            
    $Runspace = [RunspaceFactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions  = 'ReuseThread'
    $Runspace.Open()
    ForEach($Var in $ProxyVars){$Runspace.SessionStateProxy.SetVariable($Var.Name, $Var.Variable)}
    $Thread = [PowerShell]::Create('NewRunspace')
    $Thread.AddScript($ScriptBlock) | Out-Null
    $Thread.Runspace = $Runspace
    [Void]$Jobs.Add([PSObject]@{ PowerShell = $Thread ; Runspace = $Thread.BeginInvoke() })
}

$JobCleanupScript = {
    Do
    {    
        ForEach($Job in $Jobs)
        {            
            If($Job.Runspace.IsCompleted)
            {
                [Void]$Job.Powershell.EndInvoke($Job.Runspace)
                $Job.PowerShell.Runspace.Close()
                $Job.PowerShell.Runspace.Dispose()
                $Job.Powershell.Dispose()
                
                $Jobs.Remove($Job)
            }
        }

        Start-Sleep -Seconds 1
    }
    While ($SyncHash.CleanupJobs)
}

Get-ChildItem Function: | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       
    $Definition = Get-Content "function:$_" -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
    $InitialSessionState.Commands.Add($SessionStateFunction)
}


$Window.Add_Closed({
    Write-Verbose 'Halt runspace cleanup job processing'
    $SyncHash.CleanupJobs = $False
})

$SyncHash.CleanupJobs = $True
function Async($scriptBlock){ Start-RunspaceTask $scriptBlock @([PSObject]@{ Name='DataContext' ; Variable=$DataContext},[PSObject]@{Name="State"; Variable=$State},[PSObject]@{Name = "SyncHash";Variable = $SyncHash})}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name='Jobs' ; Variable=$Jobs },[PSObject]@{Name = "SyncHash";Variable = $SyncHash})



$Window.ShowDialog()
