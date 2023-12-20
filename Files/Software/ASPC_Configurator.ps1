$Global:PFXPassword
$Global:EngineTaskCreated = ""
$Global:EngineTaskUpdated = ""

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
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="900" Title="(ASPC) Azure Service Principal Checker Configurator " Height="600" Name="Window_Main" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
 <Window.Icon>
        <DrawingImage/>
    </Window.Icon>

<Grid Margin="0,-1,0,1" Name="ParentGrid">
<Grid.RowDefinitions>
<RowDefinition Height="58*"/>
<RowDefinition Height="503*"/>
</Grid.RowDefinitions>
<Border BorderBrush="Black" BorderThickness="1" Grid.Row="0" Grid.Column="0" Name="ParentBoarder">
<StackPanel Name="ParentStackPanel" Background="#384e8d" Orientation="Horizontal">
<Button Content="Home" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="HomeBTN"/>
<Button Content="Alert Defs" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="ArtDefBtn"/>

<Button Content="Eng Config" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="EngConfBtn"/>
<Button Content="Credentials" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="CredsConfBtn"/>
<Button Content="Sch Tasks" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="TasksBtn"/>
<Button Content="Update" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="8" FontStyle="Italic" Name="UpdateBtn"/>

</StackPanel>
</Border>
<TabControl Grid.Row="1" Grid.Column="0" Name="ParentTabControl" SelectedIndex="{Binding tabIndex}" Margin="0,1,0,-1"><TabItem Header="Home" Visibility="Collapsed"><Grid Background="#ffffff" Name="HomeGrid" Margin="0,2,0,-2">



<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="2" Width="866" Margin="5.837000000000003,374.188,0,0"/>





<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="67" VerticalAlignment="Top" Width="879" Margin="1.837493896484375,2.9875030517578125,0,0" Background="#717fa8">
<TextBlock TextWrapping="Wrap" Text="Welcome to the ASPC Configurator Home Page. This configuration tool is used to setup the ASPC Engine, setup the credentials to be used, and update the ASPC software. Below are the defintions of each top button and what's in that section." Name="HomeTxt" FontSize="14" FontWeight="SemiBold" Foreground="#ffffff" Margin="43,10,38,5"/>
</Border>
<StackPanel HorizontalAlignment="Left" Height="197" VerticalAlignment="Top" Width="866" Margin="5,110,0,0">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="~Alert Definitions: Section for setting up different Alert settings for each Service Principal Key/Cert that may be scanned. Allows you to to add notes (where the key goes, team it may need to go to, etc) and other emails to the alerts per Service Principal." Width="874" Height="46" Margin="4,6,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="~Engine Configuration: Set defaults for the Engine. Includes default email, what is included and excluded from daily runs, days ahead to scan, etc." Width="872" Height="27"/>

<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="~Credentials: Creates, encrypts and saves the credentials for: The SP to connect to Azure to scan the Service Principals for expiring Keys/Certs. The Email account for the report. The local/domain account to run the task." Height="44" Width="869"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="~Tasks: Section to setup the scheduled task(s) to run the Engine." Height="27" Width="875"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="~Update: Section describes and runs the software Updates if available." Height="24" Width="382"/>
</StackPanel>
<Button Content="About" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Margin="757,327,0,0" Name="AboutBTN"/>

</Grid>
</TabItem><TabItem Header="AlrtDef" Visibility="Collapsed"><Grid Background="#ffffff" Margin="0,1,0,-1" Name="AlrtDefGrid">

<Border BorderBrush="#000000" BorderThickness="1" HorizontalAlignment="Left" Height="66" VerticalAlignment="Top" Width="882" Margin="0.5,0.03125,0,0" Background="#717fa8">
<TextBlock TextWrapping="Wrap" Text="Alert Definition Page. Press the load button to load an existing definition and view or change (REMEMBER TO SAVE CHANGES!!). Clear Fields button clears the fields and the Delete Definitions button to view and delete any existing defintions. To add a new defintion: Fill in the blank fields and click save. If you opened a Definition, click the Clear Fields button and then proceed to fill in the blank fields" Foreground="#ffffff" Margin="25,4,24,5"/>
</Border>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="27" Width="534" TextWrapping="Wrap" Margin="9,145,0,0" Text="{Binding SPName}" Padding="2,6,0,0" BorderBrush="#4a90e2" Name="lqdz6538i3o5f"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="59" Width="865" TextWrapping="Wrap" Margin="10,200,0,0" Text="{Binding SpNote}" BorderBrush="#4a90e2" Name="lqdz6538zvwox"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="27" Width="532" TextWrapping="Wrap" Margin="10,337,0,0" Text="{Binding AlrtDefXMail}" Padding="2,6,0,0" BorderBrush="#4a90e2" Name="lqdz6538a81wr"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="65" Width="866" TextWrapping="Wrap" Margin="10,390,0,0" Text="{Binding AlrtDefXNote}" BorderBrush="#4a90e2" Name="lqdz6538u0xsb"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Service Principal Name. This is the name as seen in the Azure portal and must match exactly." Margin="10,125,0,0" Width="522" Height="22" FontStyle="Italic" FontWeight="SemiBold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Internal Note. This should be any information for internal teams about the Key/Certificate that is expiring." Margin="10,181,0,0" Width="586" Height="23" FontStyle="Italic" FontWeight="SemiBold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="External Email Address. Alert external parties. Must be comma seperated addresses if more than one. (first@company.com,second@company.com)" Margin="11,317,0,0" Width="863" Height="15" FontWeight="SemiBold" FontStyle="Italic"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="External Note. A message to external parties. DO NOT include any sensitive information here!" Margin="11,372,0,0" Width="507" Height="20" FontStyle="Italic" FontWeight="SemiBold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Alrt_Def_Output}" Margin="600,80,0,0" Width="227" Height="103" Background="#eeebeb" Name="lqdz653839nyj"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Output" Margin="560,73,0,0" Width="62" Height="18"/>
<Button Content="Delete Def" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="365,86,0,0" FontStyle="Italic" Name="Delete_Def_Button"/>
<Button Content="Clear Fields" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="247,86,0,0" FontStyle="Italic" Name="lqdz6538596bc"/>
<Button Content="Save Def" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="129,86,0,0" FontStyle="Italic" Name="lqdz6538607uy"/>
<Button Content="Load Alert Def" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="10,86,0,0" FontStyle="Italic" Name="lqdz6538y1xjk"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Internal Owners - Department" Margin="12,262,0,0" FontStyle="Italic" FontWeight="SemiBold"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="External Owners - Company and Email address" Margin="421,263,0,0" FontStyle="Italic" FontWeight="SemiBold"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="32" Width="395" TextWrapping="Wrap" Margin="12,280,0,0" BorderBrush="#4a90e2" Text="{Binding IntOwnerNote}" Name="lqdz6538ea67r"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="29" Width="453" TextWrapping="Wrap" Margin="420,281,0,0" BorderBrush="#4a90e2" Text="{Binding ExtOwnerNote}" Name="lqdz6538hsdz1"/>
</Grid>
</TabItem><TabItem Header="Extra" Visibility="Collapsed"><Grid Background="#ffffff" Name="ExtraGrid">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Extra" Margin="37.5,18.984375,0,0"/>
</Grid></TabItem><TabItem Header="EngConf" Visibility="Collapsed"><Grid Background="#ffffff" Name="EngConfGrid">

<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="48" VerticalAlignment="Top" Width="884" Margin="-0.15625,-1.015625,0,0" Name="EningConf_Boarder" Background="#717fa8">
<TextBlock TextWrapping="Wrap" Text="Engine Configuration page allows you to set the defaults used when the engine runs. The current configuration will load when the page loads. After making any changes to the configuration, you must hit the SAVE button else your changes will not save." Name="EngConf_Txt" Foreground="#ffffff" Width="841" Height="39" Margin="26,5,17,4" HorizontalAlignment="Center" VerticalAlignment="Center"/>
</Border>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="25" Width="489" TextWrapping="Wrap" Margin="25,114,0,0" Text="{Binding EngConfDefaultEmail}" Padding="2,2,0,0" BorderBrush="#4a90e2" Name="lqdz6538zeufi"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="27" Width="50" TextWrapping="Wrap" Margin="25,302,0,0" Text="{Binding EngConfExpireWithin}" Padding="15,5,0,0" BorderBrush="#4a90e2" Name="lqdz6538vh6tn"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="32" Width="490" TextWrapping="Wrap" Margin="25,209,0,0" Text="{Binding EngConfReportName}" Padding="2,7,0,0" BorderBrush="#4a90e2" Name="lqdz6538wq570"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Primary report email address. You can add more than one address here for your daily report, addresses must be comma seperated. (Eg. add1@company.com,add2@company.com)" Margin="27,72,0,0" Width="536" Height="41"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Report Name. This is the subject name on your daily report email. Default is ASPC Key / Certificate Expiring Report." Margin="27,166,0,0" Width="485" Height="39"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Numer of days out to Expiration. This is the number of days out from today that it looks to add to the report. This number has to be between 1-90. Default is 30." Margin="28,261,0,0" Width="486" Height="38"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="The default report includes a seperate section with expired Keys/Certs. You can EXCLUDE this section from the report by checking the box below." Margin="26,354,0,0" Width="491" Height="39"/>

<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="393" Width="2" Margin="530.5,56.98400000000001,0,0" OpacityMask="#000000"/>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="2" Width="343" Margin="536.5,295.98400000000004,0,0" OpacityMask="#000000"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="YOU MUST CLICK SAVE!! If you dont click the save button, any changes made will be removed if you navigate off this page." Margin="540.5,76.984375,0,0" Width="339" Height="45"/>
<Button Content="SAVE Eng Conf" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" Margin="642,145,0,0" FontStyle="Italic" Name="eng_conf_save_button"/>
<CheckBox HorizontalAlignment="Left" VerticalAlignment="Top" Content="Exclude expired (checked = YES / un-checked = NO)" Margin="25.5,398.46875,0,0" IsChecked="{Binding ExcludeExpiredCheckBox}" BorderBrush="#4a90e2" Name="lqdz6538eu0an"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Eng_Save_Output}" Margin="541.5,209.984375,0,0" Name="Eng_Conf_Save_output" Width="338" Height="74"/>
<Button Content="Test engine" HorizontalAlignment="Left" VerticalAlignment="Top" Width="185" Margin="614,407,0,0" Height="30" Foreground="#4a4a4a" FontSize="016" FontStyle="Italic" FontWeight="ExtraBold" Name="lqdz65381fl8d"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="If you have created Credentials and the Task then click the button below to run a test on the engine. If you dont recieve your email within a few minutes, then please check the Engine log files for errors." Margin="561.5,313.328125,0,0" Width="307" Height="84"/>
</Grid></TabItem><TabItem Header="Creds" Visibility="Collapsed"><Grid Background="#ffffff" Name="CredsGrid" Margin="-1,1,1,-1">





















<TabControl SelectedIndex="{Binding CredsTabIndex}" Name="lqdz6538ui0nq"><TabItem Header="Instructions"><Grid Background="#ffffff" Margin="0,2,0,-2">





<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="56" VerticalAlignment="Top" Width="857" Margin="4,3,0,0" Background="#717fa8" Name="Local_Domain_Boarder">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Below are the instructions on setting up the Service Principal and Email Accounts for the Engine to use when it runs. Please follow these and then move on to the Tasks page when tested and complete." Margin="36,8,0,0" Width="780" Height="39" Foreground="#ffffff"/>
</Border>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Setting up the credentials is the first step in getting ASPC up and running. Lets start with the Service Principal: The service principal is setup in Azure under Active Directory > App Registrations. There is full setup instructions in the Wiki for this projects and you should start there if you have not already. You will need the service principal Client ID, The Tenant ID and the certificate (both the Certificate thumbprint and the PFX with its pasword). Once you have added all the information and save it then move to the Email account tab." Margin="5.5,65.8125,0,0" Width="387" Height="333" Background="#edecf3"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Email account: For the email account, we currently only support sending mail using Office365 accounts. How you manage your accounts is up to you. We recommend an email account that is set to not use MFA/2fa. We understand the risk. These utilities are PowerShell based and MSFT has not updated it to use a better module than what has been around for a while. You will alos need to allow and create an App Password on the account. We highly recommend that you put an incredebly long passcode on the account. To set it up you will need the email address and its password. Everything else is autofilled to work with Office365." Margin="471,66,0,0" Width="387" Height="333" Background="#edecf3"/>
</Grid>
</TabItem><TabItem Header="Service Principal"><Grid Background="#ffffff">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Service Principal Application (client) ID" Margin="5,134,0,0" Width="262" Height="20"/>
<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="56" VerticalAlignment="Top" Width="869" Margin="4,3,0,0" Background="#717fa8">


<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Below is the input for the Service Principal / App Registration that you created along with the certificate that you uploaded to it. Fill in the fields below with the information that you copied out of the App Registration when you created it along with the Thumbprint of the certificate. Click the PFX button to point to the PFX file and then press the Get Pass button to input its password securely. Save when done." Margin="17,1,0,0" Width="840" Height="46" Foreground="#ffffff" FontSize="12"/>
</Border>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="23" Width="400" TextWrapping="Wrap" Margin="3,154,0,0" BorderBrush="#4a90e2" Text="{Binding SP_AppID}" Name="lqdz6538mgwwn"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="25" Width="401" TextWrapping="Wrap" Margin="2,219,0,0" BorderBrush="#4a90e2" Text="{Binding SP_TennantID}" Name="lqdz6538ncrbo"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="25" Width="397" TextWrapping="Wrap" Margin="3,282,0,0" BorderBrush="#4a90e2" Text="{Binding Cert_Thumbprint}" Name="lqdz65388pieb"/>
<Button Content="PFX" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="477,139,0,0" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6538yvw1t"/>

<Button Content="SAVE" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="442,364,0,0" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6538hjmpc"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Directory (tenant) ID" Margin="5,198,0,0" Width="129" Height="23"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Cert Thumbprint" Margin="5,263,0,0" Width="99" Height="20"/>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="273" Width="3" Margin="442.5,72.031,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Fill in the fields below with the information from the App registration you created in your Azure tenant." Margin="4.5,70.03125,0,0" Width="422" Height="51"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Upload the PFX version of the certificate that you used when setting up the App registration in your Azure tenant." Margin="453,70,0,0" Width="403" Height="54"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding PFX_Location}" Margin="477,178,0,0" Width="348" Height="58" Background="#e5e4eb" Name="lqdz6538tagm3"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="PFX Password" Margin="480,253,0,0" Width="99" Height="21"/>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="4" Width="886" Margin="-13.5,348.031,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Once the fields above are complete and the PFX has been choosen, click &quot;Save&quot;" Margin="11,361,0,0" Width="439" Height="34"/>



<Button Content="Get Pass" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="484.5,283.640625,0,0" Name="lqdz65386ksn8"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="57" Width="346" TextWrapping="Wrap" Margin="524.84375,359.578125,0,0" Text="{Binding SP_Page_Output}" Name="lqdz6538nvblg"/>
</Grid></TabItem><TabItem Header="Email account"><Grid Background="#ffffff" Margin="0,-1,0,1">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Email account (and from address)" Margin="73,77,0,0"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="20" Width="227" TextWrapping="Wrap" Margin="49,98,0,0" Text="{Binding EmailAddFrom}" Name="lqdz6538ewjc6"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Password - (push the get password button)" Margin="51,144,0,0" Width="241" Height="20"/>
<Button Content="Get Password" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="104,170,0,0" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6538625bd"/>
<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="56" VerticalAlignment="Top" Width="857" Margin="4,3,0,0" Background="#717fa8">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Email account test and save. Input the account you will use to send email from, click the get password button to input its password and then put in a test send to address and press the Test Account button. If you get an email, then go ahead and hit the save button to save your credentials." Margin="7,4,0,0" Width="840" Height="45" Foreground="#ffffff"/>
</Border>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="328" Width="12" Margin="357.844,70.73400000000001,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Currently this utility only supports sending mail through Office 365. We may look into adding other account types at a later date. You dont need to enter any of the SMTP info as we are autofilling it in. Please note again that this needs to be an APP Password on the account, not the login password and if you have conditional access policies or any 2fa/MFA policies they will need to be worked around (that is outside the scope of this utility)" Margin="385.84375,73.734375,0,0" Width="475" Height="119"/>
<Button Content="Test Account" HorizontalAlignment="Left" VerticalAlignment="Top" Width="109" Margin="106,278,0,0" Height="30" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6538dflgy"/>
<Button Content="Save Account" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="103,356,0,0" Height="28" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6538lly1e"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Put in a test Email address to send to and then hit Test Account" Margin="1,221,0,0" Width="346" Height="20"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="20" Width="227" TextWrapping="Wrap" Margin="51,242,0,0" Name="Test_Address" Text="{Binding TestEmailAdd}"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="If you received the Test email then click save below." Margin="29,331,0,0" Width="300" Height="18"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Output" Margin="386,211,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding EmailAccount_Output}" Margin="385.5,235.984375,0,0" Name="EmailAccount_Output" Width="477" Height="166" Background="#e5e4eb"/>
</Grid>
</TabItem></TabControl>
</Grid>
</TabItem><TabItem Header="Tasks" Visibility="Collapsed"><Grid Background="#ffffff" Name="TasksGrid" Margin="0,-1,0,1">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Scheduled Task Name: ASPC_Engine_Run (this cannot be altered)" Margin="60,82,0,0" Width="400" Height="24"/>
<ComboBox HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="56,178,0,0" ItemsSource="{Binding HourComboBox}" SelectedValue="{Binding SelectedHour}" Name="lqdz6539zcir2"/>
<ComboBox HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="174,178,0,0" SelectedValue="{Binding SelectedTOD}" ItemsSource="{Binding HourAMPM}" Name="lqdz6539laahl"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Choose Time of Day, AM or PM and &quot;Daily&quot;, &quot;Work Week&quot; (Mon-Fri) or a &quot;specific day of the week&quot;" Margin="19,107,0,0" Width="479" Height="30"/>
<ComboBox HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="292,178,0,0" ItemsSource="{Binding RunDay}" SelectedValue="{Binding SelectedRunDay}" Name="lqdz6539f2zzy"/>
<Button Content="Create / Update Task" HorizontalAlignment="Left" VerticalAlignment="Top" Width="168" Margin="177,290,0,0" Height="25" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff" FontStyle="Italic" Name="lqdz6539fetvn"/>
<Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="56" VerticalAlignment="Top" Width="877" Margin="4,3,0,0" Background="#717fa8">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text=" Scheduled Task. Choose a Time, AM or PM and which days you would like to recieve your report. Daily is everyday, Work week is Mon-Fri Or choose a specific day to have it run. Push the Create /Update Task button. It will show the settings in the output pane. It will also re-load the settings evertime you open the Sch Tasks window." Margin="10,1,0,0" Width="862" Height="47" Foreground="#ffffff"/>
</Border>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="385" Width="7" Margin="506.5,68.984375,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Output" Margin="519.5,71.03125,0,0" Width="69" Height="17"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Tasks_Output}" Margin="518.5,94.03125,0,0" Width="363" Height="280" Background="#e5e4eb" Name="lqdz6539kygzj"/>


<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Time" Margin="57,162,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="AM/PM" Margin="176,161,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Day(s)" Margin="293,161,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="(When you click the button, give it a few seconds to do its work and update the output box)" Margin="48,247,0,0" Width="406" Height="43"/>
</Grid>
</TabItem><TabItem Header="Update" Visibility="Collapsed"><Grid Background="#ffffff" Name="UpdateGrid" Margin="1,0,-1,0">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Installed Configurator Version" Margin="199,86,0,0" Width="176" Height="22"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Configurator_Ver}" Margin="248,112,0,0" Width="54" Height="20" Background="#eef3e9" Name="InstalledVersionConfig" Padding="6,2,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding AvailWebConfVer}" Margin="247,177,0,0" Width="54" Height="20" Background="#eef3e9" Padding="6,2,0,0" Name="lqdz6539t91vh"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Engine_Ver}" Margin="531,111,0,0" Background="#eef3e9" Width="54" Height="20" Padding="6,2,0,0" Name="lqdz6539tbccx"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding AvailWebEngVer}" Margin="533,178,0,0" Width="54" Height="20" Background="#eef3e9" Padding="6,2,0,0" Name="lqdz65390e8dx"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Update_Output}" Margin="81,6,0,0" Width="710" Height="45" Background="{Binding UpdateHeaderColor}" Name="lqdz6539o3fsi"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Available Configurator Version" Margin="195,151,0,0" Width="181" Height="23"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Installed Engine Version" Margin="494,88,0,0" Width="128" Height="23"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Available Engine Version" Margin="492,153,0,0" Width="130" Height="22"/>
<Button Content="Update ASPC" HorizontalAlignment="Left" VerticalAlignment="Top" Width="119" Margin="358,276,0,0" Name="UpdateButton" Height="29" Foreground="#194d4a" BorderBrush="#194d4a" FontSize="15" FontWeight="Bold" Background="#ffffff"/>
</Grid>
</TabItem><TabItem Header="About" Visibility="Collapsed"><Grid Background="#ffffff" Name="AboutGrid">
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="ASPC Configurator Version:" Margin="20,121,0,0" Width="146" Height="19"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="ASPC Engine Version:" Margin="20,150,0,0" Width="117" Height="19"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="ASPC written by Enigma-Tek" Margin="17,36,0,0" Width="378" Height="18"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Problems/Bugs, feedback or requests please email enigma-tek@outlook.com" Margin="17,56,0,0" Width="555" Height="29"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Github Project and Documentation (see Wiki) https://github.com/enigma-tek/ASPC_Pub" Margin="17,77,0,0" Width="531" Height="21"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Configurator_Ver}" Margin="171,118,0,0" Background="#edebf3" Width="58" Height="19" Padding="5,0,0,0" Name="lqdz65396yoyy"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding Engine_Ver}" Margin="141,148,0,0" Width="58" Height="19" Background="#edebf3" Padding="5,0,0,0" Name="lqdz6539wmb4w"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="ABOUT" Margin="15,6,0,0" Width="86" Height="25" FontSize="017" FontFamily="Sitka Small"/>
</Grid></TabItem></TabControl>
</Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#region Main_Page_Buttons
#Button Functions - they switch the tabindex number which in turn changes the tabs
function HomeBTNFunk {
    Async {
        $State.tabIndex = "0"
    }
}
function AlertDefBTNFunk {
    Async {
        $State.tabIndex = "1"
    }
}
function ExtraBTNFunk {
    Async {
        $State.tabIndex = "2"
    }
}
function EngConfBTNFunk {
    Async {
        $State.tabIndex = "3"
    }
}
function CredsBTNFunk {
    Async {
        $State.tabIndex = "4"
    }
}
function TasksBTNFunk {
    Async {
        $State.tabIndex = "5"
    }
}
function UpdateBTNFunk {
    Async {
        $State.tabIndex = "6"
    }
}
function AboutBTNFunk {
    Async {
        $State.tabIndex = "7"
    }
}

#endregion 
#region Alert_Def_Page
function open_alertDef {
    $OAlrtDef = Get-childItem "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs" 
    try {
    $AlertDefChoiceOpen = $OAlrtDef.Name | Out-GridView -title "Please select an Alert Definition File to open. Highlight the name and hit the OK button at the bottom of the screen" -OutputMode Single
    } catch {
    alrtDef_Clear_Entries
    return
    }
    $LoadOpenChoice = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$AlertDefChoiceOpen" | ConvertFrom-Json
    $ExternalAlertMailAddr = $LoadOpenChoice.ExternalAlertMailAddr
    $ExternalAlertMailNote = $LoadOpenChoice.ExternalAlertMailNote
    $SPNotes = $LoadOpenChoice.SPNotes
    $SPNamePortal = $LoadOpenChoice.SPNamePortal
    $InternalOwner = $LoadOpenChoice.IntOwner
    $ExtOwner = $LoadOpenChoice.ExtOwner

    $State.AlrtDefXMail = $ExternalAlertMailAddr
    $State.AlrtDefXNote = $ExternalAlertMailNote
    $State.SpNote = $SPNotes
    $State.SPName = $SPNamePortal
    $State.IntOwnerNote = $InternalOwner
    $State.ExtOwnerNote = $ExtOwner
}

function save_alertDef {
   
    Async {
        $State.Alrt_Def_Output = ""
        Start-Sleep -Seconds 1
    $ExternalAlertMailAddr = $State.AlrtDefXMail
    $ExternalAlertMailNote = $State.AlrtDefXNote
    $SPNotes = $State.SpNote
    $SPNamePortal = $State.SPName
    $InternalOwner = $State.IntOwnerNote
    $ExtOwner = $State.ExtOwnerNote
    
    $Global:AlrtDefSaveTime = Get-Date -Format 'h.mm tt - MM.dd.yyyy '
    $UserName = $env:UserName
     $AlrtDefSave = @{
        ExternalAlertMailAddr = $ExternalAlertMailAddr
        ExternalAlertMailNote = $ExternalAlertMailNote
        SPNotes = $SPNotes
        SPNamePortal = $SPNamePortal
        LastWriteTime = $Global:AlrtDefSaveTime
        LastWriteUser = $UserName
        IntOwner = $InternalOwner
        ExtOwner = $ExtOwner
        }
        
        if ($SPNamePortal -eq "" -OR $SPNamePortal -eq $null) {
            $State.Alrt_Def_Output = "Issue: Name cannot be blank"
        } else {
            if (-not(Test-Path -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPNamePortal.json" -PathType Leaf)) {
                $AlrtDefSaveTime = Get-Date -Format 'h.mm tt - MM.dd.yyyy '
                $State.Alrt_Def_Output = "New Alert Definition File Created - `n$AlrtDefSaveTime"
                $AlrtDefSave | ConvertTo-Json | Set-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPNamePortal.json"
            } else {
            $AlrtDefSaveTime = Get-Date -Format 'h.mm tt - MM.dd.yyyy '
            $State.Alrt_Def_Output = "Current Alert Definition Updated - `n$AlrtDefSaveTime"
            $AlrtDefSave | ConvertTo-Json | Set-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$SPNamePortal.json"
            }
        }
    }
}

function alrtDef_Delete {
   $AlrtDefSaveTime = Get-Date -Format 'h.mm tt - MM.dd.yyyy '
    $OAlrtDefDel = Get-childItem "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\" 
    $Global:AlertDefDelete = $OAlrtDefDel.Name | Out-GridView -title "DELETIONS CONNOT BE UNDONE!! Please select an Alert Definition File TO DELETE. Highlight the name and hit the OK button at the bottom of the screen" -OutputMode Single
    if ($Global:AlertDefDelete -like "*json") {
    Write-Output "Hi There"
        Remove-Item -Path "C:\Program Files\Enigma-Tek\ASPC\Configs\AlertDefs\$AlertDefDelete" -Verbose
    $State.Alrt_Def_Output = "Alert Definition `n$AlertDefDelete `nWas Deleted on `n$AlrtDefSaveTime"
    alrtDef_Clear_Entries
    } else {
        #do nothing here
    }
}

function alrtDef_Page_Load {
    $State.AlrtDefXMail = ""
    $State.AlrtDefXNote = ""
    $State.SpNote = ""
    $State.SPName = ""
    $State.ExtOwnerNote = ""
    $State.IntOwnerNote = ""
}

function alrtDef_Clear_Entries {
    $State.AlrtDefXMail = ""
    $State.AlrtDefXNote = ""
    $State.SpNote = ""
    $State.SPName = ""
    $State.ExtOwnerNote = ""
    $State.IntOwnerNote = ""
}
#endregion 
#region Eng_Config_Page
function eng_conf_page_load {
    Async {
    $engConfJson = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\Eng_Config.json" | ConvertFrom-Json
        $DefaultEmail = $engConfJson.DefaultEmail
        $XcludeExpired = $engConfJson.XcludeExpired
        $ExpireWithin = $engConfJson.ExpireWithin
        $ReportName = $engConfJson.ReportName
        
        $State.EngConfDefaultEmail = $DefaultEmail
        $State.EngConfExcludeExpired = $XcludeExpired
        $State.EngConfExpireWithin = $ExpireWithin
        $State.EngConfReportName = $ReportName
     
        if ($XcludeExpired -eq "YES") {
            $State.ExcludeExpiredCheckBox = "TRUE"
        } else {
            $State.ExcludeExpiredCheckBox = "FALSE"
        }
      }
}

function eng_conf_save_Button {
    Async {
        $DefaultEmail = $State.EngConfDefaultEmail
        $XcludeExpired = $State.EngConfExcludeExpired
        $ExpireWithin = $State.EngConfExpireWithin
        $ReportName = $State.EngConfReportName
         
            if ($ExpireWithin -gt "0" -AND $ExpireWithin -le "90") {
             $engConfig = @{
            ExpireWithin = $ExpireWithin
            XcludeExpired = $XcludeExpired
            DefaultEmail = $DefaultEmail
            ReportName = $ReportName
        } 
            $engConfig | ConvertTo-Json | Set-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\Eng_Config.json"
            $EngSaveTime = Get-Date -Format 'h.mm tt - MM.dd.yyyy '
            $State.Eng_Save_Output = "Engine Configuration Saved! Save time - $EngSaveTime"
    } else {
        $State.Eng_Save_Output = "ERROR - You must choose a Number of days out to Expiration between 1-90. Please change and re-save"
        break
    }
  }
}

function eng_excludeBox_checked {
    Async {
       $State.EngConfExcludeExpired = "YES" 
    }
}

function eng_excludeBox_unchecked {
    Async {
       $State.EngConfExcludeExpired = "NO" 
    }
}

function testEngineCreds {
    Start-Sleep -Seconds 2
    $cred1 = "C:\Program Files\Enigma-Tek\ASPC\Creds\emailCred.txt"
    if (-not(Test-Path -Path $cred1 -Pathtype Leaf)) {
        $cred1Num = "0"
    } else {
        $cred1Num = "1"
    }
     $cred2 = "C:\Program Files\Enigma-Tek\ASPC\Creds\emailFrom.txt"
    if (-not(Test-Path -Path $cred2 -Pathtype Leaf)) {
        $cred2Num = "0"
    } else {
        $cred2Num = "1"
    }
     $cred3 = "C:\Program Files\Enigma-Tek\ASPC\Creds\spAppID.txt"
    if (-not(Test-Path -Path $cred3 -Pathtype Leaf)) {
        $cred3Num = "0"
    } else {
        $cred3Num = "1"
    }
     $cred4 = "C:\Program Files\Enigma-Tek\ASPC\Creds\spTennantID.txt"
    if (-not(Test-Path -Path $cred4 -Pathtype Leaf)) {
        $cred4Num = "0"
    } else {
        $cred4Num = "1"
    }
     $cred5 = "C:\Program Files\Enigma-Tek\ASPC\Creds\spThumb.txt"
    if (-not(Test-Path -Path $cred5 -Pathtype Leaf)) {
        $cred5Num = "0"
    } else {
        $cred5Num = "1"
    }
    
    if (($cred1Num -eq "0") -or ($cred2Num -eq "0") -or ($cred3Num -eq "0") -or ($cred4Num -eq "0") -or ($cred5Num -eq "0")) {
        $testEngText += "Looks like one of the credentials is missing. Please review your credentials and re-run"
        $State.Eng_Save_Output = $testEngText
    } else {
        Start-Sleep -Seconds 3
        testEngineTask
       
    }
}
function testEngineTask {
        try {
        Get-ScheduledTask -taskname "ASPC_Engine_Run" -ErrorAction Stop
        Start-Sleep -Seconds 2
        $testEngText = "Task found"
        $State.Eng_Save_Output = $testEngText
        Start-ScheduledTask -TaskName "ASPC_Engine_Run"
        $testEngText += "`nASPC_Engine_Run task started. Please wait a couple of minutes to recieve the email. If you dont recieve an email, please check the Engine logs for errors."
        $State.Eng_Save_Output = $testEngText

        } catch {

        $testEngText = "There was no task found with name ASPC_Engine_Run. Was it created yet? If not please go to the tasks page and create the task"
        $State.Eng_Save_Output = $testEngText
        }
        
}
    

#endregion 
#region About_Page
function load_about {
        $versionLoad = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\Version\versions.json" | ConvertFrom-Json
        $confVer = $versionLoad.Conf_Ver
        $engVer = $versionLoad.Eng_ver
        
        $State.Configurator_Ver = $confVer
        $State.Engine_Ver = $engVer
}
#endregion 
#region Credential_Page_Service_Principal
Function Invoke-InputBoxPFX {

    [cmdletbinding(DefaultParameterSetName="plain")]
    [OutputType([system.string],ParameterSetName='plain')]
    [OutputType([system.security.securestring],ParameterSetName='secure')]

    Param(
        [Parameter(ParameterSetName="secure")]
        [Parameter(HelpMessage = "Enter the title for the input box. No more than 25 characters.",
        ParameterSetName="plain")]        

        [ValidateNotNullorEmpty()]
        [ValidateScript({$_.length -le 25})]
        [string]$Title = "User Input",

        [Parameter(ParameterSetName="secure")]        
        [Parameter(HelpMessage = "Enter a prompt. No more than 50 characters.",ParameterSetName="plain")]
        [ValidateNotNullorEmpty()]
        [ValidateScript({$_.length -le 50})]
        [string]$Prompt = "Please enter a value:",
        
        [Parameter(HelpMessage = "Use to mask the entry and return a secure string.",
        ParameterSetName="secure")]
        [switch]$AsSecureString
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type –assemblyName PresentationCore
    Add-Type –assemblyName WindowsBase

    #remove the variable because it might get cached in the ISE or VS Code
    Remove-Variable -Name myInput -Scope script -ErrorAction SilentlyContinue

    $form = New-Object System.Windows.Window
    $stack = New-object System.Windows.Controls.StackPanel

    #define what it looks like
    $form.Title = $title
    $form.Height = 150
    $form.Width = 350

    $label = New-Object System.Windows.Controls.Label
    $label.Content = "    $Prompt"
    $label.HorizontalAlignment = "left"
    $stack.AddChild($label)

    if ($AsSecureString) {
        $inputbox = New-Object System.Windows.Controls.PasswordBox
    }
    else {
        $inputbox = New-Object System.Windows.Controls.TextBox
    }

    $inputbox.Width = 300
    $inputbox.HorizontalAlignment = "center"

    $stack.AddChild($inputbox)

    $space = new-object System.Windows.Controls.Label
    $space.Height = 10
    $stack.AddChild($space)

    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = "_OK"

    $btn.Width = 65
    $btn.HorizontalAlignment = "center"
    $btn.VerticalAlignment = "bottom"

    #add an event handler
    $btn.Add_click( {
            if ($AsSecureString) {
                $script:myInput = $inputbox.SecurePassword
            }
            else {
                $script:myInput = $inputbox.text
            }
            $form.Close()
        })

    $stack.AddChild($btn)
    $space2 = new-object System.Windows.Controls.Label
    $space2.Height = 10
    $stack.AddChild($space2)

    $btn2 = New-Object System.Windows.Controls.Button
    $btn2.Content = "_Cancel"

    $btn2.Width = 65
    $btn2.HorizontalAlignment = "center"
    $btn2.VerticalAlignment = "bottom"

    #add an event handler
    $btn2.Add_click( {
            $form.Close()
        })

    $stack.AddChild($btn2)

    #add the stack to the form
    $form.AddChild($stack)

    #show the form
    $inputbox.Focus() | Out-Null
    $form.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen

    $form.ShowDialog() | out-null

    #write the result from the input box back to the pipeline
    $script:myInput

}

function pfxPasswordSet {
$Global:PFXPASS = Invoke-InputBoxPFX -Title "Secure Prompt" -Prompt "Enter your PFX Password" -AsSecureString
}

function openPFXFile([string] $initialDirectory){

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $Global:PFX = New-Object System.Windows.Forms.OpenFileDialog
    $Global:PFX.initialDirectory = [Environment]::GetFolderPath('Desktop')
    $Global:PFX.filter = "Certificate File|*.pfx|All files|*.*"
    $Global:PFX.ShowDialog() |  Out-Null
    
    $Global:PFXfileName = $Global:PFX.filename
    $State.PFX_Location = $Global:PFXfileName
}

function saveSPInformation {
            $State.SP_Page_Output = ""
            Start-Sleep -Seconds 2
        	$filepath = $Global:PFXfileName
        try {
			Import-PfxCertificate -FilePath $filepath -CertStoreLocation 'Cert:\LocalMachine\My' -Password $Global:PFXPASS -ErrorAction Stop
			Import-PfxCertificate -FilePath $filepath -CertStoreLocation 'Cert:\CurrentUser\My' -Password $Global:PFXPASS -ErrorAction Stop
            $State.SP_Page_Output = "Certificate Imported."
            }
       catch {
           $State.SP_Page_Output = "There was an issue importing the PFX."
            }
    
        $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
        $Thumbprint = $State.Cert_Thumbprint | ConvertTo-SecureString -AsPlainText -Force
        $Thumbprint | ConvertFrom-SecureString -Key $ASPCKeyData | Out-File "C:\Program Files\Enigma-Tek\ASPC\Creds\spThumb.txt"
        $State.SP_Page_Output += " Thumbprint has been encrypted and stored."
        
        $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
        $AppID = $State.SP_AppID | ConvertTo-SecureString -AsPlainText -Force
        $AppID | ConvertFrom-SecureString -Key $ASPCKeyData | Out-File "C:\Program Files\Enigma-Tek\ASPC\Creds\spAppID.txt"
        $State.SP_Page_Output += " Application ID has been encrypted and stored."
        
        $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
        $TennantID = $State.SP_TennantID | ConvertTo-SecureString -AsPlainText -Force
        $TennantID | ConvertFrom-SecureString -Key $ASPCKeyData | Out-File "C:\Program Files\Enigma-Tek\ASPC\Creds\spTennantID.txt"
        $State.SP_Page_Output += " Tenant ID has been encrypted and stored."
        
}


#endregion 
#region Credential_Page_Email_Account
Function Invoke-InputBoxEmail {

    [cmdletbinding(DefaultParameterSetName="plain")]
    [OutputType([system.string],ParameterSetName='plain')]
    [OutputType([system.security.securestring],ParameterSetName='secure')]

    Param(
        [Parameter(ParameterSetName="secure")]
        [Parameter(HelpMessage = "Enter the title for the input box. No more than 25 characters.",
        ParameterSetName="plain")]        

        [ValidateNotNullorEmpty()]
        [ValidateScript({$_.length -le 25})]
        [string]$Title = "User Input",

        [Parameter(ParameterSetName="secure")]        
        [Parameter(HelpMessage = "Enter a prompt. No more than 50 characters.",ParameterSetName="plain")]
        [ValidateNotNullorEmpty()]
        [ValidateScript({$_.length -le 50})]
        [string]$Prompt = "Please enter a value:",
        
        [Parameter(HelpMessage = "Use to mask the entry and return a secure string.",
        ParameterSetName="secure")]
        [switch]$AsSecureString
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type –assemblyName PresentationCore
    Add-Type –assemblyName WindowsBase

    #remove the variable because it might get cached in the ISE or VS Code
    Remove-Variable -Name myInput -Scope script -ErrorAction SilentlyContinue

    $form = New-Object System.Windows.Window
    $stack = New-object System.Windows.Controls.StackPanel

    #define what it looks like
    $form.Title = $title
    $form.Height = 150
    $form.Width = 350

    $label = New-Object System.Windows.Controls.Label
    $label.Content = "    $Prompt"
    $label.HorizontalAlignment = "left"
    $stack.AddChild($label)

    if ($AsSecureString) {
        $inputbox = New-Object System.Windows.Controls.PasswordBox
    }
    else {
        $inputbox = New-Object System.Windows.Controls.TextBox
    }

    $inputbox.Width = 300
    $inputbox.HorizontalAlignment = "center"

    $stack.AddChild($inputbox)

    $space = new-object System.Windows.Controls.Label
    $space.Height = 10
    $stack.AddChild($space)

    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = "_OK"

    $btn.Width = 65
    $btn.HorizontalAlignment = "center"
    $btn.VerticalAlignment = "bottom"

    #add an event handler
    $btn.Add_click( {
            if ($AsSecureString) {
                $script:myInput = $inputbox.SecurePassword
            }
            else {
                $script:myInput = $inputbox.text
            }
            $form.Close()
        })

    $stack.AddChild($btn)
    $space2 = new-object System.Windows.Controls.Label
    $space2.Height = 10
    $stack.AddChild($space2)

    $btn2 = New-Object System.Windows.Controls.Button
    $btn2.Content = "_Cancel"

    $btn2.Width = 65
    $btn2.HorizontalAlignment = "center"
    $btn2.VerticalAlignment = "bottom"

    #add an event handler
    $btn2.Add_click( {
            $form.Close()
        })

    $stack.AddChild($btn2)

    #add the stack to the form
    $form.AddChild($stack)

    #show the form
    $inputbox.Focus() | Out-Null
    $form.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen

    $form.ShowDialog() | out-null

    #write the result from the input box back to the pipeline
    $script:myInput

}

function emailPasswordSet {
$Global:EmailPASS = Invoke-InputBoxEmail -Title "Secure Prompt" -Prompt "Enter your Email Address Password" -AsSecureString
}

function testEmailAccount {
    $State.EmailAccount_Output = ""
    Start-Sleep -Seconds 2
        $State.EmailAccount_Output = "Sending Test Email"
        Try {
        $TestEmailSubject = "Test email from ASPC Configurator"
        $TestEmailBody = "Test email. If you recieved this email then you can go ahead and Click the Save button."
        $mailTestUser = $State.TestEmailAdd
        $mailFromuser = $State.EmailAddFrom
        $credential = New-Object System.Management.Automation.PSCredential ($mailFromuser,$Global:EmailPASS)
                
        $mailParams =@{
            SmtpServer          = 'smtp.office365.com'
            UseSsl              = $true
            Credential          = $credential
            From                = $mailFromuser
            To                  = $mailTestUser
            Subject             = $TestEmailSubject
            Body                = $TestEmailBody
        }
    Send-MailMessage @mailParams
    $State.EmailAccount_Output += "`nTest Email looks to have sent. Please check your inbox"
        } 
        Catch {
            
        $State.EmailAccount_Output += "`nThere was an issue sending the Test message. Please check all your settings or the account."
        }
        
}


function saveEmailAccount {
    $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
    $emailFromUserConf = $State.EmailAddFrom | ConvertTo-SecureString -AsPlainText -Force
    $emailFromUserConf | ConvertFrom-SecureString -Key $ASPCKeyData | Out-File "C:\Program Files\Enigma-Tek\ASPC\Creds\emailFrom.txt"
    $State.EmailAccount_Output += "Email account has been encrypted and saved"
    
    $ASPCKeyData = Get-Content "C:\Program Files\Enigma-Tek\ASPC\Configs\Engine\ASPC_Eng.key"
    $emailPassSave = $Global:EmailPASS | ConvertTo-SecureString -AsPlainText -Force
    $Global:EmailPASS | ConvertFrom-SecureString -Key $ASPCKeyData | Out-File "C:\Program Files\Enigma-Tek\ASPC\Creds\emailCred.txt"
    $State.EmailAccount_Output += "Email password has been encrypted and saved"
}










#endregion 
#region Scheduled_Tasks_Page

function loadTaskPage {
    try {
    $getTask = Get-ScheduledTask -TaskName "ASPC_Engine_Run"
    
    if ($getTask.Triggers.DaysInterval -eq "1") {
        $runDays = "Daily"
    } else { 
    $getTask.Triggers.DaysOfWeek
       [Flags()] enum DaysOfWeek {
            Sunday = 1
            Monday = 2
            Tuesday = 4
            Wednesday = 8
            Thursday = 16
            Friday = 32
            Saturday = 64
        }
    $runDays = [DaysOfWeek]$getTask.Triggers.DaysOfWeek
    }
    $taskRunTime = $getTask.Triggers.StartBoundary
    $getTask.Triggers.StartBoundary
        $timeSplit = $getTask.Triggers.StartBoundary.Split('T')[1]
        $timeSplit2 = $timeSplit.split('-')[0]
    
    
    $State.Tasks_Output = ""
    Start-Sleep -Seconds 1
    $State.Tasks_Output += $Global:EngineTaskUpdated
    $State.Tasks_Output += $Global:EngineTaskCreated
    $State.Tasks_Output += "Current Task Settings:"
    $State.Tasks_Output += "`nTask Name: ASPC_Engine_Run"
    $State.Tasks_Output += "`nRun Days: $runDays"
    $State.Tasks_Output += "`nStart Time (24H Format): $timeSplit2"
    } catch {
        $State.Tasks_Output = "No Task has been created yet"
    }
}

function saveTask {
    
    $SchTime = $State.SelectedHour
    $SchDay = $State.SelectedRunDay
    $SchAMPM = $State.SelectedTOD
    $ActualTime = $State.SelectedHour
    $ActualTime += $State.SelectedTOD
   
    if ($SchTime -eq "" -OR $SchDay -eq "" -OR $SchAMPM -eq "") {

        $State.Tasks_Output = ""
        $State.Tasks_Output = "You have not chosen one of the drop downs. Please make sure that you have chosen a Time, a Time of Day and Day(s)"
         
    } else {
        $scriptPath = "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Engine.ps1"
        $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-ExecutionPolicy Bypass -file ""$scriptPath"""
        $principal = New-ScheduledTaskPrincipal -UserID "SYSTEM" -RunLevel Highest -LogonType S4U
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -Compatibility Win8 -ExecutionTimeLimit (New-Timespan -Minutes 45)
        
      
  if ($SchDay -eq "Daily") {
        $trigger = New-ScheduledTaskTrigger -Daily -At $ActualTime
        Start-Sleep -seconds 4
        } elseif ($SchDay -eq "Work Week") {
            $SchDayFin = "Monday","Tuesday","Wednesday","Thursday","Friday"
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $SchDayFin -At $ActualTime
        } else {
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $SchDay -At $ActualTime
        }
        
          try {
            Get-ScheduledTask -TaskName "ASPC_Engine_Run" -Verbose -ErrorAction Stop
            Set-ScheduledTask -TaskName 'ASPC_Engine_Run' -Trigger $trigger -Action $action -Principal $principal -Settings $settings
            $Global:EngineTaskUpdated = "ASPC_Engine_Run task has been updated `n"
            $Global:EngineTaskCreated = ""
            loadTaskPage
          } catch {
            Register-ScheduledTask -TaskName 'ASPC_Engine_Run' -Trigger $trigger -Action $action -Principal $principal -Settings $settings
            $Global:EngineTaskCreated = "Task ASPC_Engine_Run has been created `n"
            $Global:EngineTaskUpdated =  ""
            loadTaskPage
          }
    }
}


    
    
    

    
#endregion 
#region Update_Page
function openUpdatePage {
    Async {
        $versionLoad = Get-content "C:\Program Files\Enigma-Tek\ASPC\Configs\Version\versions.json" | ConvertFrom-Json
        $confVer = $versionLoad.Conf_Ver
        $engVer = $versionLoad.Eng_ver
        
        $State.Configurator_Ver = $confVer
        $State.Engine_Ver = $engVer
        
        $getVersionsWeb = Invoke-WebRequest -URI "https://raw.githubusercontent.com/enigma-tek/ASPC_Pub/main/Files/Software/versions.json" | ConvertFrom-Json
        $engVersionWeb = $getVersionsWeb.Eng_Ver
        $confVersionWeb = $getVersionsWeb.Conf_Ver
        
        $State.AvailWebEngVer = $engVersionWeb
        $State.AvailWebConfVer = $confVersionWeb
        
        isThereAnUpdate
    }
}

function isThereAnUpdate {
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
            $State.UpdateHeaderColor = "#F8E71C"
            $State.Update_Output = "There is an update available for ASPC. Use the Update ASPC button to update ASPC. The configurator will close and another box will pop up. Once it is done updating, it will reopen the configurator."
            } else { 
            $State.Update_Output = "ASPC is currently up to date"
    }
}

function updateButton {
    
        $PID | Out-File -FilePath "C:\Program Files\Enigma-Tek\ASPC\Software\tempPid.tmp"
        Start-Sleep -Seconds 2
        $powershellPath = "$env:windir\system32\windowspowershell\v1.0\powershell.exe"
        $scriptPath = "C:\Program Files\Enigma-Tek\ASPC\Software\ASPC_Updater.ps1"
        Start-Process $powershellPath -ArgumentList "-ExecutionPolicy Bypass ""& '$scriptPath'"""
    
}







#endregion 


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$HomeBTN.Add_Click({HomeBTNFunk $this $_})
$ArtDefBtn.Add_Click({AlertDefBTNFunk $this $_})
$EngConfBtn.Add_Click({EngConfBTNFunk $this $_})
$CredsConfBtn.Add_Click({CredsBTNFunk $this $_})
$TasksBtn.Add_Click({TasksBTNFunk $this $_})
$UpdateBtn.Add_Click({UpdateBTNFunk $this $_})
$ParentTabControl.Add_Loaded({openUpdatePage $this $_})
$AboutBTN.Add_Click({AboutBTNFunk $this $_})
$AlrtDefGrid.Add_Loaded({alrtDef_Page_Load $this $_})
$Delete_Def_Button.Add_Click({alrtDef_Delete $this $_})
$lqdz6538596bc.Add_Click({alrtDef_Clear_Entries $this $_})
$lqdz6538607uy.Add_Click({save_alertDef $this $_})
$lqdz6538y1xjk.Add_Click({open_alertDef $this $_})
$EngConfGrid.Add_Initialized({eng_conf_page_load $this $_})
$EngConfGrid.Add_Loaded({eng_conf_page_load $this $_})
$eng_conf_save_button.Add_Click({eng_conf_save_Button $this $_})
$lqdz6538eu0an.Add_Checked({eng_excludeBox_checked $this $_})
$lqdz6538eu0an.Add_Unchecked({eng_excludeBox_unchecked $this $_})
$lqdz65381fl8d.Add_Click({testEngineCreds $this $_})
$lqdz6538yvw1t.Add_Click({openPFXFile $this $_})
$lqdz6538hjmpc.Add_Click({saveSPInformation $this $_})
$lqdz65386ksn8.Add_Click({pfxPasswordSet $this $_})
$lqdz6538625bd.Add_Click({emailPasswordSet  $this $_})
$lqdz6538dflgy.Add_Click({testEmailAccount $this $_})
$lqdz6538lly1e.Add_Click({saveEmailAccount $this $_})
$TasksGrid.Add_Loaded({loadTaskPage $this $_})
$lqdz6539fetvn.Add_Click({saveTask $this $_})
$UpdateGrid.Add_Loaded({openUpdatePage $this $_})
$UpdateButton.Add_Click({updateButton $this $_})
$AboutGrid.Add_Loaded({load_about $this $_})

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
    "tabIndex" : "6",
    "EngConfDefaultEmail" : "",
    "EngConfExcludeExpired" : "",
    "ExcludeExpiredCheckBox" : "",
    "EngConfExpireWithin" : "",
    "EngConfReportName" : "",
    "Eng_Save_Output" : "",
    "AlrtDefXMail" : "",
    "AlrtDefXNote" : "",
    "SpNote" : "",
    "SPName" : "",
    "IntOwnerNote" : "",
    "ExtOwnerNote" : "",
    "Alrt_Def_Output" : "",
    "CredsTabIndex" : "2",
    "Configurator_Ver" : "",
    "Engine_Ver" : "",
    "PFX_Location" : "File Location:",
    "SP_Page_Output" : "",
    "SP_AppID" : "",
    "SP_TennantID" : "",
    "Cert_Thumbprint" : "",
    "PFX_Location" : "",
    "PFX_Pass" : "",
    "EmailAddFrom" : "",
    "TestEmailAdd" : "",
    "EmailAccount_Output" : "",
    "HourComboBox" : ["1:00","2:00","3:00","4:00","5:00","6:00","7:00","8:00","9:00","10:00","11:00","12:00"],
    "HourAMPM" : ["AM","PM"],
    "SelectedHour" : "",
    "SelectedTOD" : "",
    "RunDay" : ["Daily","Work Week","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"],
    "SelectedRunDay" : "",
    "Tasks_Output" : "",
    "Update_Output" : "",
    "AvailWebEngVer" : "",
    "AvailWebConfVer" : "",
    "UpdateHeaderColor" : "#7ED321"
    
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("tabIndex","EngConfDefaultEmail","EngConfExcludeExpired","ExcludeExpiredCheckBox","EngConfExpireWithin","EngConfReportName","Eng_Save_Output","AlrtDefXMail","AlrtDefXNote","SpNote","SPName","IntOwnerNote","ExtOwnerNote","Alrt_Def_Output","CredsTabIndex","Configurator_Ver","Engine_Ver","PFX_Location","SP_Page_Output","SP_AppID","SP_TennantID","Cert_Thumbprint","PFX_Pass","EmailAddFrom","TestEmailAdd","EmailAccount_Output","HourComboBox","HourAMPM","SelectedHour","SelectedTOD","RunDay","SelectedRunDay","Tasks_Output","Update_Output","AvailWebEngVer","AvailWebConfVer","UpdateHeaderColor") 

$Window.DataContext = $DataContext
Set-Binding -Target $ParentTabControl -Property $([System.Windows.Controls.TabControl]::SelectedIndexProperty) -Index 0 -Name "tabIndex"  
Set-Binding -Target $lqdz6538i3o5f -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 10 -Name "SPName"  
Set-Binding -Target $lqdz6538zvwox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 9 -Name "SpNote"  
Set-Binding -Target $lqdz6538a81wr -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 7 -Name "AlrtDefXMail"  
Set-Binding -Target $lqdz6538u0xsb -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 8 -Name "AlrtDefXNote"  
Set-Binding -Target $lqdz653839nyj -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 13 -Name "Alrt_Def_Output"  
Set-Binding -Target $lqdz6538ea67r -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 11 -Name "IntOwnerNote"  
Set-Binding -Target $lqdz6538hsdz1 -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 12 -Name "ExtOwnerNote"  
Set-Binding -Target $lqdz6538zeufi -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 1 -Name "EngConfDefaultEmail"  
Set-Binding -Target $lqdz6538vh6tn -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 4 -Name "EngConfExpireWithin"  
Set-Binding -Target $lqdz6538wq570 -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 5 -Name "EngConfReportName"  
Set-Binding -Target $lqdz6538eu0an -Property $([System.Windows.Controls.CheckBox]::IsCheckedProperty) -Index 3 -Name "ExcludeExpiredCheckBox"  
Set-Binding -Target $Eng_Conf_Save_output -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 6 -Name "Eng_Save_Output"  
Set-Binding -Target $lqdz6538ui0nq -Property $([System.Windows.Controls.TabControl]::SelectedIndexProperty) -Index 14 -Name "CredsTabIndex"  
Set-Binding -Target $lqdz6538mgwwn -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 19 -Name "SP_AppID"  
Set-Binding -Target $lqdz6538ncrbo -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 20 -Name "SP_TennantID"  
Set-Binding -Target $lqdz65388pieb -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 21 -Name "Cert_Thumbprint"  
Set-Binding -Target $lqdz6538tagm3 -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 17 -Name "PFX_Location"  
Set-Binding -Target $lqdz6538nvblg -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 18 -Name "SP_Page_Output"  
Set-Binding -Target $lqdz6538ewjc6 -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 23 -Name "EmailAddFrom"  
Set-Binding -Target $Test_Address -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 24 -Name "TestEmailAdd"  
Set-Binding -Target $EmailAccount_Output -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 25 -Name "EmailAccount_Output"  
Set-Binding -Target $lqdz6539zcir2 -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 26 -Name "HourComboBox"  
Set-Binding -Target $lqdz6539zcir2 -Property $([System.Windows.Controls.ComboBox]::SelectedValueProperty) -Index 28 -Name "SelectedHour"  
Set-Binding -Target $lqdz6539laahl -Property $([System.Windows.Controls.ComboBox]::SelectedValueProperty) -Index 29 -Name "SelectedTOD"  
Set-Binding -Target $lqdz6539laahl -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 27 -Name "HourAMPM"  
Set-Binding -Target $lqdz6539f2zzy -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 30 -Name "RunDay"  
Set-Binding -Target $lqdz6539f2zzy -Property $([System.Windows.Controls.ComboBox]::SelectedValueProperty) -Index 31 -Name "SelectedRunDay"  
Set-Binding -Target $lqdz6539kygzj -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 32 -Name "Tasks_Output"  
Set-Binding -Target $InstalledVersionConfig -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 15 -Name "Configurator_Ver"  
Set-Binding -Target $lqdz6539t91vh -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 35 -Name "AvailWebConfVer"  
Set-Binding -Target $lqdz6539tbccx -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 16 -Name "Engine_Ver"  
Set-Binding -Target $lqdz65390e8dx -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 34 -Name "AvailWebEngVer"  
Set-Binding -Target $lqdz6539o3fsi -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 33 -Name "Update_Output"  
Set-Binding -Target $lqdz6539o3fsi -Property $([System.Windows.Controls.TextBlock]::BackgroundProperty) -Index 36 -Name "UpdateHeaderColor"  
Set-Binding -Target $lqdz65396yoyy -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 15 -Name "Configurator_Ver"  
Set-Binding -Target $lqdz6539wmb4w -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 16 -Name "Engine_Ver"  




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


