#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:d="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="d" Title="IT Help Tool" Width="500" Name="mainWindow">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="MinWidth" Value="125"/>
        </Style>
        <Style x:Key="HeaderStyle" TargetType="TextBlock">
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="SubheaderStyle" TargetType="TextBlock">
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style x:Key="OutputConsoleStyle" TargetType="TextBox">
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="IsReadOnly" Value="True"/>
        </Style>
        <ControlTemplate x:Key="HyperlinkLikeButtonTemplate" TargetType="{x:Type Button}">
            <TextBlock x:Name="innerText" Foreground="{DynamicResource {x:Static SystemColors.HotTrackBrushKey}}" Cursor="Hand">
        <ContentPresenter/>
            </TextBlock>
            <ControlTemplate.Triggers>
                <Trigger Property="Button.IsMouseOver" Value="true">
                    <Setter TargetName="innerText" Property="Foreground" Value="{DynamicResource {x:Static SystemColors.HighlightBrushKey}}"/>
                    <Setter TargetName="innerText" Property="TextDecorations" Value="Underline"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style x:Key="HyperlinkLikeButton" TargetType="{x:Type Button}">
            <Setter Property="Template" Value="{StaticResource HyperlinkLikeButtonTemplate}"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Border Grid.Column="0" Grid.Row="0" Height="35" Padding="5" Background="#FFE31E37">
            <TextBlock Style="{StaticResource HeaderStyle}" TextWrapping="Wrap" Text="{Binding titleText}" Name="titleTextBlock" Foreground="#ffffff"/>
        </Border>
        <TabControl Grid.Row="1" SelectedIndex="{Binding tabControlIndex}" TabStripPlacement="Left" Name="kddhyif4fnh4o">
            <TabItem Name="kdek7manke31y">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>



                    <!-- Basic Data Display -->
                    <Grid Grid.Row="1" Margin="20">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition/>
                            <RowDefinition/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition/>
                        </Grid.RowDefinitions>

                        <TextBlock Name="computerNameTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Computer Name:" Grid.Row="0" Grid.Column="0" Margin="10"/>
                        <TextBox Name="computerNameTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding computerName}" Style="{StaticResource OutputConsoleStyle}" Grid.Row="0" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="ipAddressTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="IP Address(es):" Grid.Row="1" Grid.Column="0" Margin="10"/>
                        <TextBox Name="ipAddressTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding ip}" Style="{StaticResource OutputConsoleStyle}" Grid.Row="1" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="publicIpAddressTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Public IP Address:" Grid.Row="2" Grid.Column="0" Margin="10"/>
                        <TextBox Name="publicIpAddressTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding publicIpAddress}" Style="{StaticResource OutputConsoleStyle}" Grid.Row="2" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="userNameTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" Text="Username:" Grid.Row="3" Grid.Column="0" Margin="10"/>
                        <TextBox Name="userNameTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding userName}" Style="{StaticResource OutputConsoleStyle}" Grid.Row="3" Grid.Column="1" Margin="10"/>

                        <Button Template="{StaticResource HyperlinkLikeButtonTemplate}" Content="Advanced info" Grid.Row="4" Margin="10" Name="advancedInfoButton"/>

                    </Grid>

                    <!-- Fix Addin buttons-->
                    <Border Grid.Row="3" Padding="10" BorderBrush="DodgerBlue" BorderThickness="1" Margin="10">
                        <StackPanel>
                            <TextBlock Text="Fix Office add-ins" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="Bold" Margin="0,0,0,5"/>
                            <Button Name="resetWordButton" Content="Reset Word add-ins"/>
                            <Button Name="resetOutlookButton" Content="Reset Outlook add-ins"/>
                            <Button Name="resetPowerpointButton" Content="Reset PowerPoint add-ins"/>
                            <Button Name="resetExcelButton" Content="Reset Excel add-ins"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Advanced Info page -->
            <TabItem>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition/>
                        <RowDefinition Height="Auto "/>
                    </Grid.RowDefinitions>

                    <TextBlock Style="{StaticResource SubheaderStyle}" TextWrapping="Wrap" Text="Advanced info" Name="advancedInfoTitleTextBlock"/>

                    <TextBox Name="advancedInfoTextBox" Grid.Row="1" Margin="10" Text="{Binding advancedInfoTextBox}" Style="{StaticResource OutputConsoleStyle}"/>
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                        <Button Name="copyToClipboardButton">Copy to clipboard</Button>
                        <Button Name="backButton">Back</Button>
                    </StackPanel>
                </Grid>


            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#Write your code here

#region +++++++++++++++++++++ Options +++++++++++++++++++++++++++++++
$script:ver                             = '1.0'
$script:vercaption                      = 'Created by Me'
$script:date                            = (Get-Date).ToString()
$script:utcdate                         = (get-date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
$script:formtitle                       = 'IT Help Tool'

function Get-PublicIpAddress {
    Write-Output ((New-Object Net.WebClient).DownloadString('http://ipinfo.io/ip').trim())
}

#region +++++++++++++++++++++ Gather Basic Data +++++++++++++++++++++ 
function Initialize-BasicDataPage {
    
    Async {
        $State.computerName             = $env:COMPUTERNAME
        $State.userName                 = $env:USERNAME
        $State.ip                       = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress } | Select-Object -Expand IPAddress | Where-Object { $_ -like '*.*.*.*' } | Out-String).Trim()
    }
    
    Async {
        $State.publicIpAddress          = Get-PublicIpAddress
    }
}

#endregion

#region +++++++++++++++++++++ Gather Advanced Data +++++++++++++++++++++
function Initialize-AdvancedDataPage {
    
    Add-Type -AssemblyName System.Windows.Forms
    
    #Required functions
    function Get-Uptime {
        $operatingSystem = Get-WmiObject Win32_OperatingSystem
        $uptime = (Get-Date) - ($operatingSystem.ConvertToDateTime($operatingSystem.LastBootUpTime))
        $display = '' + $uptime.Days + ' days, ' + $uptime.Hours + ' hours, ' + $uptime.Minutes + ' minutes'
        Write-Output $display
    }
    
    $netInfo = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress } | Select-Object IPAddress,MacAddress,Description,DNSServerSearchOrder,DHCPServer,DefaultIPGateway
  
    $ip0 = $netInfo | Select-Object -Expand IPAddress -First 1 | Where-Object { $_ -like '*.*.*.*' } | Out-String
    $ip0 = $ip0.trim()
    $mac0 = $netInfo| Select-Object -Expand MacAddress -First 1 | Out-String
    $nicName0 = $netInfo | Select-Object -Expand Description -First 1 | Out-String
    $ip1 = $netInfo[1] | Select-Object -Expand IPAddress | Where-Object { $_ -like '*.*.*.*' } | Out-String
    $ip1 = $ip1.trim()
    $mac1 = $netInfo[1] | Select-Object -Expand MacAddress | Out-String
    $nicname1 = $netInfo[1] | Select-Object -Expand Description | Out-String
    $domain = $env:USERDOMAIN
    $uptime = Get-Uptime
    $winVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption
    $computerSystem = Get-WmiObject -class Win32_ComputerSystem
    $compManufacturer = $computerSystem.Manufacturer
    $compModel = $computerSystem.Model
    $compSerial = (Get-WmiObject -class Win32_Bios).SerialNumber
    $processorName = (Get-WmiObject -class Win32_Processor).Name
    $memAmount = ((Get-WmiObject -class 'CIM_PhysicalMemory' | Measure-Object -Property Capacity -Sum).Sum) / 1GB
    $publicIpAddress = Get-PublicIpAddress
    $dnsServers = $netInfo | Select-Object -ExpandProperty DNSServerSearchOrder
    $dhcpServer = $netInfo | Select-Object -ExpandProperty DHCPServer
    $defaultGateway = ($netInfo | Select-Object -ExpandProperty DefaultIPGateway | Where-Object { $_ -like '*.*.*.*' } | Out-String).Trim()
    
    # Display results
    if (!$ip1) {
        $State.advancedInfoTextBox = "Computer Name:`t`t" + $State.computerName + "`nUser Name:`t`t" + $State.username + "`n`nNetwork Adapter:`t" + $nicname0 + "`nIP Address:`t`t" + $ip0 + "`nMAC Address:`t`t" + $mac0 + "`nDNS Servers:`t`t" + $dnsservers + "`nDHCP Server:`t`t" + $dhcpserver + "`nDefault Gateway:`t" + $defaultGateway + "`nPublic IP:`t`t" + $publicIpAddress + "`n`nDomain:`t`t`t" + $domain + "`n`nComputer Manufacturer:"+ $compManufacturer + "`nComputer Model:`t" + $compmodel + "`nSerial Number:`t`t" + $compSerial + "`nProcessor:`t`t" + $processorName + "`nMemory Amount:`t`t" + $memamount + "GB" + "`nSystem Uptime:`t`t" + $uptime + "`nWindows Version:`t" + $winversion
        
    }
    else {
        $State.advancedInfoTextBox = "Computer Name:`t`t" + $State.computerName + "`nUser Name:`t`t" + $State.username + "`n`nNetwork Adapter:`t" + $nicname0 + "`nIP Address:`t`t" + $ip0 + "`nMAC Address:`t`t" + $mac0 + "`nDNS Servers:`t`t" + $dnsservers + "`nDHCP Server:`t`t" + $dhcpserver + "`n`nNetwork Adapter 2:`t" + $nicname1 + "`nIP Address:`t`t" + $ip1 + "`nMAC Address:`t`t" + $mac1 + "`nDefault Gateway:`t" + $defaultGateway + "`nPublic IP:`t`t" + $publicIpAddress + "`n`nDomain:`t`t`t" + $domain + "`n`nComputer Manufacturer:"+ $compManufacturer + "`nComputer Model:`t" + $compmodel + "`nSerial Number:`t`t" + $compSerial + "`nProcessor:`t`t" + $processorName + "`nMemory Amount:`t`t" + $memamount + "GB" + "`nSystem Uptime:`t`t" + $uptime + "`nWindows Version:`t" + $winversion
    }
}
#endregion

#region +++++++++++++++++++++ Reset Addins +++++++++++++++++++++ 
function Initialize-OfficeAddins {
    param ([string]$application)
    
    $result = [Windows.Forms.MessageBox]::Show("This will reload all disabled $application add-ins. Would you like to continue?","Reset $application Add-ins","OKCancel","Warning")

    if ($result -eq "OK") {
        Try {
            $registry = Get-ChildItem -Path "HKCU:\SOFTWARE\Microsoft\Office" -Recurse | Where-Object {$_.Name -like "*.0\$application\Resiliency\DisabledItems" }
    
            if (!$registry) {
                [Windows.Forms.MessageBox]::Show("Could not find an Office installation.","Reset Office Add-ins","OK","Warning")
            }
       
            foreach ($reg in $registry) {
                Remove-Item -Path "Registry::$reg"
            }
        }
        Catch {
            [Windows.Forms.MessageBox]::Show("Failed to remove any Disabled Items for $application")
        }
    
        Try {
            $registry = Get-ChildItem -Path "HKCU:\SOFTWARE\Microsoft\Office\$application\Addins" -ErrorAction Stop
            foreach ($reg in $registry) {
                if ((Get-ItemPropertyValue -Path Registry::$reg -Name LoadBehavior) -lt 3) {
                    Set-ItemProperty -Path Registry::$reg -Name LoadBehavior -Value 3
                }
            }
        }
        Catch [System.Management.Automation.ItemNotFoundException] {
            [Windows.Forms.MessageBox]::Show("Could not find any $application Office installation.","Reset Office Add-ins","OK","Information")
        }
        Catch {
            [Windows.Forms.MessageBox]::Show("Failed to change enable add-ins for $application.","Reset Office Add-ins","OK","Warning")
        }
        
        [Windows.Forms.MessageBox]::Show("Reset $application complete.")

    }
    
}

#endregion

#region +++++++++++++++++++++ Buttons +++++++++++++++++++++ 
function Invoke-ResetWord {
    Initialize-OfficeAddins("Word")
}

function Invoke-ResetOutlook {
    Initialize-OfficeAddins("Outlook")
}

function Invoke-ResetPowerpoint {
    Initialize-OfficeAddins("Powerpoint")
}

function Invoke-ResetExcel {
    Initialize-OfficeAddins("Excel")
}

function Invoke-CopyClipboard {
    Write-Output $State.advancedInfoTextBox | & "$env:windir\system32\clip.exe"
}

function Invoke-GoBack {
    $State.tabControlIndex = "0"   
}

function Invoke-GoAdvanced {
    $State.tabControlIndex = "1"
    
    Async {
        Initialize-AdvancedDataPage
    }
    
}
#endregion
#endregion

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$kdek7manke31y.Add_Loaded({Initialize-BasicDataPage $this $_})
$advancedInfoButton.Add_Click({Invoke-GoAdvanced $this $_})
$resetWordButton.Add_Click({Invoke-ResetWord $this $_})
$resetOutlookButton.Add_Click({Invoke-ResetOutlook $this $_})
$resetPowerpointButton.Add_Click({Invoke-ResetPowerpoint $this $_})
$resetExcelButton.Add_Click({Invoke-ResetExcel $this $_})
$copyToClipboardButton.Add_Click({Invoke-CopyClipboard $this $_})
$backButton.Add_Click({Invoke-GoBack $this $_})

$State = [PSCustomObject]@{}


Function Set-Binding {
    Param($Target,$Property,$Index,$Name)
 
    $Binding = New-Object System.Windows.Data.Binding
    $Binding.Path = "["+$Index+"]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    


    [void]$Target.SetBinding($Property,$Binding)
}

function FillDataContext($props){

    For ($i=0; $i -lt $props.Length; $i++) {
   
   $prop = $props[$i]
   $DataContext.Add($DataObject."$prop")
   
    $getter = [scriptblock]::Create("return `$DataContext['$i']")
    $setter = [scriptblock]::Create("param(`$val) return `$DataContext['$i']=`$val")
    $State | Add-Member -Name $prop -MemberType ScriptProperty -Value  $getter -SecondValue $setter
               
       }
   }



$DataObject =  ConvertFrom-Json @"

{
    "titleText" : "IT Help Tool",
    "computerName" : "Computer name",
    "userName" : "User name",
    "ip" : "127.0.0.1",
    "publicIpAddress" : "127.0.0.1",
    "tabControlIndex" : "0",
    "advancedInfoTextBox" : "Loading..."
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("titleText","computerName","userName","ip","publicIpAddress","tabControlIndex","advancedInfoTextBox") 

$Window.DataContext = $DataContext
Set-Binding -Target $titleTextBlock -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 0 -Name "titleText"
Set-Binding -Target $kddhyif4fnh4o -Property $([System.Windows.Controls.TabControl]::SelectedIndexProperty) -Index 5 -Name "tabControlIndex"
Set-Binding -Target $computerNameTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 1 -Name "computerName"
Set-Binding -Target $ipAddressTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 3 -Name "ip"
Set-Binding -Target $publicIpAddressTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 4 -Name "publicIpAddress"
Set-Binding -Target $userNameTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 2 -Name "userName"
Set-Binding -Target $advancedInfoTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 6 -Name "advancedInfoTextBox"




$Global:SyncHash = [HashTable]::Synchronized(@{})
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
                $Runspace.Powershell.Dispose()
                
                $Jobs.Remove($Runspace)
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
function Async($scriptBlock){ Start-RunspaceTask $scriptBlock @([PSObject]@{ Name='DataContext' ; Variable=$DataContext},[PSObject]@{Name="State"; Variable=$State})}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name='Jobs' ; Variable=$Jobs })



$Window.ShowDialog()

