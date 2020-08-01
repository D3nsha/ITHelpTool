#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="800" Height="800">
<Grid>
  <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
          <Border Grid.Column="0" Grid.Row="0" Height="35" Padding="5" Background="#FFE31E37">
            <TextBlock TextWrapping="Wrap" Text="{Binding titleText}" Name="titleTextBlock" Foreground="#ffffff"/>
        </Border>
  <TabControl Grid.Row="1" SelectedIndex="{Binding tabControlIndex}" TabStripPlacement="Left" Name="kdc7ha079mv7v">
            <TabItem>
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
                        <TextBox Name="computerNameTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding computerName}" Grid.Row="0" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="ipAddressTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="IP Address(es):" Grid.Row="1" Grid.Column="0" Margin="10"/>
                        <TextBox Name="ipAddressTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding ip}" Grid.Row="1" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="publicIpAddressTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Public IP Address:" Grid.Row="2" Grid.Column="0" Margin="10"/>
                        <TextBox Name="publicIpAddressTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding publicIpAddress}" Grid.Row="2" Grid.Column="1" Margin="10"/>

                        <TextBlock Name="userNameTextBlock" HorizontalAlignment="Left" VerticalAlignment="Top" Text="Username:" Grid.Row="3" Grid.Column="0" Margin="10"/>
                        <TextBox Name="userNameTextBox" VerticalAlignment="Top" TextWrapping="Wrap" Text="{Binding userName}" Grid.Row="3" Grid.Column="1" Margin="10"/>

                        <Button Content="Advanced info" Grid.Row="4" Margin="10" Name="kdc7ha0970tgw"/>

                    </Grid>

                    <!-- Fix Addin buttons-->
                    <Border Grid.Row="3" Padding="10" BorderBrush="DodgerBlue" BorderThickness="1" Margin="10">
                        <StackPanel>
                            <TextBlock Text="Fix Office add-ins" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="Bold" Margin="0,0,0,5"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>
  </TabControl>
 </Grid></Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#Write your code here

#region +++++++++++++++++++++ Buttons +++++++++++++++++++++ 
function Reset-Word {
    Reset-Addins("Word")
}

function Reset-Outlook {
    Reset-Addins("Bum")
}

function Reset-Powerpoint {
    [Windows.Forms.MessageBox]::Show("Reset Powerpoint")
}

function Reset-Excel {
    [Windows.Forms.MessageBox]::Show("Reset Excel")
}

function Copy-Clipboard {
    Write-Output $State.advancedInfoTextBox | & "$env:windir\system32\clip.exe"
}

function Go-Back {
    $State.tabControlIndex = "0"   
}

function Show-Advanced {
    $State.tabControlIndex = "1"
    
    Async {
        #Gather-AdvancedData   
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


$kdc7ha0970tgw.Add_Click({Show-Advanced $this $_})

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
    "titleText":"IT Help Tool",
    "computerName":"computername",
    "ip":"127.0.0.1",
    "publicIpAddress":"127.0.0.1",
    "userName":"username"
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("titleText","computerName","ip","publicIpAddress","userName") 

$Window.DataContext = $DataContext
Set-Binding -Target $titleTextBlock -Property $([System.Windows.Controls.TextBlock]::TextProperty) -Index 0 -Name "titleText"

Set-Binding -Target $computerNameTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 1 -Name "computerName"
Set-Binding -Target $ipAddressTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 2 -Name "ip"
Set-Binding -Target $publicIpAddressTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 3 -Name "publicIpAddress"
Set-Binding -Target $userNameTextBox -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 4 -Name "userName"
$Window.ShowDialog()


