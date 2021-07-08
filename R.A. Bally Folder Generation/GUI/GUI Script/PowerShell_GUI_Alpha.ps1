[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$inputXML = @'
<Window x:Class="PowerShell_GUI.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShell_GUI"
        mc:Ignorable="d"
        Title="Window1" Height="450" Width="639" Background="#FF252526">
    <Grid x:Name="Grid_Grid" HorizontalAlignment="Center" VerticalAlignment="Stretch">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto" MinHeight="83"/>
            <RowDefinition Height="Auto" MinHeight="77"/>
            <RowDefinition Height="Auto" MinHeight="60"/>
            <RowDefinition Height="Auto" MinHeight="65"/>
            <RowDefinition Height="Auto" MinHeight="5"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="FirstName_Text" HorizontalAlignment="Left" Margin="52,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="109" Height="18"/>
        <Label x:Name="FirstName_Label" Content="First Name" HorizontalAlignment="Left" Margin="57,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="67"/>
        <Label x:Name="Middle_Label" Content="Middle" HorizontalAlignment="Left" Margin="138,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="48"/>
        <Label x:Name="LastName_Label" Content="Last Name" HorizontalAlignment="Left" Margin="209,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="66"/>
        <Label x:Name="DateOfBirth_Label" Content="Date Of Birth" HorizontalAlignment="Left" Margin="420,43,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="78"/>
        <Label x:Name="Age_Label" Content="Age" HorizontalAlignment="Left" Margin="504,43,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="31"/>
        <Label x:Name="Spouse_Label" Content="Spouse" HorizontalAlignment="Left" Margin="340,43,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="50"/>
        <Label x:Name="Card_Label" Content="Card?" HorizontalAlignment="Left" Margin="532,43,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="40"/>
        <TextBox x:Name="Middle_Text" HorizontalAlignment="Left" Margin="170,71,0,0" TextWrapping="Wrap" Width="16"/>
        <TextBox x:Name="LastName_Text" HorizontalAlignment="Left" Margin="195,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="126" Height="18"/>
        <TextBox x:Name="Spouse_Text" HorizontalAlignment="Left" Margin="335,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="91" Height="18"/>
        <TextBox x:Name="DateOfBirth_Text" HorizontalAlignment="Left" Margin="435,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="81" Height="18"/>
        <TextBox x:Name="Age_Text" HorizontalAlignment="Left" Margin="525,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="21" Height="18"/>
        <TextBox x:Name="Card_Text" HorizontalAlignment="Left" Margin="555,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="26" Height="18"/>
        <TextBox x:Name="StreetAddress_Text" HorizontalAlignment="Left" Margin="10,0,0,0" Grid.Row="1" TextWrapping="Wrap" Width="270" VerticalAlignment="Center" Height="18"/>
        <TextBox x:Name="City_Text" HorizontalAlignment="Left" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="120" Height="18" Margin="305,0,0,0"/>
        <TextBox x:Name="State_Text" HorizontalAlignment="Left" Margin="444,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="30" Height="18"/>
        <TextBox x:Name="ZipCode_Text" HorizontalAlignment="Left" Margin="514,0,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Center" Width="44" Height="18"/>
        <Label x:Name="StreetAddress_Label" Content="Street Address" HorizontalAlignment="Left" Margin="102,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="86"/>
        <Label x:Name="City_Label" Content="City" HorizontalAlignment="Left" Margin="350,6,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="30"/>
        <Label x:Name="State_Label" Content="State" HorizontalAlignment="Left" Margin="441,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="36"/>
        <Label x:Name="ZipCode_Label" Content="Zip Code" HorizontalAlignment="Left" Margin="507,5,0,0" Grid.Row="1" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="58"/>
        <TextBox x:Name="Companies_Text" HorizontalAlignment="Center" Grid.Row="2" TextWrapping="Wrap" VerticalAlignment="Center" Width="562" Height="18"/>
        <Label x:Name="Companies_Label" Content="Companies On File" HorizontalAlignment="Center" Grid.Row="2" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="110" Margin="0,5,0,0"/>
        <TextBox x:Name="ClientFolderPath_Text" HorizontalAlignment="Left" Grid.Row="3" Text="&#xD;&#xA;" TextWrapping="Wrap" VerticalAlignment="Top" Width="512" Height="18" Margin="11,30,0,0"/>
        <Label x:Name="ClientFolderPath_Label" Content="Client File Path" HorizontalAlignment="Center" Grid.Row="3" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="90" Margin="0,4,0,0"/>
        <Button x:Name="OpenClientFolder_Button" Content="Open File" Margin="527,30,0,0" Grid.Row="3" VerticalAlignment="Top" Height="20"/>
        <Button x:Name="SearchClient_Button" Content="Search Client" HorizontalAlignment="Left" Margin="63,16,0,0" VerticalAlignment="Top" Height="20" Width="73"/>
        <Button x:Name="NewClient_Button" Content="New Client" HorizontalAlignment="Left" VerticalAlignment="Top" Height="20" Width="74" Margin="414,16,0,0"/>
        <Button x:Name="ClearForm_Button" Content="Clear Form" HorizontalAlignment="Left" VerticalAlignment="Top" Width="73" Margin="176,16,0,0"/>
        <Button x:Name="UpdateClient_Button" Content="Update Client" HorizontalAlignment="Left" Margin="293,16,0,0" VerticalAlignment="Top" Width="81"/>
        <TextBox x:Name="EntryNumber_Text" HorizontalAlignment="Left" Margin="10,71,0,83" TextWrapping="Wrap" Width="37" Grid.RowSpan="2"/>
        <Label x:Name="EntryNumber_Label" Content="ID" HorizontalAlignment="Left" Margin="18,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC"/>
    </Grid>
</Window>
'@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
    try{
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
    }
    catch{
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        throw
    }
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables

echo "----------------------------------------------------------------------"," "

#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================

# Database Connection
$datasource = "C:\PowerShell Testing\Database\SampleDatabase.accdb"

$access = Connect-Access -DataSource $datasource

# Functions
function Search-Client ([string]$FirstName)
{
    [array]$client = $client = Invoke-Access -Connection $access -Query "SELECT * FROM ClientMasterFile WHERE ([FirstName] = `'$FirstName`');" | select -Property FirstName,Middle,LastName,Spouse,Age,Card,StreetAddress1,City,State,Zipcode,Companies,ClientFolderPath  
    $WPFEntryNumber_Text.Text = $client.EntryNumber
    $WPFFirstName_Text.Text = $client.FirstName
    $WPFLastName_Text.Text = $client.LastName
    $WPFMiddle_Text.Text = $client.Middle
    $WPFSpouse_Text.Text = $client.Spouse
    $WPFAge_Text.Text = $client.Age
    $WPFCard_Text.Text = $client.Card
    $WPFStreetAddress_Text.Text = $client.StreetAddress1
    $WPFCity_Text.Text = $client.City
    $WPFState_Text.Text = $client.State
    $WPFZipCode_Text.Text = $client.Zipcode
    $WPFCompanies_Text.Text = $client.Companies
    $WPFClientFolderPath_Text.Text = $client.ClientFolderPath
    $script:folder = $WPFClientFolderPath_Text.Text
}

function Open-Folder
{  
    Invoke-item $folder
}    

function Clear-Form
{
    $WPFFirstName_Text.Text = $null
    $WPFLastName_Text.Text = $null
    $WPFMiddle_Text.Text = $null
    $WPFSpouse_Text.Text = $null
    $WPFAge_Text.Text = $null
    $WPFCard_Text.Text = $null
    $WPFStreetAddress_Text.Text = $null
    $WPFCity_Text.Text = $null
    $WPFState_Text.Text = $null
    $WPFZipCode_Text.Text = $null
    $WPFCompanies_Text.Text = $null
    $WPFClientFolderPath_Text.Text = $null
    $script:folder = $WPFClientFolderPath_Text.Text
}

$Client_EntryNumber = $WPFEntryNumber_Text.Text
#$Client_FirstName = $WPFFirstName_Text.Text
$Client_LastName = $WPFLastName_Text.Text
#$Client_Middle = $WPFMiddle_Text.Text
$Client_Spouse = $WPFSpouse_Text.Text
#$Client_DateOfBirth = $WPFDateOfBirth_Text.Text
$Client_Card = $WPFCard_Text.Text
$Client_StreetAddress1 = $WPFStreetAddress_Text.Text
$Client_City = $WPFCity_Text.Text
$Client_State = $WPFState_Text.text
$Client_ZipCode = $WPFZipCode_Text.Text
$Client_Companies = $WPFCompanies_Text.Text

function Update-Client
{
    Invoke-Access -Connection $access -Query "UPDATE ClientMasterFile SET [LastName] = '$Client_LastName', [Spouse] = '$Client_Spouse', [Card] = '$Client_Card', [StreetAddress1] = '$Client_StreetAddress1', [City] = '$Client_City', [State] = '$Client_State', [ZipCode] = '$Client_ZipCode', [Companies] = '$Client_Companies' WHERE [EntryNumber] = '$Client_EntryNumber';"
}

# Buttons
$WPFSearchClient_Button.Add_Click({Search-Client -FirstName $WPFFirstName_Text.Text})

$WPFOpenClientFolder_Button.Add_Click({Open-Folder})

$WPFClearForm_Button.Add_Click({Clear-Form})

$WPFFirstName_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -FirstName $WPFFirstName_Text.Text}})

$WPFUpdateClient_Button.Add_Click({Update-Client})

#===========================================================================
# Shows the form
#===========================================================================

#write-host "To show the form, run the following" -ForegroundColor Cyan

$Form.ShowDialog() | Out-Null

Disconnect-Access -Connection $access
