[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$inputXML = @'
<Window x:Class="PowerShell_GUI.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShell_GUI"
        mc:Ignorable="d"
        Title="Window1" Height="450" Width="731" Background="#FF252526">
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
		<Label x:Name="FirstName_Label" Content="First Name" HorizontalAlignment="Left" Margin="73,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="67"/>
		<Label x:Name="Middle_Label" Content="Middle" HorizontalAlignment="Left" Margin="154,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="48"/>
		<Label x:Name="LastName_Label" Content="Last Name" HorizontalAlignment="Left" Margin="225,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="66"/>
		<Label x:Name="DateOfBirth_Label" Content="Date Of Birth" HorizontalAlignment="Left" Margin="436,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="78"/>
		<Label x:Name="Age_Label" Content="Age" HorizontalAlignment="Left" Margin="519,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="31"/>
		<Label x:Name="Spouse_Label" Content="Spouse" HorizontalAlignment="Left" Margin="356,40,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="26" Width="50"/>
		<Label x:Name="Card_Label" Content="Card?" HorizontalAlignment="Left" Margin="548,41,0,0" VerticalAlignment="Top" Foreground="#FF007ACC" Height="24" Width="40"/>
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
		<Label x:Name="EntryNumber_Label" Content="ID" HorizontalAlignment="Left" Margin="18,42,0,0" VerticalAlignment="Top" Foreground="#FF007ACC"/>
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
[System.Collections.ArrayList]$script:datasource = Import-Csv -Path "C:\PowerShell Testing\Database\Test.csv"

function Calculate-Age($person) 
{
    foreach ($client in $person)
    {
        if ($client.DateofBirth -eq "//")
        {
            continue
        }
        else 
        {
        $currentyear = Get-Date | select -ExpandProperty Year
        $dateofbirth = get-date -Date $client.DateOfBirth
        $month = $dateofbirth.Month
        $day = $dateofbirth.Day
    
        $client.birthdaycurrent = 
        @"
        $month/$day/$currentyear
"@
    
        $dateofbirth1 = $client.DateOfBirth
        $birthdaycurrent = $client.birthdaycurrent
    
        $TimeSpan =  [DateTime]$birthdaycurrent - [DateTime]$dateofbirth1;

        $Totaldays = $TimeSpan.TotalDays

        $unformattedage = ($Totaldays/365)

        $age = "{0:n0}" -f $unformattedage
    
        $client.Age = $age
   
        }
    }
    
    $script:datasource | sort { [int]$_.EntryNumber } -Unique |select * | Export-Csv -Path "C:\PowerShell Testing\Database\Test.csv" -NoTypeInformation -Force
}

Calculate-Age($script:datasource)

# Functions
function Search-Client
{
    param (
        [string]$FirstName,
        [string]$LastName
    )
    [array]$client = $client = $datasource | select * | ? {($_.FirstName -eq $FirstName) -xor ($_.LastName -eq $LastName)}
    $WPFEntryNumber_Text.Text = $client.EntryNumber
    $WPFFirstName_Text.Text = $client.FirstName
    $WPFLastName_Text.Text = $client.LastName
    $WPFMiddle_Text.Text = $client.Middle
    $WPFSpouse_Text.Text = $client.Spouse
    $WPFAge_Text.Text = $client.Age
    $WPFCard_Text.Text = $client.Card
    $WPFStreetAddress_Text.Text = $client.StreetAddress
    $WPFCity_Text.Text = $client.City
    $WPFState_Text.Text = $client.State
    $WPFZipCode_Text.Text = $client.Zipcode
    $WPFCompanies_Text.Text = $client.Companies
    $WPFClientFolderPath_Text.Text = $client.ClientFolderPath
    $WPFDateOfBirth_Text.Text = $client.DateOfBirth
}   

    
function Open-Folder
{  
    Invoke-item $folder
}    

function Clear-Form
{
    $WPFEntryNumber_Text.Text = $null
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
    $WPFDateOfBirth_Text.Text = $null
}

function Update-Client
{
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.FirstName = $WPFFirstName_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.LastName = $WPFLastName_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Middle = $WPFMiddle_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Spouse = $WPFSpouse_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Card = $WPFCard_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.StreetAddress = $WPFStreetAddress_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.City = $WPFCity_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.State = $WPFState_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Zipcode = $WPFZipCode_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.Companies = $WPFCompanies_Text.Text}
    $script:datasource | ? EntryNumber -EQ $WPFEntryNumber_Text.Text | foreach {$_.DateOfBirth = $WPFDateOfBirth_Text.Text}

    $script:datasource | sort { [int]$_.EntryNumber } -Unique |select * | Export-Csv -Path "C:\PowerShell Testing\Database\Test.csv" -NoTypeInformation -Force
}

function New-Client
{
    $entrynumber = $script:datasource.Count
    
    $newRow = New-Object PsObject -Property @{ EntryNumber = $entrynumber + 1 ; FirstName = $WPFFirstName_Text.Text ; Middle = $WPFMiddle_Text.Text ; LastName = $WPFLastName_Text.Text ; Spouse = $WPFSpouse_Text.Text ; DateOfBirth = $WPFDateOfBirth_Text.Text; Age = $null ; Card = $WPFCard_Text.Text ; StreetAddress = $WPFStreetAddress_Text.Text ; City = $WPFCity_Text.Text ; State = $WPFState_Text.Text ; ZipCode = $WPFZipCode_Text.Text ; Companies = $WPFCompanies_Text.Text; ClientFolderPath = $null ; birthdaycurrent = $null}
    
    $script:datasource += $newRow

    $script:datasource | sort { [int]$_.EntryNumber } -Unique |select * | Export-Csv -Path "C:\PowerShell Testing\Database\Test.csv" -NoTypeInformation -Force
    
    Calculate-Age($script:datasource)
}
# Buttons
$WPFSearchClient_Button.Add_Click({Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text})

$WPFOpenClientFolder_Button.Add_Click({Open-Folder})

$WPFClearForm_Button.Add_Click({Clear-Form})

$WPFFirstName_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text}})
$WPFLastName_Text.Add_KeyDown({if ($_.Key -eq "Enter"){Search-Client -FirstName $WPFFirstName_Text.Text -LastName $WPFLastName_Text.Text}})

$WPFUpdateClient_Button.Add_Click({Update-Client})

$WPFNewClient_Button.Add_Click({New-Client})

#===========================================================================
# Shows the form
#===========================================================================

#write-host "To show the form, run the following" -ForegroundColor Cyan

$Form.ShowDialog() | Out-Null
