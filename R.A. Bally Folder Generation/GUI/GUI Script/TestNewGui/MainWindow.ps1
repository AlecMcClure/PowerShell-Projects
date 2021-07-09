function LoadXamlForm
{

    Param
    (
        [Parameter(Mandatory=$true)] $xamlFile
    )

    Add-Type -AssemblyName PresentationFramework

    $inputXML = Get-Content -Path $xamlFile


    # https://stackoverflow.com/a/52416973/1644202

    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window' -replace "x:Bind", "Binding"
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML
    #Read XAML

        $reader=(New-Object System.Xml.XmlNodeReader $XAML) 
    try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
    catch [System.Management.Automation.MethodInvocationException] {
        Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
        write-host $error[0].Exception.Message -ForegroundColor Red
        if ($error[0].Exception.Message -like "*button*"){
            write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"}
    }
    catch{#if it broke some other way <span class="wp-smiley wp-emoji wp-emoji-bigsmile" title=":D">:D</span>
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
            }

    #===========================================================================
    # Store Form Objects In PowerShell
    #===========================================================================
    $Elements = @{}
    $xaml.SelectNodes("//*[@Name]") | %{ $Elements[$_.Name] = $Form.FindName($_.Name) }

    return $Form, $Elements
}




$Form,$Elements = LoadXamlForm "C:\Users\alecb\Documents\GitHub\PowerShell-Projects\R.A. Bally Folder Generation\GUI\GUI Script\TestNewGui\XAML\MainWindow.xaml"

$WPFNewClient_Button.Add_Click({LoadXamlForm "C:\Users\alecb\Documents\GitHub\PowerShell-Projects\R.A. Bally Folder Generation\GUI\GUI Script\TestNewGui\XAML\ClientSelect.xaml"})
#$Elements.lvApps.AddChild([pscustomobject]@{Name='Ben';Description="sdsd 343"})

$Form.ShowDialog() | out-null