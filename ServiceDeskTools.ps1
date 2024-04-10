Add-Type -AssemblyName PresentationFramework

$xamlFile="C:\Temp\MainWindow.xaml"
$inputXAML=Get-Content -Path $xamlFile -Raw
$inputXAML=$inputXAML -replace 'mc:Ignorable="d"', '' -replace "x:N","N" -replace '^<Win.*', '<Window'
[XML]$XAML=$inputXAML

$reader = New-Object System.Xml.XmlNodeReader $XAML

try{
$psform=[Windows.Markup.XamlReader]::Load($reader)
}
catch {
Write-Host $_.Exception
throw
}

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try{
        Set-Variable -Name "var_$($_.Name)" -Value $psform.FindName($_.Name) -ErrorAction Stop #useful as it allows a dynamic setting of variables
    }
    catch {
        throw
    }
}


#EXPORT BUTTON DEFAULT HANDLER
$var_exportBox.IsEnabled = $false
$var_outputBox.Add_SelectionChanged({
    # Check if any item is selected
    if ($var_outputBox.SelectedItem -ne $null) {
        # Enable the export button if an item is selected
        $var_exportBox.IsEnabled = $true
    } else {
        # Disable the export button if no item is selected
        $var_exportBox.IsEnabled = $false
    }
})

#SCRIPT STATUS INDICATOR
function UpdateStatusLabel{
    param([Boolean]$status)

    if ($status) {
    $var_runningLabel.Visibility = "Visible"
    }
    else {
    $var_runningLabel.Visibility = "Hidden"}
}

UpdateStatusLabel -status $false

#DATA
$data = [System.Collections.ArrayList]@()

#EXPORT FUNCTIONALITY
$var_exportBox.Add_Click({
UpdateStatusLabel -status $true
$var_runningLabel.Visibility = "Visible"
[System.Windows.Forms.MessageBox]::Show("Click 'Ok' to run the script...", "Script Running", "OK", "Information")
write-host $var_runningLabel
Start-Sleep -Milliseconds 100
$uniqueEmails = @{}

$selectedItem = $var_outputBox.SelectedItem
write-host $selectedItem
$data.Clear()

Get-NestedGroupMembers -group $selectedItem

#DESTINATION PATH PROMPT
    $saveFileDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV File"
    $result = $saveFileDialog.ShowDialog()

   if ($result -eq "OK") {
        $filePath = $saveFileDialog.FileName
$data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
[System.Windows.MessageBox]::Show("CSV file exported successfully to:`n$filePath", "Export Successful", "OK", "Information")}
elseif ($result -eq "Cancel") {
    [System.Windows.MessageBox]::Show("Export cancelled.", "Export Cancelled", "OK", "Warning")
}

UpdateStatusLabel -status $false

}

)

#SUB GROUP FUNCTION

function Get-NestedGroupMembers {
param (
        [string]$group
    )
    $groupMembers = Get-ADGroupMember -Identity $group | Where-Object { $_.objectClass -eq 'user' }

    Write-Host "Processing group: $group" -ForegroundColor Yellow

    foreach ($user in $groupMembers) {
        $userObject = Get-ADUser -Identity $user -Properties GivenName, Surname, UserPrincipalName, Office, info, Title, Department
        $userGroup = $group

        $upn = $userObject.UserPrincipalName
        if ([string]::IsNullOrEmpty($upn)) {
            $upn = "N/A"
        }

        #Check if the email address is unique
        if ($upn -ne "N/A" -and -not $uniqueEmails.ContainsKey($upn)) {
            $uniqueEmails[$upn] = $true

            $data.Add([PSCustomObject]@{
                FirstName = $userObject.GivenName
                LastName = $userObject.Surname
                Email = $upn
                Office = $userObject.Office
                Title = $userObject.Title
                Department = $userObject.Department
            }) | Out-Null
        

        }
    }

    #Recursively process nested groups
    $nestedGroups = Get-ADGroupMember -Identity $group | Where-Object { $_.objectClass -eq 'group' }

    foreach ($nestedGroup in $nestedGroups) {
        Get-NestedGroupMembers -group $nestedGroup
    }
}

#SEARCH LOGIC
$var_searchButton.Add_Click({
    $var_outputBox.Items.Clear()
    Get-ADGroupFunction -groupSearch *$($var_inputBox.Text)*
    
})
#LIST BOX FUNCTIONALITY 
function Get-ADGroupFunction {
param ([string]$groupSearch) 
$listOfGroups = Get-ADGroup -Filter {Name -like $groupSearch}  -Properties Name, SamAccountName | Select-Object -first 10
foreach ($groupName in $listOfGroups) {
$var_outputBox.Items.Add($groupName.SamAccountName)
}
}
$psform.ShowDialog()
