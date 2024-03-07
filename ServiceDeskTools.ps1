Add-Type -AssemblyName PresentationFramework

$xamlFile="C:\Users\kmorrison\Development\Service Desk Tools\MainWindow.xaml"
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

#----------------------------------Plans--------------------------------------------#

<#
    Export only runs when file location is selected?
    Perhaps have a running screen?
#>


#----------------------------------Search logic-------------------------------------#
$var_searchButton.Add_Click({
    $var_outputBox.Items.Clear()
    Get-ADGroupFunction -groupSearch *$($var_inputBox.Text)*
    
})
#----------------------------------DATA---------------------------------------------#
$data = [System.Collections.ArrayList]@()

#----------------------------------EXPORT-------------------------------------------#
$var_exportBox.Add_Click({
$uniqueEmails = @{}

$selectedItem = $var_outputBox.SelectedItem
write-host $selectedItem
$data.Clear()

Get-NestedGroupMembers -group $selectedItem

 # Prompt the user to select the destination path for the CSV file
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
##
}

)

# ---UPDATE---
function Get-ADGroupFunction {
param ([string]$groupSearch) 
#Get-ADGroup -Filter {Name -like $groupSearch}  -Properties * | select -property SamAccountName,Name
$listOfGroups = Get-ADGroup -Filter {Name -like $groupSearch}  -Properties * | Select-Object -first 10 -expandproperty Name #limits it to the first 10 -> currently it doesn't search for the first x results
foreach ($groupName in $listOfGroups) {
$var_outputBox.Items.Add($groupName)
write-host $groupName

}
}
#Subgroup colours can be different from parent group, e.g. IT Services L1/L2 will have users of a random colour & they aren't the same

#$checkbox.add_Indeterminate({ Write-Host 'indeterminate' })
#Get-Service | ForEach-Object {$var_outputBox.Items.Add($_.Name)}

$psform.ShowDialog()

#-------------------------------------------SUB GROUP FUNCTION--------------------------------------#

function Get-NestedGroupMembers {
    param (
        [string]$group,
        [boolean]$londonFilter,
        [boolean]$liverpoolFilter,
        [boolean]$cambridgeFilter
    )
    $groupMembers = Get-ADGroupMember -Identity $group | Where-Object { $_.objectClass -eq 'user' }

    Write-Host "Processing group: $group" -ForegroundColor Yellow

    foreach ($user in $groupMembers) {
        $userObject = Get-ADUser -Identity $user -Properties GivenName, Surname, UserPrincipalName, Office, info, Title, Department
        $userGroup = $group
        write-host $userObject.Office


##
        $upn = $userObject.UserPrincipalName
        if ([string]::IsNullOrEmpty($upn)) {
            $upn = "N/A"
        }

        # Check if the email address is unique
        if ($upn -ne "N/A" -and -not $uniqueEmails.ContainsKey($upn)) {
            $uniqueEmails[$upn] = $true

            $data.Add([PSCustomObject]@{
                #Group = $userGroup
                FirstName = $userObject.GivenName
                LastName = $userObject.Surname
                Email = $upn
                #PAs = $userObject.info
                Office = $userObject.Office
                Title = $userObject.Title
                Department = $userObject.Department
            }) | Out-Null
        

        }
    }

    # Recursively process nested groups
    $nestedGroups = Get-ADGroupMember -Identity $group | Where-Object { $_.objectClass -eq 'group' }

    foreach ($nestedGroup in $nestedGroups) {
        Get-NestedGroupMembers -group $nestedGroup
    }
}
#----------------------------------------------------------------------------------------------------#
# Make the UI better & a loading screen
