# Bring the results of the following commands together:
# Get-AzureRmADServicePrincipal
# Get-AzureRmADGroup
# Get-AzureRmADUser

# Read in Users
$userSpreadsheet = 'C:\TestADUser.xlsx'
# Read in Groups
$groupSpreadsheet = 'C:\TestADGroup.xlsx'
# Read in SPNs
$spnSpreadsheet = 'C:\TestSPN.xlsx'

$MassiveArray = @()

function Closing-Excel {
	Write-Host "Closing out Excel."
	$GroupWorkBook.Close()
	$UserWorkBook.Close()
  $SPNWorkBook.Close()
  $objExcel.Quit()
  [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($objExcel)
}

function UserRead{
    $Name = $UserSheet.Range('B' + $args[0]).Text # Column B
    $Id = $UserSheet.Range('C' + $args[0]).Text # Column C
    Write-Host $Name + $Id + user
    # To do add to array
    $global:MassiveArray += @($Name,$Id,'user')

}

function GroupRead{
    $DisplayName = $groupSheet.Range('C' + $args[0]).Text # Column C
    $Id = $groupSheet.Range('D' + $args[0]).Text # Column D
    Write-Host $DisplayName + $Id + group
    $global:MassiveArray += @($DisplayName,$Id,'group')
}

function SPNRead{
    $DisplayName = $spnSheet.Range('C' + $args[0]).Text # Column C
    $Id = $spnSheet.Range('D' + $args[0]).Text # Column D
    $AppId = $spnSheet.Range('B' + $args[0]).Text # Column B
    Write-Host $DisplayName + $Id + $AppId
    $global:MassiveArray += @($DisplayName,$Id,$AppId)
}

$objExcel = New-Object -ComObject Excel.Application
$objExcel.DisplayAlerts = $false
# User Work
$UserWorkBook = $objExcel.Workbooks.Open($userSpreadsheet)
Write-Host "Opening the User Spreadsheet..."
$UserSheet = $UserWorkBook.sheets.item(1)
$UserRowMax = ($userSheet.UsedRange.Rows).count
Write-Host $UserRowMax
for($UserRow = 2; $UserRow -le $UserRowMax; $UserRow++){
    $user = UserRead $UserRow
}
# Group Work
$GroupWorkBook = $objExcel.Workbooks.Open($groupSpreadsheet)
Write-Host "Opening the Group Spreadsheet..."
$GroupSheet = $GroupWorkBook.sheets.item(1)
$GroupRowMax = ($GroupSheet.UsedRange.Rows).count
for($GroupRow = 2; $GroupRow -le $GroupRowMax; $GroupRow++){
    $group = GroupRead $GroupRow
}
# SPN Work
$SPNWorkBook = $objExcel.Workbooks.Open($spnSpreadsheet)
Write-Host "Opening the SPN Spreadsheet..."
$SPNSheet = $SPNWorkBook.sheets.item(1)
$SPNRowMax = ($SPNSheet.UsedRange.Rows).count
for($SPNRow = 2; $SPNRow -le $SPNRowMax; $SPNRow++){
    $spn = SPNRead $SPNRow
}

Write-Host "========================================================="
$MassiveArray | Export-Csv -Path 'C:\What_did_I_do.csv'

# Should probably close all the excel objects as well.
Closing-Excel

# The above left me with a huge flat file. Which was total junk because it didn't export the data like I thought it would.
# Since the array was still in memory, I used the below to populate the new array:
#$NewArray = @()
#$MAMax = $MassiveArray.count
#for($MARow = 0; $MARow -le $MAMax; $MARow += 3){
#    $0 = $global:MassiveArray[$MARow]
#    $1 = $global:MassiveArray[$MARow+1]
#    $2 = $global:MassiveArray[$MARow+2]
#    $global:NewArray += ,@($0,$1,$2)
#}                      #^ that comma was the missing piece in the original.
# I now had an multidimensional array. Let's export this. Nope.
# =============================================================================================================
# The formatting was all jacked up. I don't have an example of it but I needed to add NoteProperty information.
#$LetsTryAgain = @()
#
#$CNames = @("DisplayName","ID","Type_AppID")
#foreach($line in $NewArray){
#    $obj = New-Object PSObject
#    for ($i=0;$i -lt 3;$i++){
#        $obj | Add-Member -MemberType NoteProperty -Name $CNames[$i] -Value $line[$i]
#    }
#    $LetsTryAgain += $obj
#    $obj=$null
#}
#
#$LetsTryAgain | export-csv .\LetsTryAgain.csv -NoTypeInformation
