# Original conditionalAccessPolicies.csv
$CSVRows = import-csv .\conditionalAccessPolicies.csv
# Multidimensional Array output as CSV for the users,groups,ServicePrincipal.
$IDList = import-csv 'C:\LetsTryAgain2.csv'

# Array to store the changes to write out to a new csv file.
$NewCSVArray = @()

function doWork{
    $todo = $args[0]
    # match on ID or Type_AppID since there are 2 IDs for ServicePrincipal (ID and AppID).
    if (($global:IDList.ID.Contains($todo) -or ($global:IDList.Type_AppID.Contains($todo)))){
        # Find value and the Name
        foreach ($ID in $global:IDList){
            if(($ID.ID -match $todo) -or ($ID.Type_AppID -match $todo)){
                $output = $todo+' '+$ID.DisplayName
            }
        }
    }
    else{$output = $todo + ': not found!'}
    return $output
}

function NullCheck{
    $test = $args[0]
    if ($test -eq ''){$dime = 'Blank'}
    # if it contains data, figure out if it's just 1 guid or multiple. If it's multiple, all that data needs to go back into the same cell.
    elseif($test -match ','){$dime = '';$n = $test.Split(',');foreach ($string in $n){$dime += doWork $string; $dime += ' '}}
    else {$dime = doWork $test;}
    return $dime
}

# looking for the column headings and getting the information from them. Passing to NullCheck to see if the information is empty.
foreach($line in $CSVRows){
    #
    $K = $line.UsersandGroups_included_userIds
    $KResult = NullCheck $K
    $line.UsersandGroups_included_userIds = $KResult # Write back the updated data.
    #
    $L = $line.UsersandGroups_included_groupIds
    $LResult = NullCheck $L
    $line.UsersandGroups_included_groupIds = $LResult
    #
    $M = $line.UsersandGroups_excluded_userIds
    $MResult = NullCheck $M
    $line.UsersandGroups_excluded_userIds = $MResult
    #
    $N = $line.UsersandGroups_excluded_groupIds
    $NResult = NullCheck $N
    $line.UsersandGroups_excluded_groupIds = $NResult
    #
    $T = $line.UsersandGroupsV2_included_userIds
    $TResult = NullCheck $T
    $line.UsersandGroupsV2_included_userIds = $TResult
    #
    $U = $line.UsersandGroupsV2_included_groupIds
    $UResult = NullCheck $U
    $line.UsersandGroupsV2_included_groupIds = $UResult
    #
    $Z = $line.UsersandGroupsV2_excluded_userIds
    $ZResult = NullCheck $Z
    $line.UsersandGroupsV2_excluded_userIds = $ZResult
    #
    $AA = $line.UsersandGroupsV2_excluded_groupIds
    $AAResult = NullCheck $AA
    $line.UsersandGroupsV2_excluded_groupIds = $AAResult
    #
    $AC = $line.ServicePrincipals_included
    $ACResult = NullCheck $AC
    $line.ServicePrincipals_included = $ACResult
    #
    $AD = $line.ServicePrincipals_excluded
    $ADResult = NullCheck $AD
    $line.ServicePrincipals_excluded = $ADResult
    #
    $AI = $line.ServicePrincipalsV2_included
    $AIResult = NullCheck $AI
    $line.ServicePrincipalsV2_included = $AIResult
    #
    $AJ = $line.ServicePrincipalsV2_excluded
    $AJResult = NullCheck $AJ
    $line.ServicePrincipalsV2_excluded = $AJResult
    #
    $AK = $line.ServicePrincipalsV2_includedAppContext
    $AKResult = NullCheck $AK
    $line.ServicePrincipalsV2_includedAppContext = $AKResult
    #
    $NewCSVArray += $line
}
$NewCSVArray | Export-Csv .\conditionalAccessPolicies_Decoded.csv -NoTypeInformation
