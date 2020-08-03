param (
    [Parameter(Position=0,Mandatory)]
    [ValidateScript({
        if(-Not ($_ | Test-Path) ){
            throw "File or folder does not exist"
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "The Path argument must be a file. Folder paths are not allowed."
        }
        return $true 
    })]
    [System.IO.FileInfo] $filePath,

    [Parameter(Position=1,Mandatory)]
    [string] $worksheetName,

    [Parameter(Position=2)]
    [bool] $updateUsers
)

# Small function to validate mail addresses
function IsValidEmail { 
    param([string]$EmailAddress)

    try {
        $null = [mailaddress]$EmailAddress
        return $true
    }
    catch {
        return $false
    }
}

###
# First step is to import all contacts existing online and from the single-source-of-truth-excel
# Then we compare both and delete/add/update each contact
###

# Get entries from excel
$userListFromExcel = Import-Excel $filePath -WorksheetName $worksheetName
$excelMails = $userListFromExcel | Select-Object -ExpandProperty "E-Mail-Adresse" -Unique

# Get existing online contacts
$onlineContacts = Get-MailContact
$contactMails = $onlineContacts | Select-Object -ExpandProperty WindowsEmailAddress -Unique

# Compare the email addresses and iterate over the result.
# If mail exists online but not in the excel, delete the online entry
# If mail exists in excel but not online, create new online entry
# If mail exists in both, update the online entry with the values from excel
$comparisonResult = Compare-Object -ReferenceObject $excelMails -DifferenceObject $contactMails -IncludeEqual

# Print result of compare
$comparisonResult
$cCount = $comparisonResult.Count

# Get all distribution lists that shall be filled with contacts
# by getting the excel column names that can be used as mail addresses
$distributionlistAddresses = $userListFromExcel[0] | Get-Member | Where-Object { IsValidEmail $_.Name } | Select-Object -ExpandProperty Name

$distributionlistAddresses
$dlCount = $distributionlistAddresses.Count
"$cCount contacts to be synchronised and added to $dlCount distributionlists ..."

foreach ( $result in $comparisonResult)
{
    # Skip all empty mailaddresses (null values)
    if ($result.InputObject -eq "") {
        continue
    }

    $mailAddress = $result.InputObject
    $excelEntry  = $userListFromExcel | Where-Object {$_."E-Mail-Adresse" -eq $mailAddress}

    if ( $result.SideIndicator -eq '==' -and $updateUsers) {
        "Updating $mailAddress ..."
        Set-MailContact -Identity $mailAddress.Trim() -DisplayName $excelEntry."Anzeigename".Trim()
    }

    if ( $result.SideIndicator -eq '=>' ) {
        "Deleting $mailAddress ..."
        Remove-MailContact -Identity $mailAddress.Trim() #-Confirm:$false
    }

    if ( $result.SideIndicator -eq '<=' ) {
        "Adding $mailAddress ..."
        New-MailContact -Name $excelEntry."Anzeigename".Trim() -ExternalEmailAddress $mailAddress.Trim() -DisplayName $excelEntry."Anzeigename".Trim()
    }
}

###
# Second part is to move the contacts in their specific distribution lists
###

# Foreach distribution list as list
#   Step#1 Check if given list exists
#   Step#2 Get all members of list
#   Step#3 Compare members as described by excel with current online members
#   Step#4 Work through compare result by adding/removing members from list
foreach ( $listName in $distributionlistAddresses )
{
    # Check if given list exists by getting the online object
    $list = Get-DistributionGroup -Identity $listName
    # If list does not exist, continue with the next item.
    # An error will be displayed anyway on the shell
    if ( !$list )
    {
        "Skipping distribution list $listName because it does not exist!"
        continue
    }

    # Get all members of list, that are of type MailContact (we do not want to modify nested distribution groups)
    $listmembers = Get-DistributionGroupMember -Identity $listName | Where-Object { $_.RecipientType -eq "MailContact" }

    # Get all contacts from excel that should be in the given list
    $futureMembers = $userListFromExcel | Where-Object { $_.$listName }

    # Compare current online members with given members from list by their mailaddresses
    $comparisonResult = Compare-Object -ReferenceObject ($futureMembers | Select-Object -ExpandProperty "E-Mail-Adresse" -Unique) -DifferenceObject ($listmembers | Select-Object -ExpandProperty PrimarySmtpAddress -Unique)

    foreach ( $diff in $comparisonResult )
    {
        # Skip empty mailaddresses
        if ( !$diff.InputObject ) {
            continue
        }

        $contact = Get-MailContact -Identity $diff.InputObject

        if ( $diff.SideIndicator -eq '=>' ) {
            "[$listName] Removing $( $diff.InputObject ) from group ..."
            Remove-DistributionGroupMember -Identity $listName -Member $contact.DistinguishedName
        }
    
        if ( $diff.SideIndicator -eq '<=' ) {
            "[$listName] Adding $( $diff.InputObject ) to group ..."
            Add-DistributionGroupMember -Identity $listName -Member $contact.DistinguishedName
        }
    }

}
