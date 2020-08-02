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
    [string] $worksheetName
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

    if ( $result.SideIndicator -eq '==' ) {
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

