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

    [Parameter(Position=2,Mandatory)]
    [string] $distributionListName
)

$userListFromExcel = Import-Excel $filePath -WorksheetName $worksheetName

foreach ( $user in $userListFromExcel)
{
    New-MailContact -ExternalEmailAddress $user."E-Mail-Adresse" -DisplayName $user.Anzeigename -FirstName $user.Vorname -LastName $user.Nachname 
}