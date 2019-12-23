param(
    [Parameter(Position=0,Mandatory)]
    [string]$excelPath,

    [Parameter(Position=1,Mandatory)]
    [string]$property
)

$excelData = Import-Excel $excelPath
$srcMail = $excelData | ForEach-Object { if( $_.$property ) { $_.$property } } | ForEach-Object { $_.Trim() }

$dstMail = Get-MailContact | ForEach-Object { $_.WindowsEmailAddress.Trim() }

Compare-Object -ReferenceObject $srcMail -DifferenceObject $dstMail