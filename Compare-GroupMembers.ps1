param(
    [Parameter(Position=0,Mandatory)]
    [string]$excelPath,

    [Parameter(Position=1,Mandatory)]
    [string]$excelGroupName,

    [Parameter(Position=2,Mandatory)]
    [string]$azureGroupName
)

$excelData = Import-Excel $excelPath
$srcGroupMembers = $excelData | Where-Object -Property $excelGroupName -EQ "x"
$srcMail = $srcGroupMembers | ForEach-Object { if( $_.$property ) { $_.$property } } | ForEach-Object { $_.Trim() }

$dstMail = Get-DistributionGroupMember -Identity $azureGroupName | ForEach-Object { $_.PrimarySmtpAddress }

Compare-Object -ReferenceObject $srcMail -DifferenceObject $dstMail