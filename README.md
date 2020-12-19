# azure-pwsh
Some sample scripts to manage azure things

## Sync Excel with Global Contacts and Distribution lists
```shell
# Login on to your Exchange-Online subscription
Connect-ExchangeOnline
# Run the Sync-Script
./Sync-DistributionLists.ps1 -filePath <absolute-path-to-excel> -worksheetName <WorksheetName-within-Excel>
```