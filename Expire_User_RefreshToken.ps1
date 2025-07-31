Get-AzureADUser -All $true | ForEach-Object { Revoke-AzureADUserAllRefreshToken -ObjectId $_.ObjectId }

