# Azure 연결
Connect-AzureAD

$users = Import-Csv -Path "D:\remove.csv" # 경로 설정
foreach ($user in $users) {
    $LicensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    # $LicensesToRemove.AddLicenses = @()
    $LicensesToRemove.RemoveLicenses = (Get-AzureADUser -ObjectId $user.Email).AssignedLicenses | ForEach-Object { $_.SkuId }
    Set-AzureADUserLicense -ObjectId $user.Email -AssignedLicenses $LicensesToRemove
}

