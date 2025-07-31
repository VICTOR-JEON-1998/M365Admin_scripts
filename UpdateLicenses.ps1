# Import-Excel 모듈이 설치되어 있어야 합니다.
# 엑셀에서 사용자 목록을 가져옵니다.
$excelFilePath = "test.xlsx"  # 엑셀 파일 경로
$userList = Import-Excel -Path $excelFilePath

# Azure AD, Exchange Online, Microsoft Graph에 연결 (이미 연결되어 있다고 가정)
# Connect-AzureAD
# Connect-ExchangeOnline
# Connect-MgGraph

# Get-MgSubscribedSku로 테넌트에서 사용할 수 있는 SKU 목록을 가져옵니다.
$subscribedSkus = Get-MgSubscribedSku

# 라이선스 SKU ID 매핑 (스크립트에 사용할 SKU 이름을 이곳에 정의)
$licenseSkuId = @{
    "Office 365 E1"               = "18181a46-0d4e-45cd-891e-60aabd171b4e"
    "Office 365 E3"               = "6fd2c87f-b296-42f0-b197-1e91e994b900"
    "Microsoft 365 Business Basic"= "3b555118-da6a-4418-894f-7df1e2096870"
    "Microsoft 365 Business Standard" = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"
    "Microsoft 365 Business Premium" = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
    "Microsoft 365 E3"            = "05e9a617-0261-4cee-bb44-138d3ef5d965"
    "Azure Active Directory Premium P1"       = "078d2b04-f1bd-4111-bbd4-b4b1b354cef4"
    "Intune"                      = "061f9ace-7d42-4136-88ac-31dc755f143f"
}

# 각 사용자에 대해 라이선스를 할당
foreach ($user in $userList) {
    $userId = $user.UserID
    Write-Host "UserID: $userId - 라이선스 할당 시작"
    $userLicenses = $user.Licenses -split "\+\s*"  # '+'로 구분된 라이선스 처리
    
    # 라이선스를 추가할 SKU ID 배열
    $addLicenses = @()
    
    # 사용자에게 부여할 라이선스 SKU ID 처리
    foreach ($license in $userLicenses) {
        $license = $license.Trim()  # 공백 제거
        Write-Host "Trim 완료: $license"
        
        # SKU ID가 매핑되어 있는지 확인
        if ($licenseSkuId.ContainsKey($license)) {
            $skuId = $licenseSkuId[$license]
            $addLicenses += @{ SkuId = $skuId }
            Write-Host "매핑 확인 완료: $license -> $skuId"
        } else {
            Write-Host "알 수 없는 라이선스 제품: $license"
        }
    }

    # Set-MgUserLicense로 기존 라이선스 제거 및 새 라이선스 추가
    try {
        # 사용자에게 라이선스 추가 및 제거
        Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses @()
        Write-Host "사용자 $userId 에게 라이선스 할당 완료"
    } catch {
        Write-Host "사용자 $userId 에게 라이선스 할당 실패: $_"
    }
}

Write-Host "All licenses have been updated. Finish!"
