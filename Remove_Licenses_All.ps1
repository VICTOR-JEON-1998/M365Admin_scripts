# Import-Excel 모듈을 사용하여 엑셀에서 사용자 목록 가져오기
Import-Module ImportExcel

$excelFilePath = "test.xlsx"  # 엑셀 파일 경로
$userList = Import-Excel -Path $excelFilePath

# 라이선스 SKU ID 매핑 (예시)
$licenseSkuId = @{
    "Office 365 E1"               = "18181a46-0d4e-45cd-891e-60aabd171b4e"
    "Office 365 E3"               = "6fd2c87f-b296-42f0-b197-1e91e994b900"
    "Microsoft 365 Business Basic"= "3b555118-da6a-4418-894f-7df1e2096870"
    "Microsoft 365 Business Standard" = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"
    "Microsoft 365 Business Premium" = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
    "Microsoft 365 E3"            = "05e9a617-0261-4cee-bb44-138d3ef5d965"
    "Azure Active Directory Premium P1" = "078d2b04-f1bd-4111-bbd4-b4b1b354cef4"
    "Intune"                      = "061f9ace-7d42-4136-88ac-31dc755f143f"
}

# Azure AD에 연결
Connect-MgGraph

# 각 사용자에 대해 라이선스를 제거
foreach ($user in $userList) {
    $userId = $user.UserId
    $licensesToRemove = $user.LicensesToRemove -split "\+\s*"  # '+'로 구분된 라이선스 목록 처리

    # 라이선스 제거 배열 초기화
    $removeLicenses = @()

    # 삭제할 라이선스 SKU ID를 배열로 변환
    foreach ($license in $licensesToRemove) {
        $license = $license.Trim()  # 공백 제거
        Write-Host "Trim 완료: $license"
        
        # 라이선스 SKU ID가 매핑되어 있는지 확인
        if ($licenseSkuId.ContainsKey($license)) {
            $skuId = $licenseSkuId[$license]
            # GUID만 배열에 추가
            $removeLicenses += [Guid]$skuId
            Write-Host "매핑 확인 완료 (라이선스): $license -> $skuId"
        } else {
            Write-Host "알 수 없는 라이선스 제품: $license"
        }
    }

    # 라이선스를 제거하는 명령어 실행
    if ($removeLicenses.Count -gt 0) {
        try {
            # 라이선스를 제거 (빈 배열로 addLicenses 전달)
            Set-MgUserLicense -UserId $userId -AddLicenses @() -RemoveLicenses $removeLicenses
            Write-Host "$userId 에게서 라이선스를 제거 완료"
        } catch {
            Write-Host "사용자 $userId 에게서 라이선스를 제거하는데 실패: $_"
        }
    } else {
        Write-Host "사용자 $userId 에게 제거할 라이선스가 없음"
    }
}

Write-Host "모든 라이선스 제거 완료"
