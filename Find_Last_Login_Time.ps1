# 모듈 설치
Install-Module -Name AzureAD
Install-Module -Name AzureADPreview

# AzureADPreview 모듈 최신버전 업데이트
Update-Module -Name AzureADPreview

# 모듈 주입
Import-Module AzureADPreview


# Microsoft 365에 연결합니다.
Connect-AzureAD

# 사용자의 마지막 로그인 시간을 가져오기 위한 쿼리 실행
Get-AzureADAuditSignInLogs -Filter "UserPrincipalName eq '[MailAddress]'" | Sort-Object -Property CreatedDateTime -Descending | Select-Object -First 1 | Select UserPrincipalName, CreatedDateTime
