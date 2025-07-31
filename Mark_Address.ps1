# ============================================
# Misto Malaysia 회의실 주소록 표시 스크립트
# ============================================

# 1. ExchangeOnlineManagement 모듈 설치
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force
}

# 2. Exchange Online 연결 (CLI 환경 고려)
Connect-ExchangeOnline -UseDeviceAuthentication

# 3. 회의실 목록 정의
$rooms = @(
    @{ Name = "room-MAL-2A"; Email = "room-MAL-2A@mistobrand.com" },
    @{ Name = "room-MAL-2B"; Email = "room-MAL-2B@mistobrand.com" },
    @{ Name = "room-MAL-2C"; Email = "room-MAL-2C@mistobrand.com" },
    @{ Name = "room-MAL-1A"; Email = "room-MAL-1A@mistobrand.com" }
)

# 4. 주소록 숨김 해제 설정
foreach ($room in $rooms) {
    Write-Host " Unhiding mailbox: $($room.Email)"
    Set-Mailbox -Identity $room.Email -HiddenFromAddressListsEnabled $false
}

# 5. 연결 해제
Disconnect-ExchangeOnline
