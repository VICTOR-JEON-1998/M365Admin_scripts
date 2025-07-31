# ============================================
# Misto Malaysia (Surian Tower) 회의실 등록 스크립트
# 실행 전 관리자 권한 PowerShell에서 실행 요망
# ============================================

# 1. ExchangeOnlineManagement 모듈 설치
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force
}

# 2. Exchange Online 연결
Connect-ExchangeOnline -UseDeviceAuthentication

# 3. RoomList Distribution Group 생성
$roomListName = "Misto Malaysia (Surian Tower)"
$roomListAddress = "Room-MistoMalaysia-SurianTower@mistobrand.com"

New-DistributionGroup -Name $roomListName `
    -DisplayName $roomListName `
    -PrimarySmtpAddress $roomListAddress `
    -RoomList

# 4. 회의실 메일박스 생성
$rooms = @(
    @{ Name = "room-MAL-2A"; Display = "(2F) Meeting RM, 2A"; Email = "room-MAL-2A@mistobrand.com"; Floor = 2 },
    @{ Name = "room-MAL-2B"; Display = "(2F) Meeting RM, 2B"; Email = "room-MAL-2B@mistobrand.com"; Floor = 2 },
    @{ Name = "room-MAL-2C"; Display = "(2F) Meeting RM, 2C"; Email = "room-MAL-2C@mistobrand.com"; Floor = 2 },
    @{ Name = "room-MAL-1A"; Display = "(1F) Convention RM, 1A"; Email = "room-MAL-1A@mistobrand.com"; Floor = 1 }
)

foreach ($room in $rooms) {
    New-Mailbox -Name $room.Name -DisplayName $room.Display -Room -PrimarySmtpAddress $room.Email
}

# 5. 회의실을 RoomList에 추가
foreach ($room in $rooms) {
    Add-DistributionGroupMember -Identity $roomListName -Member $room.Email
}

# 6. Set-Place: 위치 정보 등록
$building = "Surian Tower, Unit 2.4, No 1, Jalan PJU 7/3, Mutiara Damansara"
$city = "Petaling Jaya"
$state = "Selangor"
$country = "Malaysia"

foreach ($room in $rooms) {
    Set-Place -Identity $room.Email `
        -City $city `
        -Building $building `
        -Floor $room.Floor `
        -State $state `
        -CountryOrRegion $country
}

# 7. (선택) 자동 예약 처리 설정
foreach ($room in $rooms) {
    Set-CalendarProcessing -Identity $room.Email -AutomateProcessing AutoAccept
}

# 8. 연결 해제
Disconnect-ExchangeOnline
