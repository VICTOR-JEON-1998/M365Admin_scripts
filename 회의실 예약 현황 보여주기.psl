# Exchange Online 연결
Connect-ExchangeOnline 

# 회의실 목록
$roomEmails = @(
    "room-MAL-1A@mistobrand.com",
    "room-MAL-2A@mistobrand.com",
    "room-MAL-2B@mistobrand.com",
    "room-MAL-2C@mistobrand.com"
)

# 각 회의실에 대해 Reviewer 권한 부여
foreach ($roomEmail in $roomEmails) {
    $calendarPath = "$roomEmail:\Calendar"

    # 기존 권한 확인
    $existingPermission = Get-MailboxFolderPermission -Identity $calendarPath -User Default -ErrorAction SilentlyContinue

    if ($existingPermission) {
        # 권한이 있으면 업데이트
        Set-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights Reviewer
        Write-Output "✅ [$roomEmail] 기존 권한을 Reviewer로 업데이트했습니다."
    } else {
        # 권한이 없으면 새로 부여
        Add-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights Reviewer
        Write-Output "✅ [$roomEmail] Reviewer 권한을 새로 부여했습니다."
    }
}


# 권한 제대로 들어갔는지 확인하고 싶다면:
foreach ($roomEmail in $roomEmails) {
    Get-MailboxFolderPermission -Identity "$roomEmail:\Calendar" | Where-Object { $_.User -eq "Default" }
}


# 연결 종료
Disconnect-ExchangeOnline

