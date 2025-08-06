# 1. Exchange Online 연결
Connect-ExchangeOnline 
# 2. 조직 내 모든 회의실(RoomMailbox) 목록 조회
$roomEmails = Get-Mailbox -RecipientTypeDetails RoomMailbox | Select-Object -ExpandProperty PrimarySmtpAddress

# 3. 설정 시작
foreach ($roomEmail in $roomEmails) {
    Write-Host "`n===== [$roomEmail] 회의실 설정 중 =====" -ForegroundColor Cyan
    $calendarPath = "${roomEmail}:\Calendar"

    # 3-1. Default 사용자 권한 부여 (Reviewer)
    $existingPermission = Get-MailboxFolderPermission -Identity $calendarPath -User Default -ErrorAction SilentlyContinue

    if ($existingPermission) {
        Set-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights Reviewer
        Write-Host "Reviewer 권한 업데이트 완료"
    } else {
        Add-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights Reviewer
        Write-Host "Reviewer 권한 새로 부여 완료"
    }

    # 3-2. RemovePrivateProperty 설정 (비공개 해제)
    Set-CalendarProcessing -Identity $roomEmail -RemovePrivateProperty $true
    Write-Host "RemovePrivateProperty = $true 설정 완료"

    # 3-3. 권한 및 설정 상태 요약
    $perm = Get-MailboxFolderPermission -Identity $calendarPath | Where-Object { $_.User -eq "Default" }
    $calendarProc = Get-CalendarProcessing -Identity $roomEmail
    Write-Host " 현재 권한 상태: $($perm.AccessRights)"
    Write-Host " RemovePrivateProperty 상태: $($calendarProc.RemovePrivateProperty)"
}

# 4. 연결 종료
Disconnect-ExchangeOnline
Write-Host "`n 모든 회의실 설정 완료! 이제 Outlook/Teams에서 제목 + 본문 + 예약자 표시됨!" -ForegroundColor Green
