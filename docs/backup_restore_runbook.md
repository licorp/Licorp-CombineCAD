# Backup Restore Runbook

## Scripts

- Backup: `../revit-api-main/ops/backup-mongo.ps1`
- Restore test: `../revit-api-main/ops/restore-mongo-test.ps1`

## Tan suat

- Backup full: hang ngay.
- Restore drill: toi thieu 1 lan/thang.

## Quy trinh backup

1. Dat bien moi truong ket noi MongoDB.
2. Chay script backup.
3. Luu artifact vao kho luu tru an toan.

## Quy trinh restore test

1. Tao database test rieng.
2. Chay script restore vao DB test.
3. Verify so ban ghi cac bang chinh: users/devices/sessions/audit_logs.
4. Chay smoke test auth co ban.

## Mau bien ban drill gan nhat

- Thoi gian: `2026-05-13 09:29 UTC` (moc xac minh smoke test gan nhat)
- Nguoi thuc hien: `Codex + Admin`
- Backup artifact: `TODO: cap nhat sau khi chay backup-mongo.ps1`
- Ket qua restore: `PARTIAL-PASS`
- Ghi chu: `Da xac minh smoke test API auth/admin pass tren backend port 3100; restore DB drill full can thuc hien va dien artifact.`
