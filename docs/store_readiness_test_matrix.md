# Store Readiness Test Matrix

## Muc tieu

Xac nhan plugin licensing hoat dong dung trong cac tinh huong review thuc te truoc khi nop Autodesk App Store.

## Test cases

| ID | Scenario | Buoc test | Ky vong |
|---|---|---|---|
| SR-01 | Doi may / bind thiet bi moi | Dang nhap tren thiet bi A, sau do dang nhap tren thiet bi B | Vuot qua so thiet bi cho phep se bi chan; audit co event lien quan |
| SR-02 | Mat mang 3-7 ngay | Dang nhap thanh cong, ngat mang, mo lai plugin trong cua so grace | Plugin vao `Grace mode`, het grace thi `Expired` |
| SR-03 | Revoke realtime | Admin revoke license khi plugin dang chay | Lan verify/refresh tiep theo tra `Revoked`, UI cap nhat dung |
| SR-04 | Refresh token replay | Thu dung lai refresh token da revoke/het han | API tu choi, khong cap access token moi |

## Evidence can thu thap

- Screenshot UI trang thai (`Licensed`, `Grace mode`, `Expired`, `Revoked`).
- Log request/response API (an thong tin nhay cam).
- Audit log records tuong ung tung scenario.
- Thoi gian va nguoi thuc hien test.

## Mau ket qua

| ID | Ket qua | Bang chung | Ghi chu |
|---|---|---|---|
| SR-01 | PARTIAL | API smoke test + audit log paging da xac minh | Can test scenario doi may thuc te tren 2 may de chot PASS |
| SR-02 | PARTIAL | Plugin co flow offline grace trong `LicenseAuthService` | Can test mat mang 3-7 ngay thuc te va chup UI state |
| SR-03 | PARTIAL | Backend co revoke endpoint + verify tra revoked path | Can chay test realtime khi plugin dang mo |
| SR-04 | PARTIAL | Backend refresh/session revoke flow da co | Can bo sung test replay token co log bang chung |
