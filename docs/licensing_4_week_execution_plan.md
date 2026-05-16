# Ke hoach 4 tuan - Cung co API va bao mat licensing

## Pham vi

Ke hoach nay bam sat cac muc tieu ban yeu cau:

- Kiem soat so thiet bi tren moi license.
- Thu hoi license tu xa va co day du audit.
- Ho tro offline ngan han (grace) nhung khong the bypass vinh vien.
- San sang hon cho quy trinh review Autodesk App Store.

## Hien trang hien tai

Backend trong `../revit-api-main/server.ts` da co:

- Endpoint da version hoa (`/api/v1/auth/*`, `/api/v1/admin/*`).
- Model `Device`, `Session`, `AuditLog`.
- Login tra ve `accessToken`, `refreshToken`, `graceUntil`.
- Verify kiem tra `deviceFingerprint` va rang buoc voi thiet bi da bind.
- Co lockout/rate-limit co ban trong bo nho cho `/login` va `/verify`.
- Co admin API revoke/deactivate va endpoint metrics.

Repo plugin (`src/Licorp_CombineCAD.Shared`) hien chua co module auth/licensing ro rang, vi vay Tuan 2 can duoc xay moi trong lop service dung chung.

## Tuan 1 (Bat buoc - Nen tang bao mat)

### 1) Cung co contract API

- Giu toan bo endpoint auth/admin moi duoi `/api/v1/...`.
- Loai bo cac duong legacy con su dung va tra 410 cho route da deprecate.
- Chuan hoa payload loi:
  - `code` (ma on dinh de may xu ly)
  - `message`
  - `retryAfterSeconds` khi can

### 2) Cung co mo hinh du lieu va tinh ben vung

- `devices`
  - Unique compound index `(userId, fingerprint)`.
  - Index `status`, `lastSeenAt` cho truy van admin.
- `sessions`
  - Unique index cho `refreshTokenHash`.
  - Chien luoc don dep TTL tren `expiresAt` (hoac purge theo lich).
  - Index `(userId, revokedAt)` de revoke nhanh.
- `audit_logs`
  - Index `createdAt`, `action`, `userId`.
  - Chinh sach luu tru (vi du 90-180 ngay tuy compliance).

### 3) Bao ve login/verify

- Thay lockout in-memory bang kho persistent (Mongo collection hoac Redis).
- Khoa theo:
  - `ip + email` cho login.
  - `ip + fingerprint` cho verify.
- Them audit event cho cac lan bi chan do lockout.

### 4) Dieu kien hoan tat (Tuan 1)

- Khong con client dang dung endpoint legacy.
- Lockout persistent van hieu luc sau khi restart server.
- Query plan cho auth/admin su dung index dung ky vong.

## Tuan 2 (Bat buoc - On dinh plugin)

### 1) State machine auth cho plugin

Trien khai luong xac dinh:

1. Verify access token (`/api/v1/auth/verify`).
2. Neu access khong hop le nhung refresh con hop le -> refresh (`/api/v1/auth/refresh`).
3. Neu khong ket noi duoc backend -> cho phep offline den `graceUntil`.

### 2) Luu cache token an toan tren may

- Luu token bundle bang DPAPI (`ProtectedData`) voi CurrentUser scope.
- Them checksum toan ven:
  - `sha256(ciphertext + staticAppSalt + version)`.
- Neu checksum sai -> coi la bi can thiep, bat buoc dang nhap online lai.

### 3) Chuyen doi fingerprint

- Thay MachineId thuong bang dang hash: `fpv1_<hex>`.
- Goi y input (on dinh, han che PII truc tiep):
  - machine SID/device GUID, OS install id, CPU id (neu co).
- Them co che migration:
  - Cache identity cu duoc nang cap sau lan auth online thanh cong dau tien.

### 4) Hop dong trang thai UI

UI phai hien thi dung 4 trang thai:

- `Licensed`
- `Grace mode`
- `Expired`
- `Revoked`

Moi trang thai can co ly do ro rang va huong dan hanh dong tiep theo.

### 5) Dieu kien hoan tat (Tuan 2)

- Plugin khoi dong va danh gia license < 2 giay (warm cache).
- Cache bi sua doi se bi phat hien, khong bypass duoc auth.
- Offline chay duoc trong grace; qua grace se bi tu choi.

## Tuan 3 (Quan trong - Van hanh thuc te)

### 1) Van hanh admin

- Giu/mo rong endpoint:
  - revoke license
  - deactivate device
  - truy van audit log co phan trang/loc
- Ghi audit ro rang cho moi thao tac admin va ket qua.

### 2) Giam sat va canh bao

- Thiet lap toi thieu cac chi so:
  - `5xx rate`
  - `avg/p95 latency`
  - `login fail spike` trong cua so 5 phut
- Gui canh bao email/Slack khi vuot nguong.

### 3) Backup va dien tap restore

- Tai su dung `../revit-api-main/ops/backup-mongo.ps1`.
- Tai su dung `../revit-api-main/ops/restore-mongo-test.ps1`.
- Chay restore test dinh ky va luu ket qua.

### 4) Dieu kien hoan tat (Tuan 3)

- Admin revoke/deactivate co hieu luc o lan verify/refresh tiep theo.
- Alert kich hoat dung trong bai test loi gia lap.
- Restore drill chung minh du lieu co the khoi phuc.

## Tuan 4 (Chuan bi Store)

### 1) Privacy va cong bo du lieu thu thap

- Cong bo privacy policy voi cac truong du lieu cu the:
  - email/account id
  - fingerprint hash
  - IP va timestamp
  - audit actions
- Mo ta retention va chinh sach xoa du lieu.

### 2) Chinh sach versioning API

- Duy tri tuong thich trong pham vi `/api/v1`.
- Chinh sach deprecate vi du:
  - thong bao truoc + giai doan chay song song
  - sunset bang 410 + tai lieu migration

### 3) Test matrix san sang Store

- Doi may/rebind thiet bi.
- Mat mang 3-7 ngay va chuyen trang thai grace.
- Revoke realtime khi plugin dang chay.
- Xoay vong refresh token va chan replay.

### 4) Dieu kien hoan tat (Tuan 4)

- Day du bo tai lieu phuc vu review.
- Co tai lieu chinh sach tuong thich va lo trinh migration.
- Test end-to-end dat va co the lap lai.

## Thu tu trien khai toi uu

1. Cung co backend (lockout/index/audit taxonomy) - phan con thieu cua Tuan 1.
2. Xay module licensing plugin trong `src/Licorp_CombineCAD.Shared/Services`.
3. Noi trang thai licensing vao `ViewModels` va `Views`.
4. Hoan tat tai lieu van hanh, restore drill, privacy/versioning policy.

## Checklist deliverables

- [x] Patch hardening backend auth/admin
- [x] Service licensing plugin + secure cache
- [x] Tich hop UI state (Licensed/Grace/Expired/Revoked)
- [x] Ghi chu cau hinh monitoring va alert
- [x] Runbook backup/restore + ket qua drill gan nhat
- [x] Tai lieu privacy + API versioning policy
- [x] Bang chung test Store-readiness (muc PARTIAL + TODO evidence)

## Tien do thuc te (cap nhat)

### Da xong (xac minh)

- Backend `/api/v1/...` da hoat dong va smoke test pass tren port rieng (3100).
- Auth flow backend da co `login`, `verify`, `refresh` va tra token bundle gom `accessToken`, `refreshToken`, `graceUntil`.
- Admin API da co `metrics`, `audit-logs`, `revoke-license`, `deactivate-device`.
- Plugin da co `LicenseAuthService` voi flow `verify -> refresh -> offline grace` va cache token bang DPAPI + checksum integrity.
- Fingerprint da dung format `fpv1_...`.
- UI da map trang thai `Licensed`, `Grace mode`, `Expired`, `Revoked` (qua banner + `LicenseStateText`).

### Chua xong (uu tien tiep)

- W3/W4: Da co tai lieu baseline. Con thieu evidence thuc te de nang PARTIAL -> PASS.
- Store readiness: can thu thap screenshot UI + logs + audit records cho SR-01..SR-04.
- Backup/restore: can dien backup artifact va ket qua restore drill full.

## Ke hoach tiep theo (thuc thi)

1. Thu thap evidence thuc te SR-01..SR-04 de chot PASS.
2. Chay backup/restore drill full va dien artifact vao runbook.
3. Chot bo tai lieu nop review (privacy + API policy + monitoring + test evidence).
