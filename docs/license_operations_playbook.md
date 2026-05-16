# License Operations Playbook

## Dieu kien tien quyet

- Backend licensing dang chay (vi du `http://localhost:3100`).
- Co admin key/secret theo cau hinh server.

## 1) Kiem tra suc khoe he thong

```bash
curl -s "http://localhost:3100/api/v1/admin/metrics"
```

Kiem tra nhanh:
- `system.healthy`
- `system.loginFailSpike`
- `system.avgLatencyMs`, `system.p95LatencyMs`

## 2) Tra audit log theo thoi gian

```bash
curl -s "http://localhost:3100/api/v1/admin/audit-logs?page=1&limit=50"
```

Nen loc cac action quan trong:
- `login_failed`
- `login_rate_limited`, `verify_rate_limited`
- `login_lockout`, `verify_lockout`
- `license_revoked`, `device_deactivated`

## 3) Revoke license mot user

```bash
curl -X POST "http://localhost:3100/api/v1/admin/revoke-license/USER_ID"
```

Tac dong:
- Thu hoi toan bo session dang hoat dong cua user.
- Plugin se bi chan o lan verify/refresh tiep theo.

## 4) Deactivate mot device cu the

```bash
curl -X POST "http://localhost:3100/api/v1/admin/deactivate-device/DEVICE_ID"
```

Tac dong:
- Khoa 1 may cu the, khong khoa toan bo tai khoan.

## 5) Quy trinh xu ly su co nhanh

1. Phat hien spike tu `metrics`.
2. Xac minh nguon qua `audit_logs` (ip/fingerprint).
3. Deactivate device nghi ngo hoac revoke license neu nghiem trong.
4. Theo doi lai `login_failed_5m` va `5xx_rate` sau can thiep.

## 6) Lich van hanh de xuat

- Hang ngay: check `metrics` 1-2 lan.
- Hang tuan: review `audit_logs` va tong hop bat thuong.
- Hang thang: restore drill + test revoke realtime.
