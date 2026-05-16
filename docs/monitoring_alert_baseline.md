# Monitoring va Alert Baseline

## Muc tieu

- Theo doi suc khoe API licensing theo thoi gian thuc.
- Canh bao som khi co dau hieu tan cong/bruteforce hoac backend suy giam.

## Chi so toi thieu

- `5xx_rate`: ty le response 5xx tren tong request trong cua so 5 phut.
- `avg_latency_ms`: do tre trung binh trong cua so 5 phut.
- `p95_latency_ms`: do tre p95 trong cua so 5 phut.
- `login_failed_5m`: so lan login fail trong 5 phut.
- `login_fail_spike`: co/khong khi `login_failed_5m` vuot nguong.

## Nguong canh bao de xuat

- High 5xx: `5xx_rate >= 0.05` lien tuc 5 phut.
- Latency cao: `avg_latency_ms >= 1000` hoac `p95_latency_ms >= 2000` lien tuc 5 phut.
- Login fail spike: `login_failed_5m >= 20` trong cua so 5 phut.

## Kenh canh bao

- Email: gui toi nhom van hanh va chu san pham.
- Slack: gui vao kenh `#licensing-alerts` (neu da co).

## Playbook xu ly nhanh

1. Xac minh alert bang endpoint admin metrics.
2. Kiem tra audit log lien quan (`login_failed`, `*_rate_limited`, `*_lockout`).
3. Neu la dot tan cong: nang lockout window tam thoi, theo doi IP/fingerprint bat thuong.
4. Neu la loi he thong: khoanh vung thay doi gan nhat, rollback neu can.

## Endpoint phuc vu quan sat

- `GET /api/v1/admin/metrics`
- `GET /api/v1/admin/audit-logs`
