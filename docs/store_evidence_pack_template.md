# Store Evidence Pack Template

## Cau truc thu muc de xuat

```text
docs/
  evidence/
    SR-01/
      screenshot_01.png
      api_log.txt
      audit_log.json
      notes.md
    SR-02/
      screenshot_01.png
      api_log.txt
      audit_log.json
      notes.md
    SR-03/
      screenshot_01.png
      api_log.txt
      audit_log.json
      notes.md
    SR-04/
      screenshot_01.png
      api_log.txt
      audit_log.json
      notes.md
```

## Quy uoc dat ten file

- Screenshot: `screenshot_<nn>.png`
- API log: `api_log.txt`
- Audit dump: `audit_log.json`
- Mo ta ket qua: `notes.md`

## Mau `notes.md` cho moi SR

```md
# SR-XX Evidence

- Thoi gian test: YYYY-MM-DD HH:mm UTC
- Nguoi test: ...
- Moi truong: ...
- Ket qua: PASS | FAIL | PARTIAL

## Buoc thuc hien
1. ...
2. ...

## Bang chung dinh kem
- screenshot_01.png: ...
- api_log.txt: ...
- audit_log.json: ...

## Ghi chu
- ...
```

## Dieu kien chot PASS tung case

- SR-01: Co log/audit xac nhan gioi han so thiet bi hoac rebind dung ky vong.
- SR-02: Co screenshot UI `Grace mode` va `Expired` sau moc grace.
- SR-03: Co bang chung revoke va plugin chuyen `Revoked` o lan verify/refresh tiep.
- SR-04: Co bang chung refresh replay bi tu choi va khong cap token moi.
