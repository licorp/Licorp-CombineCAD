# Pre-release 10 phut checklist (Add-in + License)

## A. Build & package

- [ ] Chay `build-package.bat`
- [ ] Xac nhan output:
  - [ ] `artifacts/Licorp_CombineCAD_Setup_1.0.0.zip`
  - [ ] `artifacts/release/1.0.0/installer`

## B. Deploy local de test nhanh

- [ ] Chay `build-deploy.bat`
- [ ] Mo Revit, kiem tra tab `Licorp` xuat hien
- [ ] Mo dialog `Combine CAD` khong loi

## C. Sanity test license flow

- [ ] Login online thanh cong (`Licensed`)
- [ ] Tat mang tam thoi, xac nhan vao `Grace mode`
- [ ] Qua moc grace test, xac nhan `Expired`
- [ ] Bat mang lai, login/refresh lai thanh cong

## D. Revoke test

- [ ] Goi admin revoke license hoac deactivate device
- [ ] Tren plugin, lan verify/refresh tiep theo hien `Revoked`
- [ ] Co audit log event tuong ung

## E. Evidence & docs

- [ ] Dien evidence vao:
  - [ ] `docs/evidence/SR-01/notes.md`
  - [ ] `docs/evidence/SR-02/notes.md`
  - [ ] `docs/evidence/SR-03/notes.md`
  - [ ] `docs/evidence/SR-04/notes.md`
- [ ] Cap nhat `docs/store_readiness_test_matrix.md` -> PASS cho case dat
- [ ] Cap nhat `docs/backup_restore_runbook.md` voi backup artifact + ket qua restore drill full

## F. Go/No-Go

- [ ] Tat ca SR-01..SR-04 = PASS
- [ ] Khong con loi nghiem trong trong smoke test
- [ ] San sang dong goi/nop Autodesk App Store
