# API Versioning Policy

## Nguyen tac

- Toan bo endpoint phien ban hien tai su dung prefix `/api/v1`.
- Trong v1, uu tien backward compatibility cho thay doi nho.

## Deprecation policy

1. Cong bo truoc endpoint sap deprecate.
2. Chay song song endpoint cu/moi trong giai doan chuyen doi.
3. Den ngay sunset, endpoint cu tra `410` kem huong dan migration.

## Breaking changes

- Breaking change phai dua vao version moi (vi du `/api/v2`).
- Cung cap changelog va migration guide.
