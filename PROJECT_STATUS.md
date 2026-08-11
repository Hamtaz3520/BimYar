# وضعیت فعلی پروژه BimYar

## فاز فعلی

- **Current Phase:** 3
- **Current Phase Title:** Repository Stabilization and Source Intake
- **Next Phase:** 4 - Import Engine Stabilization

## معرفی پروژه

- **نام پروژه:** BimYar
- **هدف:** سامانه مدیریت و حسابداری دفتر بیمه با تمرکز اولیه بر Microsoft Access 2016، VBA، SQL و واردکردن گزارش‌های Excel بیمه‌نامه‌ها.
- **فایل اصلی فعلی Access:** `BimYarCRM_v0.1`

## ساختار نسخه‌های SQL

- پوشه `Database/SQL/SQL_v4` مجموعه کاری قبلی و مرجع اسکریپت‌های موجود است.
- پوشه `Database/SQL/SQL_v5` فقط برای فایل‌های SQL جدید یا فایل‌هایی است که نسبت به نسخه قبلی تغییر کرده‌اند.
- فایل‌های بدون تغییر نباید بی‌دلیل در `SQL_v5` تکرار شوند.

## جدول‌های پایگاه داده

### جدول‌های تأییدشده ساخته‌شده

- `tblRoles`
- `tblUsers`
- `tblMarketers`
- `tblInsuranceCompanies`
- `tblBranches`
- `tblInsuranceTypes`
- `tblBanks`
- `tblAccountHeads`
- `tblSettings`
- `tblCustomers`
- `tblPolicies`
- `tblInstallments`
- `tblPayments`
- `tblBankTransactions`
- `tblFollowUps`
- `tblRenewals`
- `tblRenewalStatuses`
- `tblImportPolicies`
- `tblImportPolicies_Work`
- `tblImportSessions`
- `tblImportErrors`
- `tblImportLogs`

### جدول‌های برنامه‌ریزی‌شده

- `tblImportPayments`
- `tblImportBank`
- `tblUserLogs`
- `tblPermissions`
- `tblRolePermissions`

## ماژول‌های VBA فعلی یا ایجادشده

- `modGlobals`
- `modFunctions`
- `modCreateIndexes`
- `modSeedData`
- `modImportEngine`
- `modImportExcelSizeCheck`
- `modImportDiagnostic`

## وضعیت Import

- دستور اصلی `ImportPolicyReport` ساخته شده است.
- گزارش صدور کلی و گزارش بازاریاب باید از Excel وارد شوند.
- فایل کلی ابتدا وارد می‌شود و سپس فایل‌های بازاریاب.
- ردیف‌های تکراری باید تطبیق و حذف شوند.
- ردیف‌های اقساط باید به ردیف اصلی بیمه‌نامه متصل شوند.
- داده‌های سربرگ باید برای ردیف‌های قسط تکمیل شوند.
- سربرگ‌های تکرارشده صفحات Excel باید نادیده گرفته شوند.
- ردیف‌های باقی‌مانده بدون بازاریاب باید امکان تخصیص دستی داشته باشند.

## خطای فعلی Import

```text
Run-time error 3163
The field is too small to accept the amount of data you attempted to add.
```

## اقدامات تشخیصی انجام‌شده

- فیلد `InsuredName` به Long Text تغییر داده شده است.
- فیلد `NationalCode` به دلیل طول بیشتر داده Excel بزرگ‌تر شده است.
- Mapping اولیه بر اساس عنوان فارسی ستون‌ها کنار گذاشته شد.
- Mapping بر اساس شماره ستون با `GetImportFieldNameByColumn` انجام شده است.
- تست‌های اندازه فیلد Excel و جداول Access ایجاد شده‌اند.

## فایل‌های شناخته‌شده فعلی SQL_v5

- `024_Create_tblRenewalStatuses.sql`
- `025_Insert_modSeedData.bas`
- `026_Create_tblImportPolicies_Work.sql`

## محدودیت این مرحله

در این مرحله نباید هیچ SQL یا VBA جدیدی از روی حدس تولید شود. فقط ساختار پوشه‌ها و مستندات وضعیت پروژه ایجاد و نگهداری می‌شوند.

## آمادگی Source Intake

- مرحله Source Intake برای دریافت و ثبت اولین بسته واقعی فایل‌های پروژه آماده شده است.
