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

- پوشه `Database/SQL/SQL_v4` شامل مجموعه واقعی فعلی است: ۱۸ فایل SQL غیر-placeholder که ۱۷ جدول و ۵ ایندکس `tblInstallments` را تعریف می‌کنند؛ `.gitkeep` در شمارش منبع واقعی لحاظ نمی‌شود.
- پوشه `Database/SQL/SQL_v5` فقط برای فایل‌های SQL جدید یا فایل‌هایی است که نسبت به نسخه قبلی تغییر کرده‌اند.
- فایل‌های بدون تغییر نباید بی‌دلیل در `SQL_v5` تکرار شوند.

## جدول‌های پایگاه داده

### جدول‌های دارای DDL تأییدشده در Repository

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

### جدول‌های گزارش‌شده ساخته‌شده، اما بدون DDL موجود

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

## فایل‌های واقعی تأییدشده در Source Intake

- `Database/SQL/SQL_v5/024_Create_tblRenewalStatuses.sql`
- `Database/VBA/Modules/025_Insert_modSeedData.bas`
- `Database/SQL/SQL_v5/026_Create_tblImportPolicies_Work.sql`

مسیر تکراری و نادرست `Database/VBA/025_Insert_modSeedData.bas` وجود ندارد.

## محدودیت این مرحله

در این مرحله نباید هیچ SQL یا VBA جدیدی از روی حدس تولید شود. فقط ساختار پوشه‌ها و مستندات وضعیت پروژه ایجاد و نگهداری می‌شوند.

## وضعیت Source Intake

- **Phase 3 Step 3.2:** Completed
- اولین بسته واقعی Source Intake با شناسه `BMY-CX-20260812-006-R3` از نظر وجود، مسیر، نام، پسوند، خوانایی و encoding تأیید شد.
- **Phase 3 Step 3.3:** Completed — تحلیل رسمی شکاف پیاده‌سازی در `Documents/Progress/BIMYAR_GAP_ANALYSIS.md` ثبت شد.
- **Phase 3 Step 3.4:** Completed — موجودی snapshot پیشین در زمان خالی‌بودن پوشه ثبت شد؛ موجودی Step 3.5 اکنون آن شمارش قبلی را به‌روزرسانی می‌کند.
- **Phase 3 Step 3.5 (`BMY-CX-20260813-009`):** Completed — هر ۱۸ فایل واقعی SQL_v4 فعلی (۱۷ فایل ساخت جدول و یک فایل مستقل شامل ۵ ایندکس) به‌طور کامل در `Documents/Progress/SQL_V4_SCHEMA_ANALYSIS.md` تحلیل و موجودی `Documents/Progress/SQL_V4_INVENTORY.md` به‌روزرسانی شد.
- یافته کلیدی: ۱۷ کلید اصلی و تنها ۵ ایندکس ثانویه وجود دارد؛ هیچ `FOREIGN KEY` و هیچ `DEFAULT` تعریف نشده است. رابطه‌های منطقی، ایندکس‌های FK/جستجو/تطبیق، ناسازگاری اندازه و نوع فیلدها و ابهام شماره‌گذاری 012 مستند شده‌اند.
- فایل‌های محلی تاریخی 018 تا 023 خالی بوده‌اند و در Repository فعلی موجود نیستند؛ آن‌ها فقط به‌عنوان اقلام مفقود/پیاده‌نشده ثبت شدند و در این مرحله هیچ نام، شیء یا SQL برای آن‌ها حدس زده یا ساخته نشد.
- وضعیت کلی Phase 3 همچنان **In Progress** است؛ تطبیق با Access واقعی، ترتیب اجرا، رابطه‌ها/ایندکس‌های واقعی، تکمیل SQL_v5، سایر VBAها، نمونه‌های Excel و اسناد معماری/قواعد کسب‌وکار و بستن رسمی Phase 3 باقی مانده‌اند.
