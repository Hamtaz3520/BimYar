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
- **Phase 3 Step 3.6 (`BMY-CX-20260813-010`):** Completed — برنامه تثبیت پایگاه داده در `Documents/Progress/DATABASE_STABILIZATION_PLAN.md` ثبت شد: ۲۸ بسته تغییر پیشنهادی (۱۴ مورد MUST پیش از Phase 4، ۱۰ مورد SHOULD پیش از Production و ۴ مورد CAN DEFER) با وضعیت فعلی، نوع دقیق رابطه/ایندکس، ریسک داده، پیش‌نیاز، ترتیب اجرا، کنترل پشتیبان/بازگشت و موانع سازگاری Access 2016.
- **Phase 3 Step 3.7-R1 (`BMY-CX-20260813-011-R1`):** Completed — برنامه اجرایی `Documents/Progress/PHASE3_MUST_FIX_EXECUTION_PLAN.md` با برنامه تثبیت ادغام‌شده تطبیق داده شد. هر ۱۴ مورد MUST (`DB-01` تا `DB-14`) و هر ۷ بسته اجرایی همچنان معتبرند. بسته ۱، «حفاظت و پیش‌بررسی فقط‌خواندنی» (`DB-01/02`)، کوچک‌ترین نخستین بسته امن باقی ماند و هیچ پاک‌سازی داده یا تغییر ساختاری را مجاز نمی‌کند.
- **Phase 3 Step 3.8 (`BMY-CX-20260813-012`):** بسته اجرایی Batch 1 و checklist ماشین‌خوان در `Documents/Progress/BATCH1_PROTECTION_BASELINE.md` و `Documents/Progress/BATCH1_EXECUTION_CHECKLIST.md` ایجاد و preflight مخزن اجرا شد. Gate در وضعیت **BLOCKED — NOT PASSED** است، زیرا فایل Access ناشناس‌شده/خروجی زنده، مجوز مالک و حریم خصوصی، محل خارجی backup/evidence و محیط Access 2016/ACE ارائه نشده‌اند. هیچ فایل پایگاه داده باز نشد، هیچ DDL/DML اجرا نشد و schema، داده یا منطق SQL/VBA تغییر نکرد. Batch 2 و هر تغییر ساختاری تا اجرای موفق A تا L، تطبیق hash/count و امضای مالک و بازبین ممنوع است.
- پیش‌نیازهای خطای 3163، رابطه‌ها، ایندکس‌ها، ImportSession، پرداخت/بانک و سازگاری Access 2016 در برنامه اجرایی تفکیک و با ترتیب مرجع همسو شدند. تصمیم‌های کسب‌وکار درباره session، اختیاری‌بودن/مالکیت رابطه‌ها، اتصال پرداخت، مدل حساب/بچ بانکی، دامنه یکتایی و قواعد تطبیق بازاریاب همچنان مانع اجرا هستند.
- **Exact next implementation task:** رفع موانع ثبت‌شده Step 3.8 و ادامه همان Batch 1 روی workstation دارای Access 2016 با نسخه تولیدی ناشناس‌شده و موردتأیید مالک: اجرای checklist بخش‌های A تا L، تطبیق hash و row count، ثبت evidence خارجی و اخذ امضای مالک/بازبین. پیش از `PASS` شدن gate هیچ DDL یا Batch 2 آغاز نشود.
- یافته کلیدی: ۱۷ کلید اصلی و تنها ۵ ایندکس ثانویه وجود دارد؛ هیچ `FOREIGN KEY` و هیچ `DEFAULT` تعریف نشده است. رابطه‌های منطقی، ایندکس‌های FK/جستجو/تطبیق، ناسازگاری اندازه و نوع فیلدها و ابهام شماره‌گذاری 012 مستند شده‌اند.
- فایل‌های محلی تاریخی 018 تا 023 خالی بوده‌اند و در Repository فعلی موجود نیستند؛ آن‌ها فقط به‌عنوان اقلام مفقود/پیاده‌نشده ثبت شدند و در این مرحله هیچ نام، شیء یا SQL برای آن‌ها حدس زده یا ساخته نشد.
- وضعیت کلی Phase 3 همچنان **In Progress** و برای بسته‌شدن آماده نیست؛ تطبیق و پروفایل فقط‌خواندنی Access واقعی، تأیید جدول‌های session/error/log و ترتیب اجرا، اجرای آزمایشی در Access 2016، تکمیل SQL_v5، دریافت VBAهای واقعی و نمونه‌های Excel ناشناس‌شده، تصویب قواعد کسب‌وکار، baseline نهایی و بستن رسمی Phase 3 باقی مانده‌اند. Phase 4 تا تکمیل پیش‌نیازها و آزمون‌های همه موارد MUST مسدود است.
