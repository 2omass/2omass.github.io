# HR File Organizer (Direct on Computer + Always-On for Windows ARM)

برنامج احترافي لتنظيم ملفات الموظفين **مباشرة على الكمبيوتر** مع تحليل ذكي لمحتوى الملف:
- قراءة الاسم + المحتوى (بما فيها Excel الحديثة `.xlsx/.xlsm` و Word `.docx`).
- تحديد **القسم** (مثل: `GovernmentRelations`, `Finance`, `HumanResources`).
- تحديد **نوع الملف** داخل القسم (مثل: `Payroll`, `Contracts`, `GovernmentDocuments`).
- استخراج اسم الموظف والتاريخ (إن وُجدا).
- إعادة تسمية احترافية موحدة.
- نقل تلقائي داخل هيكل مجلدات احترافي.
- دعم وضع التشغيل المستمر `--watch` للعمل طوال الوقت.

## هيكل الحفظ النهائي

```text
organized/
  <Division>/
    <Category>/
      <Employee>/
        <DEPARTMENT_DIVISION_CATEGORY_EMPLOYEE_DATE_SEQ.ext>
```

## التشغيل لمرة واحدة (ينفذ النقل فعليًا)

```bash
python3 hr_file_organizer.py /path/to/files --department "HR"
```

## معاينة فقط بدون نقل

```bash
python3 hr_file_organizer.py /path/to/files --department "HR" --dry-run
```

## تشغيل مستمر (Always-On)

```bash
python3 hr_file_organizer.py /path/to/files --department "HR" --watch --interval 30
```

- السكربت سيفحص كل 30 ثانية وينقل الملفات الجديدة تلقائيًا.
- لإيقافه: `Ctrl + C`.

## تشغيل دائم تلقائي على Windows ARM (بعد كل إعادة تشغيل)

> المتطلب: Python ARM64 مثبت على الجهاز، وPowerShell بصلاحية مناسبة.

1) افتح PowerShell داخل مجلد المشروع.
2) نفّذ:

```powershell
.\install_windows_startup_task.ps1 -SourceDir "C:\HR\Inbox" -Department "CORP" -PythonExe "python" -IntervalSeconds 30 -TaskName "HRFileOrganizerARM"
```

3) سيُسجَّل Scheduled Task يعمل عند إقلاع النظام ويشغل المنظم بشكل مستمر.

## تحديد مخرج مخصص

```bash
python3 hr_file_organizer.py /path/to/files --output /path/to/output --department "HR"
```

## ملاحظات
- يتم إنشاء تقرير CSV افتراضيًا في:
  - `<source>/organized/hr_organizer_report.csv`
- دعم Excel يعتمد على قراءة XML الداخلي لملفات OOXML (`.xlsx/.xlsm/.xltx/.xltm`) للتصنيف الذكي.
- ملفات `.xls` القديمة (ثنائية) تُصنّف غالبًا من اسم الملف/النص المتاح بدون مكتبات إضافية.
