Option Compare Database
Option Explicit

Public Sub Seed_RenewalStatuses()

    Dim db As DAO.Database
    Set db = CurrentDb

    'پاک کردن اطلاعات قبلي در صورت وجود
    db.Execute "DELETE FROM tblRenewalStatuses", dbFailOnError

    'ثبت وضعيت‌هاي تمديد
    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Pending','در انتظار پيگيري',True,1,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Contacted','تماس گرفته شده',True,2,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Interested','تمايل به تمديد',True,3,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Renewed','تمديد شده',True,4,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Declined','عدم تمديد',True,5,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('NoAnswer','پاسخگو نيست',True,6,Now())"

    db.Execute "INSERT INTO tblRenewalStatuses (StatusCode, StatusTitle, IsActive, SortOrder, CreatedDate) " & _
               "VALUES ('Expired','منقضي شده',True,7,Now())"

    MsgBox "وضعيت‌هاي تمديد با موفقيت ايجاد شد", vbInformation

    Set db = Nothing

End Sub

