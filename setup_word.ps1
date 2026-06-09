$excelPath = "C:\Users\uri\Desktop\uri test1.xlsx"
$wordFiles = Get-ChildItem "$env:USERPROFILE\Desktop\*.docx"
$wordPath = $wordFiles[0].FullName

Write-Host "Connecting $wordPath to $excelPath..."

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Open($wordPath)

# חיבור ה-Mail Merge לאקסל
$doc.MailMerge.OpenDataSource($excelPath, 0, $false, $false, $true, $false, "", "", $false, "", "", "Entire Spreadsheet", "SELECT * FROM [Sheet1$]")

# עדכון שדות בטבלה (מיפוי ידני לפי מה שראינו)
# ננסה למצוא את השדות הקיימים ולהחליף אותם
$fields = $doc.MailMerge.Fields
foreach ($field in $fields) {
    $code = $field.Code.Text
    if ($code -like "*firstheb*" -or $code -like "*lastheb*") {
        $field.Code.Text = " MERGEFIELD 'שם נבחן' "
    }
    elseif ($code -like "*formatted_id_num*") {
        $field.Code.Text = " MERGEFIELD 'תעודת זהות' "
    }
    elseif ($code -like "*שם_משתמש*") {
        $field.Code.Text = " MERGEFIELD 'קוד משתמש' "
    }
    elseif ($code -like "*סיסמה*") {
        $field.Code.Text = " MERGEFIELD 'סיסמה' "
    }
}

# טיפול ב-QR - נחפש את המילה "qr" או שדה קיים
$range = $doc.Content
if ($range.Find.Execute("qr")) {
    $range.Text = ""
    $doc.Fields.Add($range, -1, "DISPLAYBARCODE `"{ MERGEFIELD DATA_QR }`" QR \q 3 \s 100", $false)
}

$doc.Save()
Write-Host "Done! Please check your Word document."
