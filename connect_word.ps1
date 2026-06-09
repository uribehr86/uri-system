    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Open('C:\Users\uri\Desktop\טופס התחברות לבחינה.docx')
    $doc.MailMerge.OpenDataSource('C:\Users\uri\Desktop\uri test1.xlsx', 0, $false, $false, $true, $false, '', '', $false, '', '', 'Entire Spreadsheet', 'SELECT * FROM [Sheet1$]')
    
    $range = $doc.Content
    if ($range.Find.Execute('qr')) {
        $range.Text = ''
        $field = $doc.Fields.Add($range, -1, 'DISPLAYBARCODE "{ MERGEFIELD DATA_QR }" QR \q 3 \s 120', $false)
    }
    
    $doc.Save()
    Write-Host 'DONE_SUCCESSFULLY'
