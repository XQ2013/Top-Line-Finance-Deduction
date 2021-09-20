Attribute VB_Name = "Module3"

    
Sub paymemnt()

    ''Decide Destination

    Dim Mypath As String
    Mypath = Application.ThisWorkbook.Path

    ''generate date , ACH #, total amount from excel file Name


    Dim excelname As String
    Dim tempdate As String
    Dim filedate As String
    Dim ach As String
    Dim total As Variant

    excelname = ActiveWorkbook.Name

    tempdate = Left(excelname, 6)
    filedate = Left(tempdate, 2) & "/" & Mid(tempdate, 3, 2) & "/" & Right(tempdate, 2)

    ach = Mid(excelname, 20, 7)

    total = Mid(excelname, InStr(excelname, "$") + 1)
    total = Format(Replace(total, ".xlsm", ""), "0.00")



    ''create open column



    Dim src As Worksheet
    Dim tgt As Worksheet
    Dim lastRowsrc As Long
    Dim lastRowtgt As Long
    Dim invoiceTotal As Variant
    Dim amount As Variant


    Set src = ThisWorkbook.Sheets("payment")
    src.AutoFilterMode = False
    src.Range("c:c").ClearContents
    lastRowsrc = src.Range("A" & src.Rows.Count).End(xlUp).Row
    
    invoiceTotal = Round(Application.WorksheetFunction.Sum( _
    src.Range("D2:D" & lastRowsrc)), 2)

    ''generate payment csv
    Sheets.Add.Name = "tgt"
    Set tgt = ThisWorkbook.Sheets("tgt")
    
    
    If (invoiceTotal - total < 0) Then
        lastRowtgt = lastRowsrc
        tgt.Range("k2:k" & lastRowtgt).Value = src.Range("d2:d" & lastRowtgt).Value
     
     Else:
        lastRowtgt = 3
        amount = 0
            

        Do
            lastRowtgt = lastRowtgt + 1
            amount = Round(Application.WorksheetFunction.Sum(src.Range("D2:D" & lastRowtgt)), 2)
        Loop Until (amount - total) >= 0
         
        Range(src.Cells(lastRowtgt, 3), src.Cells(lastRowsrc, 3)).Value = "open"
        tgt.Range("k2:k" & lastRowtgt).Value = src.Range("d2:d" & lastRowtgt).Value
        tgt.Cells(lastRowtgt, 11).Value = Format( _
        total - Application.WorksheetFunction.Sum(tgt.Range("k2:k" & (lastRowtgt - 1))), "0.00")
        
    End If





    Dim header As Variant
    header = Array("External ID", "Customer", "Department", "Location", "Total Payment amount", _
        "posting period", "Date", "ACH Payment#", "Memo", "Payment Method", "Payment applied", _
        "Account", "A/R Account Required Field*", "Internal ID", "Discount", "undep.fund")
    tgt.Range("a1:p1") = header

    tgt.Range("a2:a" & lastRowtgt).Value = ach
    tgt.Range("b2:b" & lastRowtgt).Value = "Wayfair.com"
    tgt.Range("c2:c" & lastRowtgt).Value = "Dot com"
    tgt.Range("d2:d" & lastRowtgt).Value = "IL-S"
    tgt.Range("e2:e" & lastRowtgt).Value = total
    tgt.Range("g2:g" & lastRowtgt).Value = filedate
    tgt.Range("h2:h" & lastRowtgt).Value = ach
    tgt.Range("j2:j" & lastRowtgt).Value = "Wire"

    tgt.Range("l2:l" & lastRowtgt).Value = _
    "10021 Bank of America : Bank of America (Depository) "
    tgt.Range("m2:m" & lastRowtgt).Value = "12040 Accounts Receivable"
    tgt.Range("n2:n" & lastRowtgt).Value = src.Range("k2:k" & lastRowtgt).Value

    tgt.Range("o2:o" & lastRowtgt).Value = 0
    tgt.Range("p2:p" & lastRowtgt).Value = "FALSE"
    



    ''export single sheet csv
    
    tgt.Name = tempdate & "_Wayfair Payment"
 

    tgt.Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name & ".csv", _
    FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    ActiveWorkbook.Close
        
    Application.DisplayAlerts = False
    tgt.Delete
    Application.DisplayAlerts = True


End Sub
















