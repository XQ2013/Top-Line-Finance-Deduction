Attribute VB_Name = "Module1"

Sub taxInvoice()

    ''Decide Destination

    Dim Mypath As String
    Mypath = Application.ThisWorkbook.Path

    ''generate date , ACH #, total amount from excel file Name


    Dim excelname As String
    Dim tempdate As String
    Dim filedate As String


    excelname = ActiveWorkbook.Name

    tempdate = Left(excelname, 6)
    filedate = Left(tempdate, 2) & "/" & Mid(tempdate, 3, 2) & "/" & Right(tempdate, 2)


    ''generate tax csv

    Dim src As Worksheet
    Dim tgt As Worksheet
    Dim lastRow As Long


    Set src = ThisWorkbook.Sheets("payment")
    src.AutoFilterMode = False
    Sheets.Add.Name = "tgt"
    Set tgt = ThisWorkbook.Sheets("tgt")
    
    lastRow = src.Range("A" & src.Rows.Count).End(xlUp).Row
    
    
    Dim header As Variant
    header = Array("ExternalID", "Invoice Date", "Customer", "Department", "Location", "PO#", _
        "Memo", "Due Date", "Commission Rate", "Download To A1Warehouse", "Item", "Description", _
        "Price level", "Sell Price", "Amount", "NS Item Type", "Ship Via ")
    tgt.Range("a1:q1") = header
    
    

    tgt.Range("a2:a" & lastRow).Value = src.Range("h2:h" & lastRow).Value
    tgt.Range("b2:b" & lastRow).Value = src.Range("a2:a" & lastRow).Value
    tgt.Range("c2:c" & lastRow).Value = "Wayfair.com : Castlegate - CAN Toronto"
    tgt.Range("d2:d" & lastRow).Value = "Dot Com"
    tgt.Range("e2:e" & lastRow).Value = "CG-CAN"
    tgt.Range("f2:f" & lastRow).Value = src.Range("h2:h" & lastRow).Value
    tgt.Range("g2:g" & lastRow).Value = "13% HST (Harmonized Sales Tax) for CG-CAN only"
    tgt.Range("h2:h" & lastRow).Value = src.Range("d2:d" & lastRow).Value
    tgt.Range("i2:i" & lastRow).Value = "0.00%"
    tgt.Range("j2:j" & lastRow).Value = "FALSE"
    tgt.Range("k2:k" & lastRow).Value = "13% HST"
    tgt.Range("l2:l" & lastRow).Value = "13% HST (Harmonized Sales Tax) for CG-CAN only"
    tgt.Range("m2:m" & lastRow).Value = "custom"
    tgt.Range("n2:n" & lastRow).Value = src.Range("j2:j" & lastRow).Value
    tgt.Range("o2:o" & lastRow).Value = src.Range("j2:j" & lastRow).Value
    tgt.Range("p2:p" & lastRow).Value = "Discount"
    tgt.Range("q2:q" & lastRow).Value = "Pick Up"

    ''export
    tgt.Name = tempdate & "_WF Canada_Tax Invoice"
 

    tgt.Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name & ".csv", _
    FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    ActiveWorkbook.Close
        
    Application.DisplayAlerts = False
    tgt.Delete
    Application.DisplayAlerts = True


End Sub
   	



















