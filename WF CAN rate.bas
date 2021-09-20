Attribute VB_Name = "Module2"
Sub rate()

''Decide Destination

Dim Mypath As String
Mypath = Application.ThisWorkbook.Path

    ''generate date and ACH # from excel file Name


    Dim excelname As String
    Dim tempdate As String
    Dim filedate As String
    Dim ach As String
    
    excelname = ActiveWorkbook.Name
    
    tempdate = Left(excelname, 6)
    filedate = Left(tempdate, 2) & "/" & Mid(tempdate, 3, 2) & "/" & Right(tempdate, 2)

    ach = Mid(excelname, 20, 7)
    


    ''generate rate csv based on Location


    Dim Claim(1 To 3) As String
    Dim ItemClaim(1 To 3) As String


    Dim criteria As String
    Dim i As Integer
    Dim j As Integer

      
    Dim src As Worksheet
    Dim tgt As Worksheet
    Dim copyRange As Range
    Dim lastRow As Long
    Dim header As Variant
    

    Claim(1) = "1.5% Early Payment Discount"
    Claim(2) = "5% Defective Allowance"
    Claim(3) = "2% Advertising Co-Op"

    ItemClaim(1) = "Prompt Payment Discount"
    ItemClaim(2) = "Preset Defective"
    ItemClaim(3) = "Co-op"


    header = Array("External ID", "Credit #", "Customer", "Date", "Posting Period", "Department", _
            "Location", "Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", _
            "To Be Faxed", "Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", _
            "Sale Amnt", "Description", "Taxable", "Apply_Applied", "Apply_payment")


    Set src = ThisWorkbook.Sheets("payment")
    src.AutoFilterMode = False
    lastRow = src.Range("A" & src.Rows.Count).End(xlUp).Row


    '' Location = CG-CAN, ganerate csv based on rate
    For i = 1 To 3
    
        Sheets.Add.Name = "tgt"
        Set tgt = ThisWorkbook.Sheets("tgt")


        tgt.Range("a1:w1") = header

        tgt.Range("a2:a" & lastRow).Value = "CR0001"
        tgt.Range("b2:b" & lastRow).Value = "21"
        tgt.Range("c2:c" & lastRow).Value = "Wayfair.com : Castlegate - CAN Toronto"
        tgt.Range("d2:d" & lastRow).Value = filedate
        tgt.Range("f2:f" & lastRow).Value = "Dot Com"
        tgt.Range("g2:g" & lastRow).Value = "CG-CAN"
        tgt.Range("h2:h" & lastRow).Value = "USD"
        tgt.Range("i2:i" & lastRow).Value = "1"
        tgt.Range("j2:l" & lastRow).Value = "FALSE"
        tgt.Range("m2:m" & lastRow).Value = "Chargeback on CK#" & ach
        tgt.Range("n2:n" & lastRow).Value = Claim(i)
        tgt.Range("o2:o" & lastRow).Value = ItemClaim(i)
        tgt.Range("p2:p" & lastRow).Value = "1"
        tgt.Range("q2:q" & lastRow).Value = "Custom"
        tgt.Range("t2:t" & lastRow).Value = Claim(i)
        tgt.Range("u2:u" & lastRow).Value = "FALSE"
        tgt.Range("v2:v" & lastRow).Value = src.Range("f2:f" & lastRow).Value


        If i = 1 Then
            Set copyRange = src.Range("p2:p" & lastRow)
        ElseIf i = 2 Then
            Set copyRange = src.Range("q2:q" & lastRow)
        Else
            Set copyRange = src.Range("r2:r" & lastRow)
        End If

        copyRange.Copy tgt.Range("r2")
        copyRange.Copy tgt.Range("s2")
        copyRange.Copy tgt.Range("w2")

        tgt.Name = tempdate & "_WF " & Left(Claim(i), InStr(Claim(i), "%"))
     
    
        tgt.Copy
        ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name & ".csv", _
        FileFormat:=xlCSV, CreateBackup:=False, Local:=True
        ActiveWorkbook.Close
            
        Application.DisplayAlerts = False
        tgt.Delete
        Application.DisplayAlerts = True

    
    Next i


    
End Sub

