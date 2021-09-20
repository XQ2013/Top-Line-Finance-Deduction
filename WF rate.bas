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
    Dim csvClaim(1 To 3) As String
    Dim csvLocation(1 To 2) As String


    Dim criteria As String
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer

      
    Dim src As Worksheet
    Dim tgt As Worksheet
    Dim filterRange As Range
    Dim copyRange As Range
    Dim fillRange As Range
    Dim lastRowsrc As Long
    Dim lastRowtgt As Long
    

    Claim(1) = "1.5% Early Payment Discount"
    Claim(2) = "4% Defective Allowance"
    Claim(3) = "2% Advertising Co-Op"

    ItemClaim(1) = "Prompt Payment Discount"
    ItemClaim(2) = "Preset Defective"
    ItemClaim(3) = "Co-op"

    csvClaim(1) = "1.5 discount"
    csvClaim(2) = "4 defective"
    csvClaim(3) = "2 co-op"

    csvLocation(1) = "CA&IL"
    csvLocation(2) = "CG-ER"


    '' get location info

    Set src = ThisWorkbook.Sheets("payment")
    src.AutoFilterMode = False
    lastRowsrc = src.Range("A" & src.Rows.Count).End(xlUp).Row




    ''set filter


    ''generate based on Location
    For i = 1 To 2
    
    src.AutoFilterMode = False
    Set filterRange = src.Range("A1:Q" & lastRowsrc)
    
    Dim rangeN1 As Range
    Dim rangeN2 As Range
    
    If i = 1 Then
    
        filterRange.AutoFilter Field:=13, Criteria1:="<>CG-ER"
        n = Application.WorksheetFunction.CountIf( _
        Sheets("payment").Range("m2:m" & lastRowsrc), "<>CG-ER")
        
        Else:
        filterRange.AutoFilter Field:=13, Criteria1:="=CG-ER"
        n = Application.WorksheetFunction.CountIf( _
        Sheets("payment").Range("m2:m" & lastRowsrc), "=CG-ER")
        
    End If
    

    If n > 0 Then


        src.AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("M:M"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortNormal
        With src.AutoFilter.Sort
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        src.AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("L:L"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortNormal
        With src.AutoFilter.Sort
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
            ''generate based on claim
            For j = 1 To 3
            
            ''sort based on Location and customer


            Sheets.Add.Name = "tgt"
            Set tgt = ThisWorkbook.Sheets("tgt")

            Dim header As Variant
            header = Array("External ID", "Credit #", "Customer", "Date", _
                "Posting Period", "Department", "Location", "Currency", "Exchange Rate", _
                "To Be Printed", "To Be E-mailed", "To Be Faxed", "Memo", "PO #", _
                "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", _
                "Description", "Taxable", "Apply_Applied", "Apply_payment")
                tgt.Range("a1:w1") = header


            ' start in row 2 to prevent copying the header
            ' copy the visible cells to our target range
            Set copyRange = src.Range("L2:L" & lastRowsrc)
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("C2")
            Set copyRange = src.Range("m2:M" & lastRowsrc)
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("g2")
            Set copyRange = src.Range("k2:k" & lastRowsrc)
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("v2")

            If j = 1 Then
            Set copyRange = src.Range("n2:n" & lastRowsrc)
            ElseIf j = 2 Then
            Set copyRange = src.Range("o2:o" & lastRowsrc)
            Else
            Set copyRange = src.Range("p2:p" & lastRowsrc)
            End If

            
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("r2")
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("s2")
            copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("w2")




            ''fill range
            lastRowtgt = tgt.Range("C" & tgt.Rows.Count).End(xlUp).Row

            tgt.Range("D2:D" & lastRowtgt).Value = filedate
            tgt.Range("f2:f" & lastRowtgt).Value = "Dot Com"
            tgt.Range("H2:H" & lastRowtgt).Value = "USD"
            tgt.Range("I2:I" & lastRowtgt).Value = "1"
            tgt.Range("J2:L" & lastRowtgt).Value = "FALSE"
            tgt.Range("M2:M" & lastRowtgt).Value = "Chargeback on CK#" & ach
            tgt.Range("N2:N" & lastRowtgt).Value = src.Cells(1, 13 + j).Value

            tgt.Range("O2:O" & lastRowtgt).Value = ItemClaim(j)
            tgt.Range("P2:P" & lastRowtgt).Value = "1"
            tgt.Range("Q2:Q" & lastRowtgt).Value = "Custom"
            tgt.Range("t2:t" & lastRowtgt).Value = src.Cells(1, 13 + j).Value
            tgt.Range("U2:U" & lastRowtgt).Value = "FALSE"

            'External ID
            tgt.Cells(2, 2).Value = 21
            tgt.Range("a2").Value = "CR0001"
            Dim m As Integer
            m = 3
            
            If lastRowtgt = 2 Then
                Else:
                If lastRowtgt = 3 Then

                    If (tgt.Cells(m, 3).Value = tgt.Cells(m - 1, 3).Value) And _
                    (tgt.Cells(m, 7).Value = tgt.Cells(m - 1, 7).Value) Then
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2)
                        Else:
                        tgt.Cells(m, 2).Value = tgt.Cells(m - 1, 2).Value + 1
                    End If
                    Sheets("tgt").Select
                    Range("A3").Select
                    ActiveCell.FormulaR1C1 = _
                    "=""CR00""&TEXT(RC[1]-20,""00"")"
                    
                    Else:
                
                    Do
                    If (tgt.Cells(m, 3).Value = tgt.Cells(m - 1, 3).Value) And _
                    (tgt.Cells(m, 7).Value = tgt.Cells(m - 1, 7).Value) Then
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2)
                        m = m + 1
                        Else:
                        tgt.Cells(m, 2).Value = tgt.Cells(m - 1, 2).Value + 1
                        m = m + 1
                    End If
                    Loop Until m = lastRowtgt + 1

                    Sheets("tgt").Select
                    Range("A2").Select
                    ActiveCell.FormulaR1C1 = _
                    "=""CR00""&TEXT(RC[1]-20,""00"")"
                    Selection.AutoFill Destination:=Range("A2:A" & lastRowtgt)
                End If
  

            End If


        ''export single sheet csv
        
        tgt.Name = tempdate & "_WF " & csvClaim(j) & "(" & csvLocation(i) & ")"
     
    
        tgt.Copy
        ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name & ".csv", _
        FileFormat:=xlCSV, CreateBackup:=False, Local:=True
        ActiveWorkbook.Close
            
        Application.DisplayAlerts = False
        tgt.Delete
        Application.DisplayAlerts = True

        Next j
        src.AutoFilterMode = False
    
    Else:
    src.AutoFilterMode = False
    
    End If

    Next i


End Sub















