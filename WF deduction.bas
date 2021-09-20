Attribute VB_Name = "Module1"


Sub deductions()


Dim Mypath As String
Mypath = Application.ThisWorkbook.Path


''generate date and ACH # from excel file name

Dim excelname As String
Dim tempdate As String
Dim filedate As String
Dim ach As String

excelname = ActiveWorkbook.Name

tempdate = Left(excelname, 6)
filedate = Left(tempdate, 2) & "/" & Mid(tempdate, 3, 2) & "/" & Right(tempdate, 2)

ach = Mid(excelname, 20, 7)


''generate location Array
Dim src As Worksheet
Set src = ThisWorkbook.Sheets("deduction")

src.AutoFilterMode = False
Dim lastRow As Long
lastRow = src.Range("A" & src.Rows.Count).End(xlUp).Row

src.Range("l:l").ClearContents

src.Range("j1:j" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=src.Range("l1"), Unique:=True

Dim n As Integer
n = src.Range("l" & src.Rows.Count).End(xlUp).Row - 1


ReDim Location(1 To n) As String

Dim filterRange As Range
Dim copyRange As Range
Dim fillRange As Range
Dim lastRowtgt As Long



For i = 1 To n

    Location(i) = src.Cells(i + 1, 12).Value

    criteria = Location(i)

    Dim tgt As Worksheet
    Sheets.Add.Name = "tgt"
    Set tgt = ThisWorkbook.Sheets("tgt")
    
    src.AutoFilterMode = False
    Set filterRange = src.Range("A1:K" & lastRow)
    ' filter range based on column J (Location)

    filterRange.AutoFilter Field:=10, Criteria1:=Location(i)

    
    ActiveWorkbook.Worksheets("deduction").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("I:I"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("deduction").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("deduction").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("J:J"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("deduction").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    


    Dim header As Variant
    header = Array("External ID", _
    "Credit #", "Customer", "Date", "Posting Period", "Department", "Location", _
    "Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", "To Be Faxed", _
    "Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", _
    "Description", "Taxable", "PO details", "Apply_Applied", "Apply_payment")
    tgt.Range("A1:X1").Value = header

    ' start in row 2 to prevent copying the header
    ' copy the visible cells to our target range
    Set copyRange = src.Range("I2:I" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("C2")
    Set copyRange = src.Range("J2:J" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("G2")
    Set copyRange = src.Range("H2:J" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("O2")

    Set copyRange = src.Range("K2:K" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("R2")
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("S2")
    Set copyRange = src.Range("B2:B" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("T2")
    Set copyRange = src.Range("A2:A" & lastRow)
    copyRange.SpecialCells(xlCellTypeVisible).Copy tgt.Range("V2")

    ' fill range
    lastRowtgt = tgt.Range("C" & tgt.Rows.Count).End(xlUp).Row

    tgt.Range("D2:D" & lastRowtgt).Value = filedate
    tgt.Range("F2:F" & lastRowtgt).Value = "Dot com"
    tgt.Range("H2:H" & lastRowtgt).Value = "USD"
    tgt.Range("I2:I" & lastRowtgt).Value = "1"
    tgt.Range("J2:L" & lastRowtgt).Value = "FALSE"
    tgt.Range("M2:M" & lastRowtgt).Value = "Chargeback on CK#" & ach
    tgt.Range("N2:N" & lastRowtgt).Value = "Extra Deductions(except 4%)"
    tgt.Range("P2:P" & lastRowtgt).Value = "1"
    tgt.Range("Q2:Q" & lastRowtgt).Value = "Custom"
    tgt.Range("U2:U" & lastRowtgt).Value = "FALSE"

    ''generate keys
    
            tgt.Cells(2, 2).Value = 21
            tgt.Range("a2").Value = "CR0001"
            Dim m As Integer
            m = 3
            
            If lastRowtgt = 2 Then
                Else:
                If lastRowtgt = 3 Then

                    If tgt.Cells(m, 3) = tgt.Cells(m - 1, 3).Value Then
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2)
                        Else:
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2) + 1
                    End If
                    Sheets("tgt").Select
                    Range("A3").Select
                    ActiveCell.FormulaR1C1 = _
                    "=""CR00""&TEXT(RC[1]-20,""00"")"
                    
                    Else:
                
                    Do
                    If tgt.Cells(m, 3) = tgt.Cells(m - 1, 3).Value Then
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2)
                        m = m + 1
                        Else:
                        tgt.Cells(m, 2) = tgt.Cells(m - 1, 2) + 1
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

    
    tgt.Name = tempdate & "_WF " & Left(criteria, 2) & " extra deductions"

    tgt.Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name, _
    FileFormat:=xlCSV, CreateBackup:=True
    ActiveWorkbook.Close
        
    Application.DisplayAlerts = False
    Sheets(tgt.Index).Delete
    Application.DisplayAlerts = True
    
    src.AutoFilterMode = False

    
Next i

src.Range("l:l").ClearContents

        

End Sub







































