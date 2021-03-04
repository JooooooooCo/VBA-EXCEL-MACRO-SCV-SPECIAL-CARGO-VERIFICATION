Attribute VB_Name = "SPECIAL_CARGO_VERIF"
Public selectedbkgtype As String

Sub SPECIAL_CARGO_VERIF()
Attribute SPECIAL_CARGO_VERIF.VB_ProcData.VB_Invoke_Func = "P\n14"

' SELECT BOOKING TYPE - NORMAL OR BRCAB
    selectedbkgtype = ""
   
    Application.ScreenUpdating = True
    BKGTYPEINPUT.StartUpPosition = 0
    BKGTYPEINPUT.Left = Application.Left + (0.5 * Application.Width) - (0.5 * BKGTYPEINPUT.Width)
    BKGTYPEINPUT.Top = Application.Top + (0.5 * Application.Height) - (0.5 * BKGTYPEINPUT.Height)
    BKGTYPEINPUT.Show
    Application.ScreenUpdating = False
    
If selectedbkgtype <> "" Then
    

        'OPENING CARGO READINESS REPORT
          Dim BKGTYPE As Integer
          Dim directory As String, fileName As String, sheet As Worksheet, total As Integer
          Dim fd As Office.FileDialog
        
          Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
          With fd
            .AllowMultiSelect = False
            .Title = "Please select the file."
            .Filters.Clear
            .Filters.Add "Excel 2003", "*.xls?"
        
            If .Show = True Then
              fileName = dir(.SelectedItems(1))
        
            End If
          End With
        
          Application.ScreenUpdating = False
          Application.DisplayAlerts = False
        
          Workbooks.Open (fileName)
          
        'PREPARING PLAN
            Rows("1:1").Select
            Selection.Delete Shift:=xlUp
            Rows("5:17").Select
            Selection.Delete Shift:=xlUp
            ActiveWindow.FreezePanes = False
        
        ' Searching fields for validation
            Dim searchvalidationfield(1 To 12) As String
            Dim i As Integer
            Dim headervalidationfield As Range
            Dim headervalidationfieldcolumn As String
                
            i = 1
            searchvalidationfield(1) = "DG document enclosed"
            searchvalidationfield(2) = "Remarks for Load List and Terminals"
            searchvalidationfield(3) = "TP Invalidity Reasons"
            searchvalidationfield(4) = "Amendment / Cancellation request"
            searchvalidationfield(5) = "Additional Activity Reference"
            searchvalidationfield(6) = "EQU Number"
            searchvalidationfield(7) = "Booking Equipment Line Item"
            searchvalidationfield(8) = "Commodity"
            searchvalidationfield(9) = "Cargo Type"
            searchvalidationfield(10) = "Status"
            searchvalidationfield(11) = "Booking Number"
            searchvalidationfield(12) = "Booking Type"
            
        Do While i <= 12
            
            Set headervalidationfield = Cells.Find(What:=searchvalidationfield(i), LookIn:=xlFormulas, LookAt _
                :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
        
            If headervalidationfield Is Nothing Then
                
                MsgBox ("The column * " & searchvalidationfield(i) & " * was not found in the selected file / report." _
                & vbCrLf & vbCrLf & "Please verify if the Cargo Readiness Report contains all the below columns:" _
                & vbCrLf & vbCrLf & "Booking Type, Booking Number, Status, Cargo Type, Commodity, Booking Equipment Line Item, EQU Number, Additional Activity Reference,Amendment / Cancellation request, TP Invalidity Reasons, Remarks for Load List and Terminals or DG document enclosed.")
                Workbooks(fileName).Close SaveChanges:=False
                End
                
                Else
        
                headervalidationfieldcolumn = headervalidationfield.Address & ":" & headervalidationfield.Address
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "$", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "0", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "1", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "2", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "3", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "4", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "5", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "6", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "7", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "8", "")
                headervalidationfieldcolumn = Replace(headervalidationfieldcolumn, "9", "")
                Columns(headervalidationfieldcolumn).Select
                Selection.Copy
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
        
            End If
            
            i = i + 1
        Loop
            
            Columns("M:M").Select
            Application.CutCopyMode = False
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        
        'INCL HEADER
            Range("C1:C4").Select
            Selection.Copy
            Range("F1:F4").Select
            ActiveSheet.Paste
            Range("M5").Select
            ActiveCell.FormulaR1C1 = "CSVE Action Required"
            Range("N5").Select
            ActiveCell.FormulaR1C1 = "Bkg Number VS Bkg Movement"
            Range("O5").Select
            ActiveCell.FormulaR1C1 = "Booking Status "
            Range("P5").Select
            ActiveCell.FormulaR1C1 = "Coal / Charcoal / Wetblue / Scrap / Pharma Approval"
            Range("Q5").Select
            ActiveCell.FormulaR1C1 = "Amendment / Cancellation Pending"
            Range("R5").Select
            ActiveCell.FormulaR1C1 = "TP Invalid"
            Range("S5").Select
            ActiveCell.FormulaR1C1 = "DG final docs"
            
        
        'RUN VALIDATIONS
            Range("B6").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Range("M6").Select
            ActiveSheet.Paste
            Range("M6").Select
        
                ' IF BKG NORMAL
                If selectedbkgtype = "NORMAL" Then
                    
                    ActiveCell.FormulaR1C1 = _
                        "=IF(RC[-12]=""NORMAL"",IF(OR(RC[1]<>""OOOKKK"",RC[2]<>""OOOKKK"",RC[3]<>""OOOKKK"",RC[4]<>""OOOKKK"",RC[5]<>""OOOKKK"",RC[6]<>""OOOKKK""),""YES"",""NO""),""NO"")"
                
                    Range("G1").Select
                    ActiveCell.FormulaR1C1 = "Booking Type: NORMAL"
                
                End If
                
                'IF BKG BRCAB
                If selectedbkgtype = "BRCAB" Then
            
                    ActiveCell.FormulaR1C1 = _
                        "=IF(RC[-12]=""BRCAB"",IF(OR(RC[1]<>""OOOKKK"",RC[2]<>""OOOKKK"",RC[3]<>""OOOKKK"",RC[4]<>""OOOKKK"",RC[5]<>""OOOKKK"",RC[6]<>""OOOKKK""),""YES"",""NO""),""NO"")"
                
                    Range("G1").Select
                    ActiveCell.FormulaR1C1 = "Booking Type: BRCAB"
                
                End If
        
        
            Range("N6").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(RC[-7]="""",""OOOKKK"",IF(RC[-12]=LEFT(RC[-8],10),""OOOKKK"",CONCATENATE(""Booking Number: "",RC[-12],CHAR(10),""Booking Movement: "", LEFT(RC[-8],10))))"
            Range("O6").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(RC[-12]=""CONFIRMED"",""OOOKKK"",CONCATENATE(""Error: "",RC[-12]))"
            Range("P6").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(OR(ISERR(SEARCH(""COAL"",RC[-11]))=FALSE,ISERR(SEARCH(""BLUE,WET"",RC[-11]))=FALSE,ISERR(SEARCH(""SCRAP"",RC[-11]))=FALSE,ISERR(SEARCH(""PHARMA"",RC[-11]))=FALSE),CONCATENATE(""Approval Remarks: "",RC[-5],CHAR(10),CHAR(10),""Commodity: "",RC[-11]),""OOOKKK"")"

            'FORMULA DESABILITADA DEVIDO A FALHA NA INTERPRETACAO DA FUNCIONALIDADE DA COLUNA DE AMEND E CANCEL.
            'DEVERA SER HABILITADA APOS CONCLUSAO DO CR.
            Range("Q6").Select
            ActiveCell.FormulaR1C1 = "OOOKKK"
           'ActiveCell.FormulaR1C1 = _
           '    "=IF(RC[-8]="""",""OOOKKK"",CONCATENATE(""Pending: "",RC[-8]))"
            Range("R6").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(RC[-8]="""",""OOOKKK"",CONCATENATE(""TP Invalid. Reason: "",RC[-8]))"
            Range("S6").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(ISERR(SEARCH(""DG"",RC[-15])=FALSE),""OOOKKK"",IF(RC[-7]=""Y"",""OOOKKK"",""DG document not enclosed""))"
            Range("M6:S6").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
    
            Application.Calculation = xlAutomatic

            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        ' REMOVE "OOOKKK"
            Range("M6:S6").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Replace What:="OOOKKK", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
        
        'GENERAL FORMATATION
            Cells.Select
            With Selection.Font
                .Name = "Arial"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            ActiveWindow.DisplayGridlines = False
        
            Range("A6:FZ6").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = xlThin
            End With
 
        'ADJUST HEADER VERIFICATION
            Range("M5:S5").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
            Range("A5:FZ5").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            Range("A5:FZ5").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        
            Range("M4:S4").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Font.Bold = True
            
        ' Searching Shipper column
            Dim searchadditionalfield(1 To 2) As String
            'Dim i As Integer (duplicated)
            Dim headeradditionalfield As Range
            Dim headeradditionalfieldcolumn As String
                
            i = 1
            searchadditionalfield(1) = "Booking Agreement Party"
            searchadditionalfield(2) = "Shipper"
                        
            Do While i <= 2
                
                Set headeradditionalfield = Cells.Find(What:=searchadditionalfield(i), LookIn:=xlFormulas, LookAt _
                    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                    False, SearchFormat:=False)
            
                If headeradditionalfield Is Nothing Then
                    
                    MsgBox ("The column * " & searchadditionalfield(i) & " * was not found in the selected file / report." _
                    & vbCrLf & vbCrLf & "Please verify if the Cargo Readiness Report contains all the below columns:" _
                    & vbCrLf & vbCrLf & "Booking Type, Booking Number, Status, Cargo Type, Commodity, Booking Equipment Line Item, EQU Number, Additional Activity Reference,Amendment / Cancellation request, TP Invalidity Reasons, Remarks for Load List and Terminals or DG document enclosed.")
                    Workbooks(fileName).Close SaveChanges:=False
                    End
                    
                    Else
                    
                    headeradditionalfieldcolumn = headeradditionalfield.Address & ":" & headeradditionalfield.Address
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "$", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "0", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "1", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "2", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "3", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "4", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "5", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "6", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "7", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "8", "")
                    headeradditionalfieldcolumn = Replace(headeradditionalfieldcolumn, "9", "")
                    Columns(headeradditionalfieldcolumn).Select
                    Selection.Copy
                    Columns("T:T").Select
                    Selection.Insert Shift:=xlToRight
            
                End If
                
                i = i + 1
            Loop
            
        'ADJUST COLUMN WIDTH
            Columns("G:G").ColumnWidth = 16.5
            Columns("M:M").ColumnWidth = 10.17
            Columns("N:N").ColumnWidth = 22.86
            Columns("O:O").ColumnWidth = 17.71
            Columns("P:P").ColumnWidth = 30.14
            Columns("Q:Q").ColumnWidth = 18.86
            Columns("R:R").ColumnWidth = 38.5
            Columns("S:S").ColumnWidth = 26.57
            Columns("T:T").ColumnWidth = 25.5
            Columns("U:U").ColumnWidth = 25.5
            
        ' ADJUST HEADER VVD
            Range("A2").Select
            ActiveCell.FormulaR1C1 = "Op. Voyage"
            Range("A3").Select
            ActiveCell.FormulaR1C1 = "Coml. Voyage"
            Rows("1:4").RowHeight = 11.25

        'FILTER ONLY CSVE ACTION REQUIRED
            Range("A6").Select
            ActiveWindow.FreezePanes = True
            
            If ActiveSheet.AutoFilterMode Then
            'FILTER ON
            Rows("5:5").Select
            Selection.AutoFilter
            Selection.AutoFilter Field:=13, Criteria1:="YES"
            Else
            'FILTER OFF
            Rows("5:5").Select
            Selection.AutoFilter Field:=13, Criteria1:="YES"
            End If
        
        'Add Qtt Actions info
            Range("M4").Select
            ActiveCell.FormulaR1C1 = "Qtt Actions:"
            Range("N4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
            Range("O4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
            Range("P4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
            Range("Q4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
            Range("R4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
            Range("S4").Select
            ActiveCell.FormulaR1C1 = "=SUBTOTAL(103,R[2]C:R[5000]C)"
 
            Application.Calculation = xlAutomatic
            
            Range("M4:S4").Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False

        ' Hidden not used columns
            Cells.Select
            Selection.EntireColumn.Hidden = False
            Selection.EntireRow.Hidden = False
            
            Range("C:E,H:L").Select
            Selection.EntireColumn.Hidden = True
            Columns("V:GB").Select
            Selection.EntireColumn.Hidden = True
 
    Range("A1").Select

End If

End Sub



