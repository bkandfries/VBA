Attribute VB_Name = "Module2"
Option Explicit
Private Function getConsolWorkBookName() As String
    Dim srcWB As Workbook
    Dim wbCnt As Integer
    For wbCnt = 1 To Workbooks.Count
        If InStr(1, Workbooks(wbCnt).Name, "Consol") Then
            Set srcWB = Workbooks(wbCnt)
        End If
    Next wbCnt
        
    getConsolWorkBookName = srcWB.Name
    Set srcWB = Nothing
End Function
Sub updateacts()
Attribute updateacts.VB_ProcData.VB_Invoke_Func = "U\n14"
        Dim trgws As Worksheet
        Dim consolws As Worksheet
        Dim consolwb As Workbook
        Dim c As Range
        
        If ActiveWorkbook.Name = ThisWorkbook.Name Or ActiveWorkbook.Name = getConsolWorkBookName Then
        MsgBox "select target wb"
        Exit Sub
        End If
        Dim consolName As String
        consolName = getConsolWorkBookName
        Application.Calculation = xlCalculationManual
        
        For Each trgws In ActiveWindow.SelectedSheets
        Application.StatusBar = "Working on " & trgws.Name
        If WorksheetExists(trgws.Name, Workbooks(consolName)) Then
                                        
                    Set consolws = Workbooks(getConsolWorkBookName).Worksheets(trgws.Name)
                    For Each c In consolws.Range("X1:X372")
                      '  Debug.Assert c.Row <> 6
                        If c.Row = 6 Then
                            trgws.Range(c.Address).Value = "ACT"
                        End If
                        
                        If c.Row = 9 Then
                            trgws.Range(c.Address).Value = 8
                        End If
                        If c.Interior.color = 16183538 Then
                            
                            With trgws.Range(c.Address)
                                    .Formula = "='[" & consolName & "]" & trgws.Name & "'!X" & c.Row
                                    .Interior.color = 16183538
                            End With
                          End If
                            
                    Next c
                    Else
                    trgws.Visible = xlSheetHidden
        End If
        trgws.Cells(1, "X").EntireColumn.Hidden = False
    Next trgws
    Application.StatusBar = ""
End Sub

Public Function IsWorksheet(ByVal sheetName As String, Optional ByVal Workbook As String) As Boolean
    Dim trgtWorkBookStr As String
    On Error Resume Next

    If Workbook <> "" Then
        If Workbooks(Workbook).Name <> "" Then
            trgtWorkBookStr = ActiveWorkbook.Name
        End If

        If Err Then Exit Function
        Err.Clear
    Else
        trgtWorkBookStr = ActiveWorkbook.Name
    End If

    On Error Resume Next
    IsWorksheet = Workbooks(trgtWorkBookStr).Sheets(sheetName).Name <> ""
    If Err And IsNumeric(sheetName) Then
        IsWorksheet = Workbooks(trgtWorkBookStr).Sheets(Val(sheetName)).Name <> ""
    End If

    On Error GoTo 0
End Function
Sub FixBudgetFileTF()
Attribute FixBudgetFileTF.VB_ProcData.VB_Invoke_Func = "B\n14"
    Dim srcWB As Workbook
    Dim SrcWS As Worksheet
    Dim TrgWB As Workbook
    Dim trgws As Worksheet
    Dim RefWS As Worksheet
    Dim SelectedSheets As Collection
    Dim Sheet As Variant
    Dim userInput As Variant
    Const checkConstant As String = "CONSTANT"  ' Replace "CONSTANT" with your specified constant
    
    
    
   If ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Sub
   
    
    ' Set references to your source and target workbooks and worksheets
    Set srcWB = ThisWorkbook  ' Assuming the source workbook is the one currently open
    Set SrcWS = srcWB.Sheets("02005")  ' Replace "SourceSheetName" with your source sheet name
    Set TrgWB = Workbooks(ActiveWorkbook.Name)  ' Replace "TargetWorkbookName.xlsx" with your target workbook name
    

    
    On Error GoTo 0  ' Error handling
    
    ' Disable screen updating, automatic calculations, and events
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    TrgWB.Activate
   
   
    ' Prompt user for input
    userInput = InputBox("Please enter a value for cell A1:", "User Input")
   
    ' Store selected sheetnames into array
    Dim shtNames() As String
    Dim cntr As Long
    
    cntr = 0
    
    Dim w As Worksheet
    For Each w In ActiveWindow.SelectedSheets
        ReDim Preserve shtNames(cntr)
        shtNames(cntr) = w.Name
        cntr = cntr + 1
    Next w
    
    ' Copy "REF" worksheet from SrcWB to TrgWB and hide it
    If Not WorksheetExists("REF", TrgWB) Then
        srcWB.Sheets("REF").Copy After:=TrgWB.Sheets(TrgWB.Sheets.Count)
        Set RefWS = TrgWB.Sheets(TrgWB.Sheets.Count)
        RefWS.Visible = xlSheetHidden
    End If

    ' Loop through each selected sheet in the collection
    Dim iLoop As Long
    Dim consolwb As Workbook
    Dim consolws As Worksheet
    
    For iLoop = LBound(shtNames) To UBound(shtNames)
        Set trgws = TrgWB.Worksheets(shtNames(iLoop))
        trgws.Activate
        Application.DisplayStatusBar = True
        Application.StatusBar = "Working on " & trgws.Name
        ' Set value of cell A1 to user input
        trgws.Cells(1, 1).Value = userInput
        If userInput = "" Then
        MsgBox "enter something"
        Exit Sub
        End If
        
        
        With trgws
            
            
            SrcWS.Range("A394:CF500").Copy Destination:=.Range("A394:CF500")
            SrcWS.Range("BQ1:CF500").Copy Destination:=.Range("BQ1:CF500")
            SrcWS.Range("c53:Be55").Copy Destination:=.Range("c53:Be55")
            SrcWS.Range("AS3:BE3").Copy Destination:=.Range("AS3:BE3")
            
            
            .Range("AS26").Resize(1, 12).Formula = SrcWS.Range("AS26").Resize(1, 12).Formula
            .Range("AS27").Resize(1, 12).Formula = SrcWS.Range("AS27").Resize(1, 12).Formula
            .Range("AS59").Resize(1, 12).Formula = SrcWS.Range("AS59").Resize(1, 12).Formula
            .Range("AS67").Resize(1, 12).Formula = SrcWS.Range("AS67").Resize(1, 12).Formula
            .Range("AS75").Resize(1, 12).Formula = SrcWS.Range("AS75").Resize(1, 12).Formula
            .Range("AS95").Resize(1, 12).Formula = SrcWS.Range("AS95").Resize(1, 12).Formula
            .Range("AS105").Resize(1, 12).Formula = SrcWS.Range("AS105").Resize(1, 12).Formula
             
             SrcWS.Range("as98").Resize(1, 12).Copy
            .Range("as98").Resize(1, 12).PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
             SrcWS.Range("AS26").Resize(3, 12).Copy
            .Range("AS26").Resize(3, 12).PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
             SrcWS.Range("x14:ab352").Copy
            .Range("x14:ab352").PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
             
             SrcWS.Range("as95").Resize(1, 12).Copy
            .Range("as95").Resize(1, 12).PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
             
             
             Application.CutCopyMode = False
            
            .Rows("486:" & .Rows.Count).Clear
            .Columns("CG:EG").Clear
            
            
        End With
        

        Dim clr As Variant
        
        
        ' Loop through rows 58 to 352
        Dim r As Range
        For Each r In trgws.Range("AS53:BD353").Rows
        If r.Row <> 58 And r.Row <> 59 And r.Row <> 60 And r.Row <> 61 And r.Row <> 63 And r.Row <> 67 And r.Row <> 93 And r.Row <> 94 Then
            clr = -1
            If r.Cells(1, 1).HasFormula Then
                          If Not r.Cells(1, 1).Interior.color = 65535 And Not r.Cells(1, 1).Interior.color = 855309 Then
                            clr = r.Cells(1, 1).Interior.color
                            Else
                             clr = -1
                            End If
                            
                            
                          
                SrcWS.Range(r.Address).Copy Destination:=r
                If clr <> -1 Then r.Interior.color = clr
                
            End If
        End If
        Next r
        
        Set r = Nothing
        
        Dim n As Long

        
        
        trgws.Range("AS67").Formula = trgws.Range("AS67").Formula & "/1.5*IFERROR(AS$403,1.5)"
        trgws.Range("AS67").Resize(1, 12).FillRight
        Dim mgmtsal As String
        
        
        ' Copy column widths
      Dim c As Range
        Dim SalSTR As String
        Dim vacastr As String
            trgws.Range("AS93").Value2 = 0
            trgws.Range("AS93").Interior.color = xlNone
            trgws.Range("AS93").Resize(1, 12).FillRight
              
              trgws.Range("AS94").Formula = "=IFERROR((SUMIFS('Compensation Report'!$AC:$AC,'Compensation Report'!$AJ:$AJ,""602"",'Compensation Report'!$X:$X,""Salaried"",'Compensation Report'!$AF:$AF,$A$3)/364*DAY(EOMONTH(DATE(2024,AS$9,1),0))),0)+IFERROR((SUMIFS('Compensation Report'!$AC:$AC,'Compensation Report'!$AJ:$AJ,""602"",'Compensation Report'!$X:$X,""Hourly"",'Compensation Report'!$AF:$AF,$A$3)/364*DAY(EOMONTH(DATE(2024,AS$9,1),0))),0)"
              trgws.Range("AS94").Resize(1, 12).FillRight
            
              trgws.Range("AS98").Formula = "=IFERROR(SUM($BZ$94) * $BV98 / SUM($BV$94) * IF(Q$9 = 0, $BV98 / $BV$398 * AS$398, Q98) / $BV98, 0) * IFERROR(1 + $BT98, 1) * (IFERROR((SUMIFS('Compensation Report'!$AC:$AC, 'Compensation Report'!$AJ:$AJ, ""602"", 'Compensation Report'!$X:$X, ""Hourly"", 'Compensation Report'!$AF:$AF, $A$3) / 364 * DAY(EOMONTH(DATE(2024, AS$9, 1), 0))) / ((SUMIFS('Compensation Report'!$AC:$AC, 'Compensation Report'!$AJ:$AJ, ""602"", 'Compensation Report'!$X:$X, ""Salaried"", 'Compensation Report'!$AF:$AF, $A$3) / 364 * DAY(EOMONTH(DATE(2024, AS$9, 1), 0))) + (SUMIFS('Compensation Report'!$AC:$AC, 'Compensation Report'!$AJ:$AJ, ""602"", 'Compensation Report'!$X:$X, ""Hourly"", 'Compensation Report'!$AF:$AF, $A$3) / 364 * DAY(EOMONTH(DATE(2024, AS$9, 1), 0)))), 0))"
                trgws.Range("AS98").Resize(1, 12).FillRight

        Dim colRng As Range
        For Each colRng In SrcWS.UsedRange.Columns
            trgws.Columns(colRng.Column).ColumnWidth = colRng.ColumnWidth
        Next colRng
   
    
        Set colRng = Nothing
        Set trgws = Nothing

    Next iLoop
    Application.DisplayStatusBar = True
    ' Run the DeleteInvalidNames macro from the PERSONAL.xlsb workbook
    Application.Run "PERSONAL.xlsb!DeleteInvalidNames"
    
    ' Change links from SrcWB to TrgWB
  Dim linkSources As Variant
    linkSources = TrgWB.linkSources(xlExcelLinks)  ' Get an array of source file names
    
    If Not IsEmpty(linkSources) Then
        Dim i As Integer
        For i = LBound(linkSources) To UBound(linkSources)
            Dim sourceFileName As String
            If linkSources(i) = srcWB.FullName Then
            On Error Resume Next
            TrgWB.ChangeLink Name:=linkSources(i), NewName:=TrgWB.FullName, Type:=xlExcelLinks
            On Error GoTo 0
            
            End If
        Next i
   End If
    Exit Sub  ' Exit before reaching the error handler
    
  
    
ErrorHandler:
    ' Reset settings in case of an error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Application.CalculateFull
    Application.Calculation = xlCalculationAutomatic
    
    ' Display error message
    MsgBox Err.Description, vbExclamation
    
    ' Release object references
    Set trgws = Nothing
    Set SrcWS = Nothing
    Set TrgWB = Nothing
    Set srcWB = Nothing
    Set RefWS = Nothing
    Set SelectedSheets = Nothing
    Set consolwb = Nothing
    Set consolws = Nothing
End Sub

Sub CPFormats2()
Attribute CPFormats2.VB_ProcData.VB_Invoke_Func = "W\n14"

    Dim srcWB As Workbook
    Dim SrcWS As Worksheet
    Dim trgtWB As Workbook
    Dim TrgtWS As Worksheet
            
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Activate Other WB First"
            Exit Sub
    End If
    
    ' Set references to your source and target workbooks and worksheets
    Set srcWB = ThisWorkbook  ' Assuming the source workbook is the one currently open
    Set trgtWB = Workbooks(ActiveWorkbook.Name)  ' Replace "TargetWorkbookName.xlsx" with your target workbook name
    
    Dim sheetNames(0 To 2) As String
    sheetNames(0) = "Compensation Report"
    sheetNames(1) = "Rev Events Inputs"
    sheetNames(2) = "GenLabor Inputs"
    Dim rng As Range
    
    
    
    
    
    
    
    
    TurnOffToBegin
    Dim i As Long
    
    Dim rg As Range
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        If Not WorksheetExists(sheetNames(i)) Then
            MsgBox (sheetNames(i) & " not found")
            Err.Clear
            GoTo nextseq:
        End If
        
        
        Set TrgtWS = trgtWB.Worksheets(sheetNames(i))
        Set SrcWS = srcWB.Worksheets(sheetNames(i))
        
        TrgtWS.Activate

        If TrgtWS.AutoFilterMode = True Then
            Range(TrgtWS.AutoFilter.Range.Address).AutoFilter
        End If
            If sheetNames(i) = "Rev Events Inputs" Then
                TrgtWS.Range("A1", TrgtWS.Cells.SpecialCells(xlLastCell)).AutoFilter
                ElseIf sheetNames(i) = "GenLabor Inputs" Then
                TrgtWS.Range("A1", TrgtWS.Cells.SpecialCells(xlLastCell)).AutoFilter
                
            End If

        
        
        
        
        TrgtWS.Range("A1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        TrgtWS.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        Application.CutCopyMode = False
        
            
        TrgtWS.UsedRange.ColumnWidth = SrcWS.UsedRange.ColumnWidth
        TrgtWS.UsedRange.RowHeight = SrcWS.UsedRange.RowHeight

        TrgtWS.Activate
        TrgtWS.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.Zoom = "90"
        If sheetNames(i) = "GenLabor Inputs" Then
            SrcWS.Range("K1").Copy Destination:=TrgtWS.Range("K1")
            SrcWS.Range("An2:AT3").Copy Destination:=TrgtWS.Range("An2:AT3")
            TrgtWS.Range("Ao3").Resize(TrgtWS.Range("A4").CurrentRegion.Rows.Count - 1, 6).FillDown
            If TrgtWS.AutoFilterMode = True Then TrgtWS.AutoFilterMode = False
            TrgtWS.Range("A4").CurrentRegion.Offset(1, 0).AutoFilter
            
            TrgtWS.Calculate
            
        End If
        
                
        If sheetNames(i) = "Rev Events Inputs" Then
            SrcWS.Range("S1:X2").Copy Destination:=TrgtWS.Range("s1:X2")
            TrgtWS.Range("t2").Resize(TrgtWS.Range("A4").CurrentRegion.Rows.Count - 1, 6).FillDown
            If TrgtWS.AutoFilterMode = True Then TrgtWS.AutoFilterMode = False
            TrgtWS.Range("A4").CurrentRegion.AutoFilter
            TrgtWS.Calculate
                
            
        End If
        Application.CutCopyMode = False
        
        TrgtWS.Activate
        Range("A1").Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        
        
nextseq:
        ' Release objects for the next iteration
        Set TrgtWS = Nothing
        Set SrcWS = Nothing
        Set rg = Nothing
        
    Next i

    ' Release objects
    Set trgtWB = Nothing
    Set srcWB = Nothing
    TurnOnToEnd
    
    End Sub

Sub TurnOffToBegin()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook
    ' Turn off Excel functionality to improve performance.
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    ' Note: this is a sheet-level setting.
    ActiveSheet.DisplayPageBreaks = False
    Application.AutoRecover.Enabled = False
    
    For Each ws In wb.Worksheets
    
        Debug.Assert (WorksheetExists(ws.Name))
        If Not ws.Visible = xlSheetVeryHidden Then
            ws.DisplayPageBreaks = False
        End If
    
    Next ws
    Set ws = Nothing
    Set wb = Nothing
End Sub

Private Sub TurnOnToEnd()
    ' Restore Excel settings to original state.
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.AutoRecover.Enabled = True
    ' If Not IsEmpty(calcState) Then Application.Calculation = calcState
End Sub

Public Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Public Function lastRow(sh As Worksheet)
    On Error Resume Next
    lastRow = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            searchdirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Private Function lastCol(sh As Worksheet)
    On Error Resume Next
    lastCol = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            searchdirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function



Public Function LastRange(Optional Worksheet As Variant, Optional SkipFirstRows As Integer = 0) As Range
        Dim ws As Worksheet
        If IsMissing(Worksheet) Then
            Set ws = ActiveSheet
        ElseIf VarType(Worksheet) = vbString Then
            Set ws = Worksheets(Worksheet)
            
        
        Else
            If IsObject(Worksheet) Then
                
                Set ws = Worksheet
            Else
              Set ws = ActiveWorkbook.Worksheets(Worksheet)
            End If
       End If
        Dim rng As Range
       
       
        With ws
            Set rng = .Range(.Cells(1, 1), .Cells(lastRow(ws), lastCol(ws)))
        End With
               
        If SkipFirstRows > 0 Then
        Set LastRange = rng.Offset(SkipFirstRows).Resize(rng.Rows.Count - SkipFirstRows, rng.Columns.Count)
        Else
        Set LastRange = rng
        End If
        
End Function





