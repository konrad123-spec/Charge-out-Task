Attribute VB_Name = "charge_out_petla_new"
Sub main_new()

    ' Declaring SAP GUI and Excel objects
    Dim sapGuiApp As Object
    Dim excelApp As Object
    Dim excelSheet As Object
    Dim session As Object
    Dim sapValue As String
    Dim sapValue1 As String
    Dim Bu As String
    Dim PG As String
    
    ' Retrieving credentials and description from the "Macro" worksheet
    Dim Crudencials As String
    Crudencials = Worksheets("Macro").Range("H1")
    
    Dim Description As String
    Description = Worksheets("Macro").Range("H2")
    
    ' Retrieving the file path
    Dim Path As String
    Path = Worksheets("Macro").Range("H3").Value
    
    ' Loading nodes from the "Macro" worksheet
    Dim Nodes() As Variant
    Worksheets("Macro").Activate
    Nodes = Range("C3").CurrentRegion
    
    ' Setting up SAP session
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    
    ' Opening the report in SAP
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "gr55"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = "N016"
    session.findById("wnd[0]/tbar[1]/btn[20]").press
    session.findById("wnd[0]/usr/radONLYUSER").Select
    session.findById("wnd[0]/usr/txtLTEXT").Text = Description
    session.findById("wnd[0]/usr/txtI_USER").Text = Crudencials
    session.findById("wnd[0]/usr/ctxtR_RGJNR-LOW").Text = "N016"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "LTEXT"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
    
    ' Setting column width and decimal places
    session.findById("wnd[0]/usr/lbl[5,8]").SetFocus
    session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
    session.findById("wnd[1]/usr/txtRGRWF-RTITW").Text = "65"
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/mbar/menu[5]/menu[1]").Select
    session.findById("wnd[1]/usr/ctxtLGRWO-SUM_FROM").Text = "2"
    session.findById("wnd[1]/usr/ctxtLGRWO-SUM_TO").Text = "2"
    session.findById("wnd[1]/usr/ctxtLGRWO-SUM_TO").SetFocus
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/usr/lbl[71,8]").SetFocus
    session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select
    session.findById("wnd[1]/usr/txtRGRWF-COLWD").Text = "16"
    session.findById("wnd[1]/usr/cmbRGRWF-ROUND").Key = "0"
    session.findById("wnd[1]/usr/ctxtRGRWF-DECIP").Text = "2"
    session.findById("wnd[1]").sendVKey 0
    
    ' Categories processing
    Dim Categories() As Variant
    Categories = Array("UA3201 - Charge-out non-order rel COS AB", "UA3211 - Charge-out non-order rel COS IN")
                    
    For Category = LBound(Categories) To UBound(Categories)
    
        ' Calling Exporting subroutine
        Call Exporting(CStr(Categories(Category)), Path)
    
        ' Processing nodes
        For Node = LBound(Nodes) To UBound(Nodes)
        
            If Left(Nodes(Node, 1), 2) = "IA" Then
                Bu = Nodes(Node, 1)
            Else
                PG = Nodes(Node, 1)
                Call Accounts(Bu, PG, CStr(Categories(Category)), Path)
            End If
        
        Next Node
        
        Range("X1").Copy
        Workbooks("Export.xlsx").Close True
    
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    Next Category
    
    MsgBox "Done"

End Sub

Sub Exporting(Category As String, Path As String)

    ' Setting up SAP session
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

    ' Opening accounts in SAP
    session.findById("wnd[0]").sendVKey 71
    session.findById("wnd[1]/usr/txtRSYSF-STRING").Text = Category
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' Error handling if category is missing
    On Error GoTo Finish
    session.findById("wnd[2]/usr/lbl[17,2]").SetFocus
    session.findById("wnd[2]").sendVKey 2
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[1]/usr/lbl[1,2]").SetFocus
    session.findById("wnd[0]").sendVKey 2

    ' Resetting error handling
    On Error GoTo 0

    ' Layout adjustments and date selection
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[1]/btn[33]").press
    session.findById("wnd[1]/tbar[0]/btn[71]").press
    session.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "Charge out new"
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[3]/usr/lbl[15,2]").SetFocus
    session.findById("wnd[3]").sendVKey 2
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ' Changing the date range for the report
    Dim currentDate As Date
    Dim firstDay As Date
    Dim lastDay As Date
    
    currentDate = Date
    firstDay = Format(DateSerial(Year(currentDate), Month(currentDate) - 1, 1), "dd.mm.yyyy")
    lastDay = Format(DateSerial(Year(currentDate), Month(currentDate), 0), "dd.mm.yyyy")

    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/btnB_SEARCH").press
    session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "data di reg"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/btnAPP_WL_SING").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = firstDay
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = lastDay
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]").maximize
    
    ' Exporting data to file
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Export.XLSX"
    session.findById("wnd[1]/tbar[0]/btn[7]").press
    
    ' Opening and closing the exported Excel file
    Workbooks.Open Path & "\Export.xlsx"
    Workbooks("Export.xlsx").Close False
    Workbooks.Open Path & "\Export.xlsx"
    
    Windows("Export.xlsx").Activate
    
    Exit Sub
    
Finish:
    ' Error handling for issues with export
    On Error GoTo -1
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[12]").press
    Exit Sub

End Sub

Sub Accounts(Bu As String, PG As String, Category As String, Path As String)

    ' Error handling
    On Error GoTo 0
    
    ' Shortening category names for Chargeout file
    Dim CategoryTipo As String
    
    Select Case Category
        Case "UA3201 - Charge-out non-order rel COS AB"
            CategoryTipo = "CHARGE OUT COS AB"
        Case "UA3211 - Charge-out non-order rel COS IN"
            CategoryTipo = "CHARGE OUT COS IN"
        Case "UA3221 - Charge-out non-order rel GA "
            CategoryTipo = "CHARGE OUT G&A"
        Case "UA3241 - Charge-out non-order rel Sales"
            CategoryTipo = "CHARGE OUT SALES"
    End Select
    
    ' Activating Export workbook
    Windows("Export.xlsx").Activate
    
    ' Filtering the Export file
    ActiveSheet.Range("S1").CurrentRegion.AutoFilter Field:=18, Criteria1:=Bu
    ActiveSheet.Range("S1").CurrentRegion.AutoFilter Field:=19, Criteria1:=Right(PG, 4)
    
    Dim row As Integer
    Dim lastRow As Integer
    
    ' Check if any rows remain visible after filtering
    If Range("S1").End(xlDown).Row <> Rows.Count Then
        lastRow = Range("S1").End(xlDown).Row
    Else
        Exit Sub
    End If
    
    ' Deleting rows with "0" values
    For Each viscell In Range("O1:O" & lastRow).SpecialCells(xlCellTypeVisible)
        If viscell.Value = 0 Then
            viscell.EntireRow.Delete
        End If
    Next viscell
    
    ' Rechecking for visible rows after deletion
    If Range("S1").End(xlDown).Row <> Rows.Count Then
        lastRow = Range("S1").End(xlDown).Row
    Else
        Exit Sub
    End If
    
    ' Calculating the number of visible rows
    Dim numberOfRows As Integer
    numberOfRows = Range("O2:Q" & lastRow).SpecialCells(xlCellTypeVisible).Cells.Count / 3
    
    ' Filtering the Chargeout file and copying data
    Windows("TABELA.xlsm").Activate
    Workbooks("TABELA.xlsm").Worksheets("datacharge").Activate
    
    ActiveSheet.Range("A1:H" & Range("H1").End(xlDown).Row).AutoFilter Field:=2, Criteria1:=Right(PG, 4)
    ActiveSheet.Range("A1:H" & Range("H1").End(xlDown).Row).AutoFilter Field:=3, Criteria1:=CategoryTipo
    
    ' Copying rows to the Chargeout file
    Dim startingRow As Integer
    If Range("C1").End(xlDown).Row = Rows.Count Then
        MsgBox "Something is wrong"
        Exit Sub
    Else
        startingRow = Range("C1").End(xlDown).Row
    End If
    
    Range("A" & startingRow).EntireRow.Copy
    Rows(startingRow + 1).Select
    Dim i As Integer
    i = 2
    
    Do While i <= numberOfRows
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A" & startingRow).EntireRow.Copy
        i = i + 1
    Loop
    
    ' Pasting values from Export to Chargeout
    Windows("Export.xlsx").Activate
    Range("O2:Q" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    Windows("TABELA.xlsm").Activate
    Workbooks("TABELA.xlsm").Worksheets("datacharge").Range("F" & startingRow).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Deleting rows with empty or zero values
    For row = startingRow To startingRow + numberOfRows
        If Range("F" & row).Value + Range("F" & row + 1).Value = 0 And Not IsEmpty(Range("F" & row)) And Range("H" & row) = Range("H" & row + 1) Then
            Range("F" & row & ":H" & row + 1).ClearContents
        End If
    Next row
    
    ' Deleting empty rows
    Dim placeHolder As Integer
    placeHolder = 0
    For row = startingRow To startingRow + numberOfRows
        If (IsEmpty(Range("F" & row)) And Range("C" & row) = CategoryTipo) Then
            If Right(Range("B" & row), 4) = Right(PG, 4) Then
                Rows(row).Delete
                row = row - 1
                placeHolder = placeHolder + 1
            End If
        End If
        If placeHolder >= 1000 Then Exit Sub
    Next row
    
    Exit Sub

End Sub

Sub GA()
    ' Handles "UA3221 - Charge-out non-order rel GA "
    ' This is similar to the main_new routine, but specific for the GA category
    
    Dim sapGuiApp As Object
    Dim excelApp As Object
    Dim excelSheet As Object
    Dim session As Object
    Dim sapValue As String
    Dim sapValue1 As String
    Dim Bu As String
    Dim PG As String
    
    Dim Crudencials As String
    Crudencials = Worksheets("Macro").Range("H1")
    
    Dim Description As String
    Description = Worksheets("Macro").Range("H2")
    
    Dim Path As String
    Path = Worksheets("Macro").Range("H3").Value
    
    Dim Nodes() As Variant
    Worksheets("Macro").Activate
    Nodes = Range("C3").CurrentRegion

    
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    
    ' Setting the category to GA
    Dim Category As String
    Category = "UA3221 - Charge-out non-order rel GA "
    
    ' Exporting GA data and calling Accounts subroutine
    Call Exporting(Category, Path)
    
    For Node = LBound(Nodes) To UBound(Nodes)
        
        If Left(Nodes(Node, 1), 2) = "IA" Then
            Bu = Nodes(Node, 1)
        Else
            PG = Nodes(Node, 1)
            Call Accounts(Bu, PG, CStr(Category), Path)
        End If
    
    Next Node
    
    ' Closing the workbook
    Range("X1").Copy
    Workbooks("Export.xlsx").Close True

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    MsgBox "Done"
    
End Sub

Sub SALES()
    ' Handles "UA3241 - Charge-out non-order rel Sales"
    ' Similar to the GA subroutine but specific to the Sales category

    Dim sapGuiApp As Object
    Dim excelApp As Object
    Dim excelSheet As Object
    Dim session As Object
    Dim sapValue As String
    Dim sapValue1 As String
    Dim Bu As String
    Dim PG As String
    
    Dim Crudencials As String
    Crudencials = Worksheets("Macro").Range("H1")
    
    Dim Description As String
    Description = Worksheets("Macro").Range("H2")
    
    Dim Path As String
    Path = Worksheets("Macro").Range("H3").Value
    
    Dim Nodes() As Variant
    Worksheets("Macro").Activate
    Nodes = Range("C3").CurrentRegion
    
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    
    ' Setting the category to Sales
    Dim Category As String
    Category = "UA3241 - Charge-out non-order rel Sales"
    
    ' Exporting Sales data and calling Accounts subroutine
    Call Exporting(Category, Path)
    
    For Node = LBound(Nodes) To UBound(Nodes)
        
        If Left(Nodes(Node, 1), 2) = "IA" Then
            Bu = Nodes(Node, 1)
        Else
            PG = Nodes(Node, 1)
            Call Accounts(Bu, PG, CStr(Category), Path)
        End If
    
    Next Node
    
    ' Closing the workbook
    Range("X1").Copy
    Workbooks("Export.xlsx").Close True

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    MsgBox "Done"
    
End Sub

Sub Abaccli()

    ' Starting by showing every row with value. lastRowUltimate is the last row in this table.
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter Field:=6, Operator:=xlFilterNoFill
    lastRowUltimate = Range("H1").End(xlDown).Row
    
    ' Defining long names of categories as written in SAP
    Dim Categories() As Variant
    Categories = Array("UA3201 - Charge-out non-order rel COS AB", "UA3211 - Charge-out non-order rel COS IN", _
                       "UA3221 - Charge-out non-order rel GA ", "UA3241 - Charge-out non-order rel Sales")
    
    Dim Nodes() As Variant
    Worksheets("Macro").Activate
    Nodes = Range("C3").CurrentRegion
    
    Dim r As Integer
    r = 2
    
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    
    ' Iterating through each category
    For Category = LBound(Categories) To UBound(Categories)
        
        Dim CategoryTipo As String
    
        Select Case Categories(Category)
            Case "UA3201 - Charge-out non-order rel COS AB"
                CategoryTipo = "CHARGE OUT COS AB"
            Case "UA3211 - Charge-out non-order rel COS IN"
                CategoryTipo = "CHARGE OUT COS IN"
            Case "UA3221 - Charge-out non-order rel GA "
                CategoryTipo = "CHARGE OUT G&A"
            Case "UA3241 - Charge-out non-order rel Sales"
                CategoryTipo = "CHARGE OUT SALES"
        End Select
    
        ' Iterating through each node
        For Node = LBound(Nodes) To UBound(Nodes)
            
            ' Expanding the appropriate node for division
            If Left(Nodes(Node, 1), 2) = "IA" Then
                Bu = Nodes(Node, 1)
                session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").expandNode Nodes(Node, 2)
            
            ' Processing for PG
            Else
                PG = Nodes(Node, 1)
                
                ' Filtering and showing rows from Chargeout file
                Workbooks("TABELA.xlsm").Worksheets("datacharge").Activate
                Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter
                Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter Field:=6, Operator:=xlFilterNoFill
                
                ' Filtering for the current PG and Category
                ActiveSheet.Range("A1:H" & lastRowUltimate).AutoFilter Field:=2, Criteria1:=Right(PG, 4)
                ActiveSheet.Range("A1:H" & lastRowUltimate).AutoFilter Field:=3, Criteria1:=CategoryTipo
                
                ' If no rows are visible, go to NextPG
                If Range("H1").End(xlDown).Row = Rows.Count Then
                    GoTo NextPG
                End If
                
                ' Handling account details in SAP
                session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").selectedNode = Nodes(Node, 2)
                session.findById("wnd[0]").sendVKey 71
                session.findById("wnd[1]/usr/txtRSYSF-STRING").Text = Categories(Category)
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                
                ' Error handling for missing values
                On Error GoTo TryDifferentLabel
                session.findById("wnd[2]/usr/lbl[17,2]").SetFocus
                GoTo AfterSelecting
                
TryDifferentLabel:
                On Error GoTo -1
                On Error GoTo Finish
                session.findById("wnd[2]/usr/lbl[16,2]").SetFocus
                
AfterSelecting:
                session.findById("wnd[2]").sendVKey 2
                session.findById("wnd[0]").sendVKey 2
                session.findById("wnd[1]/usr/lbl[1,2]").SetFocus
                session.findById("wnd[0]").sendVKey 2
                
                ' Resetting error handling
                On Error GoTo 0
                
                ' Adjusting layout and date filtering in SAP
                session.findById("wnd[0]").maximize
                session.findById("wnd[0]/tbar[1]/btn[33]").press
                session.findById("wnd[1]/tbar[0]/btn[71]").press
                session.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "Charge out new"
                session.findById("wnd[2]").sendVKey 0
                session.findById("wnd[3]/usr/lbl[15,2]").SetFocus
                session.findById("wnd[3]").sendVKey 2
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                
                ' Changing the date
                Dim currentDate As Date
                Dim firstDay As Date
                Dim lastDay As Date
                
                currentDate = Date
                firstDay = Format(DateSerial(Year(currentDate), Month(currentDate) - 1, 1), "dd.mm.yyyy")
                lastDay = Format(DateSerial(Year(currentDate), Month(currentDate), 0), "dd.mm.yyyy")
            
                session.findById("wnd[0]/tbar[1]/btn[38]").press
                session.findById("wnd[1]/usr/btnB_SEARCH").press
                session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "data di reg"
                session.findById("wnd[2]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/usr/btnAPP_WL_SING").press
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = firstDay
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = lastDay
                session.findById("wnd[1]").sendVKey 0
                session.findById("wnd[0]").maximize
                
                ' Going through rows in the table
                lastRow = Range("H1").End(xlDown).Row
                For r = 2 To lastRow
                    If Right(Range("B" & r), 4) = Right(PG, 4) And Range("C" & r) = CategoryTipo Then
                        session.findById("wnd[0]").sendVKey 71
                        session.findById("wnd[1]/usr/txtRSYSF-STRING").Text = AddDot(Range("F" & r))
                        session.findById("wnd[1]/tbar[0]/btn[0]").press
                        session.findById("wnd[2]/usr/lbl[8,2]").SetFocus
                        session.findById("wnd[2]").sendVKey 2
                        session.findById("wnd[0]/tbar[1]/btn[43]").press
                        session.findById("wnd[0]/tbar[1]/btn[9]").press
                        
                        ' Retrieving values from SAP
                        Range("D" & r) = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getcellvalue(0, "ZUONR")
                        Range("E" & r) = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getcellvalue(0, "KTONR")
                        
NextItem:
                        session.findById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                Next r
                
                ' Returning in SAP if no further action needed
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                
                GoTo NextPG
                
Finish:
                On Error GoTo -1
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/tbar[0]/btn[12]").press
                
NextPG:
            
            End If
        Next Node
    Next Category

    ' Resetting filters and showing the entire table
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter Field:=6, Operator:=xlFilterNoFill
    
    MsgBox "Done"
    
End Sub


