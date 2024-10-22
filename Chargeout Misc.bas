Attribute VB_Name = "Misc"
Sub CleaningTable()

    ' Clears filter in the datacharge table, deletes everything,
    ' copies the table from "datacharge copy", pastes it back into "datacharge"
    ' and filters to hide cells with yellow color in column F (which are usually empty).
    ' Rows with yellow cells may have data someday, so we keep them.
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter
    Worksheets("datacharge").Range("A2:H" & Range("A2").End(xlDown).Row).EntireRow.Delete
    Worksheets("datacharge copy").Activate
    Worksheets("datacharge copy").Range("M2:T" & Range("M2").End(xlDown).Row).Copy _
    Destination:=Worksheets("datacharge").Range("A2")
    
    ' Filtering rows in column F with no fill (to hide the yellow cells)
    Worksheets("datacharge").Range("A1").CurrentRegion.AutoFilter Field:=6, Operator:=xlFilterNoFill
    MsgBox "Done"

End Sub

Function AddDot(Number As String) As String

    ' This function adds dots to a number string to format it correctly for SAP.
    
    ' Find the position of "," in the number
    Position = InStr(Number, ",")
    
    ' Adding 0 to the end if there is only one digit after the comma (e.g., 123,4 -> 123,40)
    If Len(Number) - Position = 1 Then Number = Number & "0"
    
    ' Adding ",00" if there's no comma (e.g., 123 -> 123,00)
    If Position = 0 Then Number = Number & ",00"
    
    ' For negative numbers, move the minus sign to the end of the string for SAP compatibility
    If Left(Number, 1) = "-" Then
        Number = Right(Number, Len(Number) - 1) & "-"
        n = Len(Number)
        
        ' Add a dot for numbers larger than 1000 (e.g., 52000,00- -> 52.000,00-)
        If n > 7 Then
            AddDot = Left(Number, n - 7) & "." & Right(Number, 7)
        Else
            AddDot = Number
        End If
    Else
        ' For positive numbers, same logic without the minus sign
        n = Len(Number)
        If n > 6 Then
            AddDot = Left(Number, n - 6) & "." & Right(Number, 6)
        Else
            AddDot = Number
        End If
    End If

    ' Handling numbers larger than 1 million manually if needed
    ' (e.g., 1000.000,00 should be 1.000.000,00)
    
End Function

Sub Insert_Line_In_Main_File()

    ' This subroutine inserts new lines in the main file based on specific PG and category filters
    
    Dim period As String
    period = Format(DateAdd("m", -1, Date), "yymm") ' Define the period as the previous month (YYMM format)
    
    Dim Nodes() As Variant
    Worksheets("Macro").Activate
    Nodes = Range("C3:C" & Range("C3").End(xlDown).Row) ' Load nodes from the "Macro" sheet
    
    ' Define categories of Chargeouts
    Dim Categories() As Variant
    Categories = Array("CHARGE OUT COS AB", "CHARGE OUT COS IN", "CHARGE OUT G&A", "CHARGE OUT SALES")
    
    ' Loop through each node and category
    For Each Node In Nodes
        
        If Left(Node, 2) = "PG" Then ' If the node starts with "PG", it's a valid entry for the process
            PG = Right(Node, 4)
            
            For Each Category In Categories
                
                ' Activate Tabela workbook and filter datacharge sheet based on PG and category
                Windows("Tabela.xlsm").Activate
                Sheets("datacharge").Activate
                Sheets("datacharge").Range("A1:H" & Range("H1").End(xlDown).Row).AutoFilter Field:=2, Criteria1:=PG
                Sheets("datacharge").Range("A1:H" & Range("H1").End(xlDown).Row).AutoFilter Field:=3, Criteria1:=Category
                
                ' If filtered rows exist, proceed
                If Range("A1").End(xlDown).Row <> Rows.Count Then
                
                    ' Count the number of visible rows after filtering
                    Dim numberOfRows As Integer
                    numberOfRows = Range("A1:A" & Range("A1").End(xlDown).Row).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                    
                    ' Activate the IA file and filter based on PG and category
                    Windows("Charge out Retrieve IA " & period & ".xlsx").Activate
                    Sheets("INPUT_SAP_IA").Range("A1:V" & Range("V1").End(xlDown).Row).CurrentRegion.AutoFilter Field:=2, Criteria1:=PG
                    Sheets("INPUT_SAP_IA").Range("A1:V" & Range("V1").End(xlDown).Row).CurrentRegion.AutoFilter Field:=13, Criteria1:=Category
                    
                    ' If the filtered rows exist, insert new rows as needed
                    If Range("A1").End(xlDown).Row <> Rows.Count Then
                        Dim lastRow As Integer
                        lastRow = Range("A1").End(xlDown).Row
                        
                        ' Insert new rows based on the number of rows in the datacharge sheet
                        Dim i As Integer
                        i = 1
                        Do While i <= numberOfRows
                            Rows(lastRow + 1).Insert Shift:=xlDown
                            i = i + 1
                        Loop
                    End If
                End If
            Next Category
        End If
    Next Node

    ' Reapply the filter and finish up
    Windows("Charge out Retrieve IA " & period & ".xlsx").Activate
    Range("A1").CurrentRegion.AutoFilter
    Windows("TABELA.xlsm").Activate
    Sheets("datacharge").Range("A1").CurrentRegion.AutoFilter
    Sheets("datacharge").Range("A1").CurrentRegion.AutoFilter Field:=6, Operator:=xlFilterNoFill
    
    MsgBox "Done"
    
End Sub


