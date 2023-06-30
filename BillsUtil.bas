Option Explicit

Attribute VB_Name = "BillsUtil"

Function PreviousSheet(Optional CellAddress As String) As Variant
'Returns the name of the previous worksheet
'If CellAddress is furnished, function returns a range reference to that range on the previous worksheet
Dim cel As Range
Dim s As String
Dim ws As Worksheet
Dim wb As Workbook
Application.Volatile
Set cel = Application.Caller
Set ws = cel.Worksheet
Set wb = ws.Parent
If ws.Index > 1 Then
    s = wb.Worksheets(ws.Index - 1).Name
    If InStr(1, s, "'") > 0 Then
        s = "'" & Replace(s, "'", "''") & "'"   'If a worksheet name contains a single quote, you must escape it by using two single quotes in formulas
    ElseIf InStr(1, s, " ") > 0 Then            'If a worksheet name contains a space, you must surround it with single quotes in formulas
        s = "'" & s & "'"
    End If
    If CellAddress <> "" Then
        s = s & "!" & CellAddress
        Set PreviousSheet = Range(s)
    Else
        PreviousSheet = s
    End If
End If
End Function

Sub ImportModule(moduleName As String)
    ' Need to change trust center developer security setting "Trust access
    ' to project object model" to True
    Dim wb As Workbook
    Dim basFile As String
    Dim vbComp As Object

    Set wb = ThisWorkbook
    basFile = wb.Path & "\" & moduleName & ".bas"

    ' Remove module if it exists
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Name = moduleName Then
            wb.VBProject.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp

    ' Add newest version of this module    
    wb.VBProject.VBComponents.Import basFile
End Sub