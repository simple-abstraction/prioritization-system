Attribute VB_Name = "Module2"
Option Explicit

Sub ApplyPriorityTemplate()

    Const TEMPLATE_TABLE As String = "Priorities_Template"
    Const CUSTOM_STYLE As String = "My Custom Table Style"
    Const FALLBACK_STYLE As String = "TableStyleMedium2"

    Dim ws As Worksheet
    Dim wsTarget As Worksheet
    Dim tblTemplate As ListObject
    Dim tblTarget As ListObject
    Dim templateDataRow As Range
    Dim targetDataRange As Range
    Dim i As Long

    On Error GoTo CleanFail

    'Find the template table anywhere in the workbook
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tblTemplate = ws.ListObjects(TEMPLATE_TABLE)
        On Error GoTo CleanFail

        If Not tblTemplate Is Nothing Then Exit For
    Next ws

    If tblTemplate Is Nothing Then
        MsgBox "Could not find the table named '" & _
               TEMPLATE_TABLE & "'.", vbCritical
        Exit Sub
    End If

    'Use the currently open worksheet as the destination
    Set wsTarget = ActiveSheet

    If wsTarget.Name = tblTemplate.Parent.Name Then
        MsgBox "Open a priority sheet, then run the macro.", vbExclamation
        Exit Sub
    End If

    'Find the Power Query output table on the active sheet
    If wsTarget.ListObjects.Count = 0 Then
        MsgBox "No table was found on the active sheet.", vbExclamation
        Exit Sub
    End If

    Set tblTarget = wsTarget.ListObjects(1)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Apply custom table style, or fall back to Light 1
    If TableStyleExists(CUSTOM_STYLE) Then
        tblTarget.TableStyle = CUSTOM_STYLE
    Else
        tblTarget.TableStyle = FALLBACK_STYLE
    End If

    'Copy column widths
    For i = 1 To Application.Min( _
        tblTemplate.Range.Columns.Count, _
        tblTarget.Range.Columns.Count)

        tblTarget.Range.Columns(i).ColumnWidth = _
            tblTemplate.Range.Columns(i).ColumnWidth

    Next i

    'Copy header formatting
    tblTemplate.HeaderRowRange.Copy
    tblTarget.HeaderRowRange.PasteSpecial Paste:=xlPasteFormats

    'Copy body formatting, conditional formatting, and validation
    If Not tblTemplate.DataBodyRange Is Nothing _
       And Not tblTarget.DataBodyRange Is Nothing Then

        Set templateDataRow = tblTemplate.DataBodyRange.Rows(1)
        Set targetDataRange = tblTarget.DataBodyRange

        templateDataRow.Copy
        targetDataRange.PasteSpecial Paste:=xlPasteFormats

        templateDataRow.Copy
        targetDataRange.PasteSpecial Paste:=xlPasteValidation

        targetDataRange.RowHeight = templateDataRow.RowHeight

    End If

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Template applied to '" & tblTarget.Name & "'.", vbInformation
    Exit Sub

CleanFail:

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The template could not be applied:" & vbCrLf & _
           Err.Description, vbCritical

End Sub


Private Function TableStyleExists(ByVal StyleName As String) As Boolean

    Dim ts As TableStyle

    On Error Resume Next
    Set ts = ThisWorkbook.TableStyles(StyleName)
    TableStyleExists = Not ts Is Nothing
    On Error GoTo 0

End Function

