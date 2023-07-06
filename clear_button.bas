Attribute VB_Name = "Module2"
' Author: Isaac Finehout
' Date: 6 July 2023
' Title: Weekly Slides Updater CLEAR Button
' Purpose: Clear all contents from each slide
' Contact: isaac.finehout@us.af.mil
Sub Clear()
    ' Declare workbook variables
    ' ==========================
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TxtRng As range
    Dim count As Integer
    Dim user_input As String
    Set wb = ActiveWorkbook
    
    user_input = MsgBox("Are you sure you'd like to clear all document number input and document number output?", vbYesNo, "Clear Confirmation")
    
    If (user_input = 6) Then
        user_input = MsgBox("Click OK and wait for the next message box. Clearing contents takes approximately 2 seconds.", vbInformation, "Clear Accepted")
    
        ' Clear INPUT_TICMS_Requisitions
        ' ==============================
        Set ws = wb.Sheets("INPUT_TICMS_Requisitions")
        Set TxtRng = ws.range("A1:A1048576")
        TxtRng.ClearContents
        TxtRng.Interior.Color = RGB(117, 113, 113)
        For Each Border In TxtRng.Borders
            Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next Border
    
        ' Clear INPUT_SLIDES_Requisitions
        ' ===============================
        Set ws = wb.Sheets("INPUT_SLIDES_Requisitions")
        Set TxtRng = ws.range("A1:A1048576")
        TxtRng.ClearContents
        TxtRng.Interior.Color = RGB(117, 113, 113)
        For Each Border In TxtRng.Borders
            Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next Border
    
        ' Clear INPUT_TICMS_Outbounds
        ' ===========================
        Set ws = wb.Sheets("INPUT_TICMS_Outbounds")
        Set TxtRng = ws.range("A1:A1048576")
        TxtRng.ClearContents
        TxtRng.Interior.Color = RGB(117, 113, 113)
        For Each Border In TxtRng.Borders
            Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next Border
    
        ' Clear INPUT_SLIDES_Outbounds
        ' ============================
        Set ws = wb.Sheets("INPUT_SLIDES_Outbounds")
        Set TxtRng = ws.range("A1:A1048576")
        TxtRng.ClearContents
        TxtRng.Interior.Color = RGB(117, 113, 113)
        For Each Border In TxtRng.Borders
            Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next Border
    
        ' Clear OUTPUT
        ' ============
        Set ws = wb.Sheets("OUTPUT")
        Set TxtRng = ws.range("A4:F1048576")
        TxtRng.ClearContents
    
        user_input = MsgBox("All data has been cleared.", vbOKOnly, "Clear Acknowledgment")
    Else
        user_input = MsgBox("You chose not to clear data.", vbInformation, "Clear Declined")
    End If
    
    
End Sub
