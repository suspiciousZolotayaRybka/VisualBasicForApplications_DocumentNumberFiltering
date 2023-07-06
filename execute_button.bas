Attribute VB_Name = "Module1"
' Author: Isaac Finehout
' Date: 5 July 2023
' Title: Weekly Slides Updater EXECUTE Button
' Purpose: Compare documents numbers from TICMS and the Weekly Slides to determine how to update the Weekly Slides
' Contact: isaac.finehout@us.af.mil

' This is the main function
'==========================
Sub Execute()
    ' Declare the arrays of comparable document numbers
    Dim req_ticms() As String
    Dim out_ticms() As String
    Dim req_slides() As String
    Dim out_slides() As String
    ' Get the doc numbers from user input
    ' GetDocNumbers checks for arrays size 0 and exits the program if this is true
    req_ticms = GetDocNumbers("INPUT_TICMS_Requisitions")
    out_ticms = GetDocNumbers("INPUT_TICMS_Outbounds")
    req_slides = GetDocNumbers("INPUT_SLIDES_Requisitions")
    out_slides = GetDocNumbers("INPUT_SLIDES_Outbounds")
    
    
    ' Find the various useful comparisons between doc numbers for output
    Dim repeat_req_slides() As Boolean
    Dim repeat_out_slides() As Boolean
    repeat_req_slides = GetRepeatDocNumbers(req_slides)
    repeat_out_slides = GetRepeatDocNumbers(out_slides)
    
    Dim new_req_doc_numbers() As Boolean
    Dim new_out_doc_numbers() As Boolean
    new_req_doc_numbers = GetNewDocNumbers(req_ticms, req_slides)
    new_out_doc_numbers = GetNewDocNumbers(out_ticms, out_slides)
    
    Dim old_req_doc_numbers() As Boolean
    Dim old_out_doc_numbers() As Boolean
    old_req_doc_numbers = GetNewDocNumbers(req_slides, req_ticms)
    old_out_doc_numbers = GetNewDocNumbers(out_slides, out_ticms)
    
    
    ' Output the data
    placeholder = OutputData(req_ticms, out_ticms, req_slides, out_slides, repeat_req_slides, repeat_out_slides, new_req_doc_numbers, new_out_doc_numbers, old_req_doc_numbers, old_out_doc_numbers)
    
    
End Sub

' Output all the data onto the output slide
'==========================================
Private Function OutputData(req_ticms, out_ticms, req_slides, out_slides, repeat_req_slides, repeat_out_slides, new_req_doc_numbers, new_out_doc_numbers, old_req_doc_numbers, old_out_doc_numbers) As String
    ' Declare workbook variables
    ' ==========================
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TxtRng As range
    Dim count As Integer
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("OUTPUT")
    
    ' Output repeating document numbers for SLIDES Requistions and Outbounds
    ' ======================================================================
    ' Requisitions
    Set TxtRng = ws.range("A4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(repeat_req_slides) + 1))
        If repeat_req_slides(count) Then
            TxtRng.Value = (req_slides(count) + "_REPEAT")
        Else
            TxtRng.Value = (" " + req_slides(count))
        End If
        Set TxtRng = ws.range("A" + CStr(count + 5))
        count = count + 1
    Loop
    ' Outbounds
    Set TxtRng = ws.range("B4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(repeat_out_slides) + 1))
        If repeat_out_slides(count) Then
            TxtRng.Value = (out_slides(count) + "_REPEAT")
        Else
            TxtRng.Value = (" " + out_slides(count))
        End If
        Set TxtRng = ws.range("B" + CStr(count + 5))
        count = count + 1
    Loop
    
    ' Output new document numbers
    ' ===========================
    ' Requisitions
    Set TxtRng = ws.range("C4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(new_req_doc_numbers) + 1))
        If new_req_doc_numbers(count) Then
            TxtRng.Value = (req_ticms(count) + "_NEW")
        Else
            TxtRng.Value = (" " + req_ticms(count))
        End If
        Set TxtRng = ws.range("C" + CStr(count + 5))
        count = count + 1
    Loop
    ' Outbounds
    Set TxtRng = ws.range("D4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(new_out_doc_numbers) + 1))
        If new_out_doc_numbers(count) Then
            TxtRng.Value = (out_ticms(count) + "_NEW")
        Else
            TxtRng.Value = (" " + out_ticms(count))
        End If
        Set TxtRng = ws.range("D" + CStr(count + 5))
        count = count + 1
    Loop
    
    
    ' Output old document numbers
    ' ===========================
    ' Requisitions
    Set TxtRng = ws.range("E4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(old_req_doc_numbers) + 1))
        If old_req_doc_numbers(count) Then
            TxtRng.Value = (req_slides(count) + "_OLD")
        Else
            TxtRng.Value = (" " + req_slides(count))
        End If
        Set TxtRng = ws.range("E" + CStr(count + 5))
        count = count + 1
    Loop
    
    ' Outbounds
    Set TxtRng = ws.range("F4")
    count = 0
    TxtRng.Value = "NULL"
    Do While (count < (UBound(old_out_doc_numbers) + 1))
        If old_out_doc_numbers(count) Then
            TxtRng.Value = (out_slides(count) + "_OLD")
        Else
            TxtRng.Value = (" " + out_slides(count))
        End If
        Set TxtRng = ws.range("F" + CStr(count + 5))
        count = count + 1
    Loop
    
    OutputData = "NULL; you should never see this string output"
End Function

' Finds doc numbers that are repeated more than once in the slides
'=================================================================
Private Function GetRepeatDocNumbers(doc_numbers) As Boolean()
    ' Assign boolean array to test for repeat values
    Dim repeat_doc_numbers() As Boolean
    
    ' Find the length of the boolean array based on doc_numbers upper bound
    Dim len_repeat_doc_numbers As Integer
    len_repeat_doc_numbers = UBound(doc_numbers)
    ReDim repeat_doc_numbers(len_repeat_doc_numbers)

    ' Use a nested for loop to find repeat doc numbers
    ' Awful compute time, but doc numbers should never be above 100 (100*100 = 10,000 max iterations)
    Dim i_count As Integer
    Dim j_count As Integer
    i_count = 0
    
    For Each i_doc_number In doc_numbers
        j_count = 0
        For Each j_doc_number In doc_numbers
            If ((i_doc_number = j_doc_number) And (Not (j_count = i_count))) Then
                ' If the doc numbers are equal, use the j_count index to set equal to True
                repeat_doc_numbers(j_count) = True
            End If
            j_count = j_count + 1
        Next j_doc_number
        i_count = i_count + 1
    Next i_doc_number
    GetRepeatDocNumbers = repeat_doc_numbers
End Function

' Return an array of document numbers that are in TICMS, but not yet on the slides
'=============================================================================
Private Function GetNewDocNumbers(ticms_doc_numbers, slides_doc_numbers) As Boolean()
    ' Assign boolean array to test for new values
    Dim new_doc_numbers() As Boolean
    
    ' Find the length of the boolean array
    Dim len_new_doc_numbers As Integer
    Dim count As Integer
    len_new_doc_numbers = UBound(ticms_doc_numbers)
    ReDim new_doc_numbers(len_new_doc_numbers)
    count = 0
    ' If doc_number is in TICMS and NOT in slides it is new
    For Each doc_number In ticms_doc_numbers
        If (Not (IsIn(doc_number, slides_doc_numbers))) Then new_doc_numbers(count) = True
        count = count + 1
    Next doc_number
    GetNewDocNumbers = new_doc_numbers
End Function

' Return an array of document numbers that are on the slides, but no longer in TICMS
'===================================================================================
Private Function GetOldDocNumbers(slides_doc_numbers, ticms_doc_numbers) As Boolean()
    ' Assign boolean array to test for new values
    Dim old_doc_numbers() As Boolean
    
    ' Find the length of the boolean array
    Dim len_old_doc_numbers As Integer
    Dim count As Integer
    len_old_doc_numbers = UBound(slides_doc_numbers)
    ReDim old_doc_numbers(len_old_doc_numbers)
    count = 0
    ' If doc_number is in the slides and NOT in TICMS it is old
    For Each doc_number In slides_doc_numbers
        If (Not (IsIn(doc_number, ticms_doc_numbers))) Then old_doc_numbers(count) = True
        count = count + 1
    Next doc_number
    GetOldDocNumbers = old_doc_numbers
End Function

' Use the A column from the sheet_name passed as an argument to input document numbers from an excel sheet
'=========================================================================================================
Private Function GetDocNumbers(sheet_name) As String()
    ' Assign the number of documents numbers to declare array doc_nums
    '=================================================================
    Dim num_doc_numbers As Integer
    num_doc_numbers = 0
    ' Assign cell_name to control loop while data exists in cell
    Dim cell_name As String
    Dim cell_range As String
    Dim cell_count As Integer
    cell_count = 1
    cell_range = "A" + CStr(cell_count)
    cell_name = ThisWorkbook.Sheets(sheet_name).range(cell_range)
    
    ' Iterate through loop while data exists in each cell to get the number of doc numbers
    '=====================================================================================
    Do While (Not (cell_name = ""))
        num_doc_numbers = num_doc_numbers + 1
        cell_count = cell_count + 1
        cell_range = "A" + CStr(cell_count)
        cell_name = ThisWorkbook.Sheets(sheet_name).range(cell_range)
    Loop
    
    If (num_doc_numbers = 0) Then
        Dim placeholder As String
        placeholder = MsgBox("All or some 'INPUT' sheets do not have doc numbers input. Please ensure all 'INPUT' sheets have input.", vbExclamation, "Empty Input Warning")
        End
    End If
    ' get the doc_numbers array filled with a for loop
    '=================================================
    ' Re-state variable to control cell input
    cell_name = ""
    ' Declare the doc_numbers array
    Dim doc_numbers() As String
    ReDim doc_numbers(num_doc_numbers - 1)
    Dim i As Integer
    ' Iterate for each doc number index and fill in the array
    For i = 0 To (num_doc_numbers - 1)
        cell_name = ThisWorkbook.Sheets(sheet_name).range("A" + CStr(i + 1))
        ' Remove spaces TODO
        cell_name = Replace(cell_name, " ", "")
        doc_numbers(i) = cell_name
    Next i
    GetDocNumbers = doc_numbers
End Function

' Finds if an element is inside of an array, returns boolean
'=================================================================
Private Function IsIn(search, grouping) As Boolean
    Dim is_in As Boolean
    For Each element In grouping
        If (element = search) Then is_in = True
    Next element
    IsIn = is_in
End Function

' This function does absolutely nothing I'm just too afraid to delete the debug print statements
' VBA debugging seems counter intuitive to me so I decided to just use lots of print statements instead
' Make fun of me all you want
Private Function DebuggingComments()
    'Debug.Print ("=============$=========================")
    'Debug.Print ("LowerBound" + CStr(LBound(old_out_doc_numbers)) + "     " + "UpperBound" + CStr(UBound(old_out_doc_numbers)))
    'Dim count As Integer
    'count = 0
    'For Each boolean_doc_number In old_out_doc_numbers
        'Debug.Print (CStr(count) + ": " + CStr(boolean_doc_number))
        'count = count + 1
    'Next boolean_doc_number
    'Debug.Print ("==============$========================")
    
    
    'Debug.Print ("Count" + CStr(count) + CStr(old_out_doc_numbers(count)) + out_slides(count))
    'Debug.Print ("Count" + CStr(count) + CStr(old_out_doc_numbers(count)) + out_slides(count))
    'Debug.Print ("============^================")
    'Debug.Print ("===============^=============")
End Function
