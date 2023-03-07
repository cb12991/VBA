Attribute VB_Name = "AVID"
Sub fileCheck()
'Desc:          This quickly checks if files are in import directory for AVID import.
'Author:        Cody B. Buehler
'Created:       7/28/2020
'Updated:       2/22/2021
    
    With ActiveSheet
        
        If IsEmpty(.Range("inpYr")) Or IsEmpty(.Range("inpQtr")) Then
            MsgBox "You need to update the year and quarter input cells for this input control file before activating macro."
            Exit Sub
        End If
        
        inpCol = .Range("inpCol").Column
        ultRow = .Cells(3, 6).End(xlDown).Row
        inpYr = .Range("inpYr")
        inpQtr = Right(.Range("inpQtr"), 1)
        
        If inpQtr = 4 Then mon = "12" Else mon = "0" & (inpQtr * 3)
        
        For i = 3 To ultRow
            
            If .Cells(i, 6) = "Input Control.*" Then
                file = .Cells(i, 6)
                Path = "P:\Avid\Stage\Import\" & inpYr & "-" & mon & "\" & file
            Else
                file = Mid(.Cells(i, 6), 7)
                Path = "P:\Avid\Stage\Import\" & inpYr & "-" & mon & "\" & inpYr & mon & file
            End If
            
            exists = Dir(Path)
            
            If exists = "" Then
                    .Cells(i, inpCol) = "Missing"
            Else
                .Cells(i, inpCol) = "Included"
            End If
        Next i
    End With
End Sub

Sub fileCheck2()
'Desc:      This quickly checks for files that are in import directory and not listed in worksheet.
'Author:    Cody B. Buehler
'Date:      12/11/2020
'Updated:   2/22/2021
    
    With ActiveSheet
        
        If IsEmpty(.Range("inpYr")) Or IsEmpty(.Range("inpQtr")) Then
            MsgBox "You need to update the year and quarter input cells for this input control file before activating macro."
            Exit Sub
        End If
        
        ultRow = .Cells(3, 6).End(xlDown).Row
        inpYr = .Range("inpYr")
        inpQtr = Right(.Range("inpQtr"), 1)
        
        If inpQtr = 4 Then mon = "12" Else mon = "0" & (inpQtr * 3)

        Set rng = .Range(.Cells(3, 6), .Cells(ultRow, 6))
        i = 3
        file = Dir("P:\Avid\Stage\Import\" & inpYr & "-" & mon & "\")
        Do While file <> ""
            Set c = rng.Find(Mid(file, 7, Len(file) - 10), , , xlPart)
            If c Is Nothing Then
                .Cells(i, 16) = file
                i = i + 1
            End If
            file = Dir()
        Loop
    End With
End Sub

Public Sub adjContractNumbers()
' BY:                Cody B. Buehler
' LAST UPDATED:      09/22/2021
'
' DESCRIPTION:       This program standardizes contract numbers, formats date fields, and
' removes blank rows from various import file structures.
'
    ' Macro optimization.
    With Application
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    On Error Resume Next
    
    ' Declare arrays.
    Dim structure(1 To 11) As String: Dim unadj_files() As String: Dim adj_files() As String
    Dim bad_hdr_address() As String: Dim bad_hdr() As String: Dim bad_hdr_file() As String
    Dim correct_hdr() As String: Dim bad_hdr_row() As Integer: Dim bad_hdr_col() As Integer
    Dim unadj_reason() As String
    
    ' Initialize variables.
    protError = False: hdrError = False: dtError = False: save_changes = True
    bad_yr = False: bad_mm = False: count_tot = 0: count_err = 0: count_bad_hdr = 0:
    
    ' Set original workbook.
    Set wb_home = ActiveWorkbook
    
    ' Add structures to array.
    structure(1) = "PData"
    structure(2) = "RData"
    structure(3) = "Trad VData"
    structure(4) = "UL VData"
    structure(5) = "AssumedTrad VData"
    structure(6) = "AssumedUL VData"
    structure(7) = "ReinsYRT VData"
    structure(8) = "IndVAnn VData"
    structure(9) = "GrpVAnn VData"
    structure(10) = "NonVAnn VData"
    structure(11) = "Payout VData"
        
    ' Set column name reference table.
    Set col_ref = wb_home.Sheets("AVID Structure Mapping").ListObjects("AVID_Structure_Mapping")

    ' Get valuation date & import directory path.
    valDate = Excel.Application.WorksheetFunction.EoMonth( _
                DateSerial(ActiveSheet.Range("inpYr").Value, _
                    Right(ActiveSheet.Range("inpQtr").Value, 1) * 3, 1), 0)
    import_path = "P:\Avid\Stage\Import\" & Year(valDate) & "-" & Right("0" & Month(valDate), 2) & "\Curating\"
    
    ' Confirm import directory path.
    usr_input = _
        InputBox( _
            Prompt:="Confirm or update input file path.", _
            Title:="Adjust " & Right(Year(valDate), 2) & ActiveSheet.Range("inpQtr").Value & " Inputs", _
            Default:=import_path)
    If usr_input = "" Then Exit Sub Else import_path = usr_input
           
    If Right(import_path, 1) <> "\" Then import_path = import_path & "\"
           
    ' Initialize path for new folder to store formatted files.
    formatted_path = import_path & "Formatted\"
    
    ' If Formatted folder doesn't exist, make a new directory.
    If Dir(formatted_path) = "" Then MkDir formatted_path
        
    ' Initialize file counter for all structures.
    no_files = 0
           
    ' See if user wants interactive prompts displayed throughout loop.
    If MsgBox("Execute program quietly?", vbYesNo) = vbYes Then
        quietly = True
    Else
        quietly = False
    End If
           
    ' Outer for loop w.r.t. structure().
    For j = LBound(structure) To UBound(structure)
        ' Initialze counters with respect to structure j.
        count_err_j = 0
        count_j = 0
        no_files_j = 0
        
        ' Get file name (retain space with wildcard to prevent partial matches).
        file_j = Dir(import_path & "* " & structure(j) & "*")
        
        ' Inner while loop w.r.t. files of structure(j).
        While file_j <> ""
            ' Increment counters.
            no_files = no_files + 1
            no_files_j = no_files_j + 1
            
            ' Open input file and screen for errors.
            Set wb = Workbooks.Open(Filename:=import_path & file_j, Format:=2) 'format code 2 is comma delimiter
            
            With wb.Worksheets(1)
            
                ' We will manually change types to correct data type, so set all to General to
                ' avoid ambiguities.
                .Cells.NumberFormat = "General"
                
                If structure(j) = "PData" Or structure(j) = "RData" Then
                    hdrEnd = 1
                    firstRow = 2
                Else
                    hdrEnd = 3
                    firstRow = 5
                End If
                
                lastCol = .Cells(hdrEnd, .Columns.Count).End(xlToLeft).Column
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                
                ' Set headings range.
                Set colNames = .Range(.Cells(1, 1), .Cells(hdrEnd, lastCol))
                
                ' Set data range.
                Set data_rng = .Range(.Cells(firstRow, 1), .Cells(lastRow, lastCol))
                
                ' Check column headings.
                For Each cn In colNames
                    If Not Trim(cn.Value) = "" Then
                        ' Reset filter in case previously filtered.
                        col_ref.AutoFilter.ShowAllData
                        
                        ' Filter column reference table for current structure j, row, and column number.
                        col_ref.HeaderRowRange.AutoFilter _
                            field:=1, _
                            Criteria1:=structure(j)
                        col_ref.HeaderRowRange.AutoFilter _
                            field:=2, _
                            Criteria1:=cn.Row
                        col_ref.HeaderRowRange.AutoFilter _
                            field:=3, _
                            Criteria1:=cn.Column
                        
                        correct_col = col_ref.DataBodyRange.SpecialCells(xlCellTypeVisible).Columns(4).Value
                        
                        ' Now compare filtered reference with looped cell.
                        If Trim(cn.Value) <> correct_col Then
                            If Not quietly Then
                                inp = MsgBox(wb.Name & " has header value """ & Trim(cn.Value) & _
                                          """ instead of expected """ & correct_col & _
                                          """. Update header to match expected?", vbYesNo)
                            Else
                                inp = vbNo
                            End If
                                
                            If inp = vbYes Then
                                cn.Value = correct_col
                            Else
                                hdrError = True
                                count_bad_hdr = count_bad_hdr + 1
                        
                                ReDim Preserve bad_hdr_address(1 To count_bad_hdr)
                                bad_hdr_address(count_bad_hdr) = cn.Address
                                
                                ReDim Preserve bad_hdr_col(1 To count_bad_hdr)
                                bad_hdr_col(count_bad_hdr) = cn.Column
                                
                                ReDim Preserve bad_hdr_row(1 To count_bad_hdr)
                                bad_hdr_row(count_bad_hdr) = cn.Row
                                
                                ReDim Preserve bad_hdr(1 To count_bad_hdr)
                                bad_hdr(count_bad_hdr) = Trim(cn.Value)
                                
                                ReDim Preserve bad_hdr_file(1 To count_bad_hdr)
                                bad_hdr_file(count_bad_hdr) = file_j
                                
                                ReDim Preserve correct_hdr(1 To count_bad_hdr)
                                correct_hdr(count_bad_hdr) = correct_col
                            End If
                        End If
                    End If
                Next cn
                  
                ' Check if sheet is protected.
                If .ProtectContents = True Then protError = True
                
                ' Specify structure-specific column locations.
                Select Case structure(j)
                
                    Case "PData"
                        ReDim dt_cols(-1)
                        cntr_no_col = 3
                        as_of_mm_col = 4
                        as_of_yr_col = 5
                                                
                    Case "RData"
                        ReDim dt_cols(-1)
                        cntr_no_col = 1
                        as_of_mm_col = 9
                        as_of_yr_col = 10
                     
                    Case "Trad VData", "AssumedTrad VData"
                        ReDim dt_cols(1 To 4)
                        dt_cols(1) = 1
                        dt_cols(2) = 4
                        dt_cols(3) = 16
                        dt_cols(4) = 23
                        cntr_no_col = 3
                        val_dt_col = 1
                                            
                    Case "UL VData", "AssumedUL VData"
                        ReDim dt_cols(1 To 4)
                        dt_cols(1) = 1
                        dt_cols(2) = 12
                        dt_cols(3) = 27
                        dt_cols(4) = 33
                        cntr_no_col = 3
                        val_dt_col = 1
                        
                    Case "ReinsYRT VData"
                        ReDim dt_cols(1 To 2)
                        dt_cols(1) = 3
                        dt_cols(2) = 13
                        cntr_no_col = 2
                        val_dt_col = 0
                        
                    Case "GrpVAnn VData", "NonVAnn VData", "IndVAnn VData"
                        ReDim dt_cols(1 To 3)
                        dt_cols(1) = 1
                        dt_cols(2) = 7
                        dt_cols(3) = 15
                        cntr_no_col = 2
                        val_dt_col = 1
                        
                    Case "Payout VData"
                        ReDim dt_cols(1 To 6)
                        dt_cols(1) = 1
                        dt_cols(2) = 6
                        dt_cols(3) = 7
                        dt_cols(4) = 8
                        dt_cols(5) = 9
                        dt_cols(6) = 30
                        cntr_no_col = 3
                        val_dt_col = 1
                        
                    Case Else
                        MsgBox "Structure not found?"
                        Debug.Assert True
                        
                End Select
                    
                ' Check valuation date values (PData and RData have as-of month and year whereas
                ' VData files have an actual date).
                If structure(j) = "PData" Or structure(j) = "RData" Then
                    ' Get all distinct years in column.
                    distinct_yrs = WorksheetFunction.Unique(data_rng.Columns(as_of_yr_col))
                    
                    ' Check as-of year.
                    yrs = vbNewLine
                    For Each yr In distinct_yrs
                        If yr <> Year(valDate) Then
                            bad_yr = True
                            yrs = yrs & yr & vbNewLine
                        End If
                    Next yr
                                        
                    If bad_yr And Not quietly Then
                        If MsgBox(wb.Name & " has the following values for AS-OF-YR:" & _
                            vbNewLine & yrs & vbNewLine & _
                            "Change all to valuation year = " & Year(valDate) & "?", _
                            vbYesNo) = vbYes Then
                            data_rng.Columns(as_of_yr_col) = Year(valDate)
                            bad_yr = False
                        End If
                    End If
                    
                    ' Get all distinct months in column.
                    distinct_mm = WorksheetFunction.Unique(data_rng.Columns(as_of_mm_col))
                    
                    ' Check as-of month.
                    mmths = vbNewLine
                    For Each mm In distinct_mm
                        If mm <> Month(valDate) Then
                            bad_mm = True
                            mmths = mmths & mm & vbNewLine
                        End If
                    Next mm
                    
                    If bad_mm And Not quietly Then
                        If MsgBox(wb.Name & " has the following values for AS-OF-MM:" & _
                            vbNewLine & mmths & vbNewLine & _
                            "Change all to valuation month = " & Month(valDate) & "?", _
                            vbYesNo) = vbYes Then
                            data_rng.Columns(as_of_mm_col) = Month(valDate)
                            bad_mm = False
                        End If
                    End If
                    
                    If bad_mm Or bad_yr Then dtError = True
                    
                ElseIf val_dt_col > 0 Then
                    ' Get all distinct valuation dates in column.
                    distinct_dts = WorksheetFunction.Unique(data_rng.Columns(val_dt_col))
                    
                    ' Check valuation date.
                    dts = vbNewLine
                    For Each dt In distinct_dts
                        If dt <> valDate Then
                            dtError = True
                            dts = dts & Format(dt, "mm/dd/yyyy") & vbNewLine
                        End If
                    Next dt
                    
                    If dtError And Not quietly Then
                        If MsgBox(wb.Name & " has the following values for VALUATION DATE:" & _
                            vbNewLine & dts & vbNewLine & _
                            "Change all to current valuation date " & _
                            Format(valDate, "mm/dd/yyyy") & "?", vbYesNo) = vbYes Then
                            
                            data_rng.Columns(val_dt_col) = valDate
                            dtError = False
                        End If
                    End If
                End If
                    
                ' Format contract numbers
              For Each c In Range(data_rng.Columns(cntr_no_col).Address)
                  c.Value = "'" & WorksheetFunction.Text(c, "0000000000")
              Next c
              
              'Reformat date fields (VData only).
              If Not (structure(j) = "PData" Or structure(j) = "RData") Then
                  ' Sometimes an invalid date will be listed (e.g., 09/31/2022). These will be
                  ' stored as text since Excel can't convert to a date serial number. For now,
                  ' the fix will be to change to the first day of the next month. Other fixes
                  ' can be incorporated should they arise.
                  For i = 1 To UBound(dt_cols)
                      With data_rng.Columns(dt_cols(i))
                           For Each c In Range(.Address)
                                If Not (IsNumeric(c.Value) Or Trim(c.Value) = "") Then
                                     new_dt = DateSerial(Right(c.Text, 4), Left(c.Text, 2), Mid(c.Text, 4, 2))
                                     dtError = True
                                     If Not quietly Then
                                          If MsgBox("Change Date cell " & c.Text & " to " & new_dt & "?", _
                                                    vbYesNo) = vbYes Then
                                               c.Value = new_dt
                                               dtError = False
                                          End If
                                     End If
                                End If
                           Next c
                           .NumberFormat = "mm/dd/yyyy"
                      End With
                  Next i
              End If
              
              ' Delete trailing blank rows.
              .Range(.Cells(lastRow + 1, 1), .Cells(.Rows.Count, _
                  .Columns.Count)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
              
              ' Replace blanks in data (that will be imported) with zeros.
              data_rng.Replace _
                  What:="", _
                  Replacement:=0, _
                  LookAt:=xlWhole
          
              ' Remove any double pound signs. These are most likely escape characters
              ' and have caused import issues in the past.
              data_rng.Replace _
                  What:="##", _
                  Replacement:="", _
                  LookAt:=xlPart
              
              ' Dates will be changed to 1/0/1900 or 01/00/1900 which will cause AVID to
              ' error upon import. Change to a single space to avoid.
              data_rng.Replace _
                  What:="1/0/1900", _
                  Replacement:=" ", _
                  LookAt:=xlWhole
              data_rng.Replace _
                  What:="01/00/1900", _
                  Replacement:=" ", _
                  LookAt:=xlWhole
        
              'Ran into truncation errors for fields with length > 50, so truncate to
              'avoid future errors.
              For Each c In data_rng
                  If Len(c.Value) > 50 And IsNumeric(c.Value) Then
                      c.Value = CDbl(Left(c.Value, 50))
                  ElseIf Len(c.Value) > 50 Then
                      c.Value = Left(c.Value, 50)
                  End If
              Next c
              
              ' Some ReinsYRT imports have date columns in YYYYMMDD format instead of date.
              ' These are set to 0 from previous formatting steps; clear contents instead to
              ' avoid importing errors.
              If structure(j) = "ReinsYRT VData" Then
                  For i = 22 To 23
                      data_rng.Columns(i).Replace _
                          What:=0, _
                          Replacement:=" ", _
                          LookAt:=xlWhole
                  Next i
              End If
          
            End With
            
            ' If header errors, build array for summary and flag to not save.
            If hdrError Then
                count_err_j = count_err_j + 1
                count_err = count_err + 1
                
                ReDim Preserve unadj_files(1 To count_err)
                unadj_files(count_err) = wb.Name
                
                ReDim Preserve unadj_reason(1 To count_err)
                unadj_reason(count_err) = "DIFFERENT COLUMN HEADINGS"
                
                save_changes = False
            
            ' If date errors, build array for summary and flag to not save.
            ElseIf dtError Then
                count_err_j = count_err_j + 1
                count_err = count_err + 1
                
                ReDim Preserve unadj_files(1 To count_err)
                unadj_files(count_err) = wb.Name
                
                ReDim Preserve unadj_reason(1 To count_err)
                unadj_reason(count_err) = "BAD DATE VALUE"
                
                save_changes = False
            
            ' If sheet is protected, build array for summary and flag to not save.
            ElseIf protError Then
                count_err_j = count_err_j + 1
                count_err = count_err + 1
                
                ReDim Preserve unadj_files(1 To count_err)
                unadj_files(count_err) = wb.Name
                
                ReDim Preserve unadj_reason(1 To count_err)
                unadj_reason(count_err) = "WORKBOOK HAS PROTECTED SHEET"
                
                save_changes = False
            
            ' If no blank cells were found, error 1004 will be thrown. This is fine.
            ' If no errors thrown, error number will be 0.
            ' Any other error number should be reviewed thus do not save.
            ElseIf Err.Number <> 1004 And Err.Number <> 0 Then
                count_err_j = count_err_j + 1
                count_err = count_err + 1
                
                ReDim Preserve unadj_files(1 To count_err)
                unadj_files(count_err) = wb.Name
                
                ReDim Preserve unadj_reason(1 To count_err)
                unadj_reason(count_err) = "UNKNOWN ERROR"
                                
                ' Pause execution to debug since unknown error.
                Debug.Assert True
                                                
                save_changes = False
            
            ' If no errors then save file in formatted folder.
            Else
                count_j = count_j + 1
                count_tot = count_tot + 1
                
                ReDim Preserve adj_files(1 To count_tot)
                adj_files(count_tot) = wb.Name
                
                formatted_file_path = formatted_path & Left(wb.Name, InStrRev(wb.Name, ".") - 1) & ".CSV"
                
                ' Check if file already in folder.
                ' If not running quietly, ask user if to overwrite or append version number. If quietly,
                ' just append version number.
'                If Not quietly Then
'                    inp = MsgBox("Overwrite existing " & wb.Name & " in " & _
'                        Left(formatted_path, Len(formatted_path) - 1) & "? " & _
'                        "Select 'Yes' to overwrite, 'No' to append a version number, 'Cancel' to abort.", _
'                        vbYesNoCancel)
'                Else
'                    inp = vbNo
'                End If
'
'                If inp = vbCancel Then
'                    Exit Sub
'                ElseIf inp = vbNo Then
'                    v = 0
'                    While Dir(formatted_file_path) <> ""
'                        v = v + 1
'                        formatted_file_path = formatted_path & Left(wb.Name, InStrRev(wb.Name, ".") - 1) & "_v" & CStr(v) & ".CSV"
'                    Wend
'                End If
                
                wb.SaveAs Filename:=formatted_file_path, FileFormat:=xlCSV
            End If
            
            ' Close the file. If formatted, it will be saved in a new folder so delete original.
            wb.Close saveChanges:=False
            If save_changes Then Kill import_path & file_j
            
            ' Clear variables pertinent to file adjusted.
            Set wb = Nothing
            Err.Clear: protError = False: hdrError = False: dtError = False
            save_changes = True: bad_yr = False: bad_mm = False
            
            ' Loop to next file.
            file_j = Dir()
        Wend
                
        If Not quietly Then
            ' If a file of structure j was provided, print summary for current structure in loop.
            If no_files_j > 0 Then
                
                ' Build message for number of files adjusted.
                If count_j > 0 Then
                    adj_msg_j = vbNewLine & count_j & " input files formatted for AVID import."
                Else
                    adj_msg_j = ""
                End If
                
                ' Build message for number of files not adjusted because of errors.
                If count_err_j > 0 Then
                    err_msg_j = vbNewLine & count_err_j & " input files were not formatted due to errors."
                Else
                    err_msg_j = ""
                End If
                
                ' Compile summary message (differs slightly for last structure).
                If j < UBound(structure) Then
                    If MsgBox("Out of " & no_files_j & " files for " & structure(j) & ":" & _
                              vbNewLine & adj_msg_j & err_msg_j & vbNewLine & vbNewLine & _
                              "Proceed to next structure?", vbYesNo) = vbNo Then
                        Exit For
                    End If
                Else
                    MsgBox "Out of " & no_files_j & " files for " & structure(j) & ":" & _
                           vbNewLine & adj_msg_j & err_msg_j
                End If
            End If
        End If
    
    ' Loop to next structure.
    Next j
    
    ' If files were provided, print a summary.
    If no_files > 0 Then
    
        adj_msg = count_tot & " input files were formatted for AVID import."
        
        ' Build error message.
        If count_err > 0 Then
            err_msg = count_err & " input files had errors and were not formatted." & vbNewLine & vbNewLine & _
            "See details in 'Formatting Results' worksheet."
        Else
            err_msg = "No input files had errors. " & vbNewLine & vbNewLine & _
            "See details in 'Formatting Results' worksheet."
        End If
    
        msg = Trim("In summary, out of a total " & no_files & " input files:" & _
                   vbNewLine & vbNewLine & adj_msg & vbNewLine & vbNewLine & err_msg)
    Else
        msg = "No input files matched an AVID structure in the path provided. " & _
              "Double check the directory and try again."
    End If
    
    ' Print summary.
    MsgBox msg
    
    ' Build details sheet.
    With wb_home
        .Sheets("Formatting Results").Delete
        .Sheets.Add After:=.Sheets(2)
                
        With .Sheets(3)
            .Name = "Formatting Results"
            .Tab.ColorIndex = 10
            
            .Cells(1, 1).Value = "FORMATTED INPUT FILES"
            .Cells(1, 1).Style = "Heading 3"
            For i = 1 To count_tot
               .Cells(i + 1, 1).Value = adj_files(i)
            Next i
            .Columns(1).AutoFit
            
            .Cells(1, 3).Value = "ERRORED INPUT FILES"
            .Cells(1, 3).Style = "Heading 3"
            For i = 1 To count_err
               .Cells(i + 1, 3).Value = unadj_files(i)
            Next i
            .Columns(3).AutoFit
            
            .Cells(1, 4).Value = "ERROR DESCRIPTION"
            .Cells(1, 4).Style = "Heading 3"
            For i = 1 To count_err
               .Cells(i + 1, 4).Value = unadj_reason(i)
            Next i
            .Columns(4).AutoFit
            
            .Cells(1, 6).Value = "FILE WITH BAD HEADING"
            .Cells(1, 6).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 6).Value = bad_hdr_file(i)
            Next i
            .Columns(6).AutoFit
            
            .Cells(1, 7).Value = "BAD HEADING ROW #"
            .Cells(1, 7).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 7).Value = bad_hdr_row(i)
            Next i
            .Columns(7).AutoFit
            
            .Cells(1, 8).Value = "BAD HEADING COL #"
            .Cells(1, 8).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 8).Value = bad_hdr_col(i)
            Next i
            .Columns(8).AutoFit
            
            .Cells(1, 9).Value = "BAD HEADING ADDRESS"
            .Cells(1, 9).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 9).Value = bad_hdr_address(i)
            Next i
            .Columns(9).AutoFit
            
            .Cells(1, 10).Value = "INCORRECT HEADING TEXT"
            .Cells(1, 10).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 10).Value = bad_hdr(i)
            Next i
            .Columns(10).AutoFit
            
            .Cells(1, 11).Value = "EXPECTED HEADING TEXT"
            .Cells(1, 11).Style = "Heading 3"
            For i = 1 To count_bad_hdr
               .Cells(i + 1, 11).Value = correct_hdr(i)
            Next i
            .Columns(11).AutoFit
            
            .Range(.Cells(2, 6), .Cells(count_bad_hdr + 1, 11)).Sort _
                key1:=.Range("F2"), order1:=xlAscending, _
                key2:=.Range("H2"), order2:=xlAscending, _
                key3:=.Range("G2"), order3:=xlAscending
                
            .Cells(1, 1).Select
            
        End With
    End With
    
    ' Reset application properties.
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
    
End Sub


Public Sub openSaveAsCSV()
    'BY:                Cody B. Buehler
    'LAST UPDATED:      03/09/2022
    '
    'DESCRIPTION:
    '
     
    'Get valuation date & import directory path.
    valDate = Excel.Application.WorksheetFunction.EoMonth( _
                DateSerial(ActiveSheet.Range("inpYr").Value, _
                    Right(ActiveSheet.Range("inpQtr").Value, 1) * 3, 1), 0)
    import_path = "P:\Avid\Stage\Import\" & Year(valDate) & "-" & Right("0" & Month(valDate), 2) & "\"
    
    'Confirm import directory path.
    usr_input = _
        InputBox( _
            Prompt:="Confirm or update input file path.", _
            Title:="Adjust " & Right(Year(valDate), 2) & ActiveSheet.Range("inpQtr").Value & " Inputs", _
            Default:=import_path)
    If usr_input = "" Then Exit Sub Else import_path = usr_input
           
    If Not Right(import_path, 1) = "\" Then import_path = import_path & "\"
           
    file = Dir(import_path)
    
    'Macro optimization.
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With
        
    'Inner while loop w.r.t. files of structure(j).
    While file <> ""
        Set wb = Workbooks.Open(Filename:=(import_path & file))
        new_name = Left(wb.FullName, InStrRev(wb.FullName, ".") - 1) & ".CSV"
        wb.SaveAs Filename:=new_name, FileFormat:=xlCSV
        wb.Close
        Set wb = Nothing
        file = Dir()
    Wend
                
    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub



