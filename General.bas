Attribute VB_Name = "General"
Sub StyleKiller()
Dim N As Long, i As Long

'This clears all custom cell styles from a workbook. Since many spreadsheets we use are ancient, duplicate cell styles have been copied over.
'The number of duplicate styles is so large that it prevents the default styles from being displayed.
    With ActiveWorkbook
    
        N = .Styles.Count
        For i = N To 1 Step -1
            If Not .Styles(i).BuiltIn Then .Styles(i).Delete
        Next i
    End With
End Sub

Sub Relative2Absolute()
Dim sel As Range: Set sel = Application.Selection
Dim cell As Range, i As Integer, j As Integer, k As Integer
ReDim formulas(0) As Variant

'This changes all formula references from relative to absolute in a given selection.

    
    For Each cell In sel
        'First see if each cell contains a formula; if not then skip.
        If cell.HasFormula = True Then
            
            'If the cell does have a formula, clear the previous array values and reset variables.
            Erase formulas
            ReDim formulas(0)
            i = 0
            j = 0
            formulas(0) = cell.Formula
            
            'The ConvertFormula method returns a #VALUE error for formulas over length ~110. Need to split up formulas if length is longer.
            While Len(formulas(i)) > 50
                ReDim Preserve formulas(UBound(formulas) + 1)
                
                'This is where we want to split our formula. The substrings need to still be valid formulas,
                'so we need to split this at a binary operator within the formula (assuming there is one due to long length).
                'We check for the location of the nearest binary operator from 100 characters out in the formula (to the left) and
                'return the max location, i.e. the longest valid formula substring under length 100.
                j = -1 + WorksheetFunction.Max(InStrRev(Left(formulas(i), 50), "+"), _
                                            InStrRev(Left(formulas(i), 50), "-"), _
                                            InStrRev(Left(formulas(i), 50), "*"), _
                                            InStrRev(Left(formulas(i), 50), "/"))
                     
                'Now split the formula at the acceptable position.
                formulas(i + 1) = Right(formulas(i), Len(formulas(i)) - j)
                formulas(i) = Left(formulas(i), j)
                i = i + 1
            Wend
            
            'Now convert each formula segment.
            For k = 0 To i
                On Error Resume Next
                formulas(k) = Application.ConvertFormula(Formula:=formulas(k), _
                                                            FromReferenceStyle:=xlA1, _
                                                            ToReferenceStyle:=xlA1, _
                                                            ToAbsolute:=xlAbsolute)
            Next k
            
            'Lastly combine all segments into one string and set cell formula to cobined string.
            cell.Formula = Join(formulas)
        End If
    Next cell
End Sub

Sub Absolute2Relative()
Dim sel As Range: Set sel = Application.Selection
Dim cell As Range, i As Integer, j As Integer, k As Integer
ReDim formulas(0) As Variant

'This changes all formula references from relative to absolute in a given selection.

    
    For Each cell In sel
        'First see if each cell contains a formula; if not then skip.
        If cell.HasFormula = True Then
            
            'If the cell does have a formula, clear the previous array values and reset variables.
            Erase formulas
            ReDim formulas(0)
            i = 0
            j = 0
            formulas(0) = cell.Formula
            
            'The ConvertFormula method returns a #VALUE error for formulas over length ~110. Need to split up formulas if length is longer.
            While Len(formulas(i)) > 50
                ReDim Preserve formulas(UBound(formulas) + 1)
                
                'This is where we want to split our formula. The substrings need to still be valid formulas,
                'so we need to split this at a binary operator within the formula (assuming there is one due to long length).
                'We check for the location of the nearest binary operator from 100 characters out in the formula (to the left) and
                'return the max location, i.e. the longest valid formula substring under length 100.
                j = -1 + WorksheetFunction.Max(InStrRev(Left(formulas(i), 50), "+"), _
                                            InStrRev(Left(formulas(i), 50), "-"), _
                                            InStrRev(Left(formulas(i), 50), "*"), _
                                            InStrRev(Left(formulas(i), 50), "/"))
                     
                'Now split the formula at the acceptable position.
                formulas(i + 1) = Right(formulas(i), Len(formulas(i)) - j)
                formulas(i) = Left(formulas(i), j)
                i = i + 1
            Wend
            
            'Now convert each formula segment.
            For k = 0 To i
                On Error Resume Next
                formulas(k) = Application.ConvertFormula(Formula:=formulas(k), _
                                                            FromReferenceStyle:=xlA1, _
                                                            ToReferenceStyle:=xlA1, _
                                                            ToAbsolute:=xlRelative)
            Next k
            
            'Lastly combine all segments into one string and set cell formula to cobined string.
            cell.Formula = Join(formulas)
        End If
    Next cell
End Sub

Function distinct(x As Variant) As Variant
' BY:           Cody B. Buehler
' CREATED:      03/13/2022
'
' DESCRIPTION:  Returns an array containing only unique values in provided range or array.
'
    Dim vals As Variant: Dim uniques() As Variant
    vals = Application.Transpose(x)
    n_distinct = 0
     
    For Each c In vals
        dup = False
        
        If n_distinct > 0 Then
            For i = 1 To n_distinct
                If c = vals(i) Then
                    dup = True
                    Exit For
                End If
            Next i
        End If
        
        If Not (dup Or IsEmpty(c)) Then
            n_distinct = n_distinct + 1
            ReDim Preserve vals(n_distinct)
            vals(n_distinct) = c
        End If
    Next c
    
    distinct = vals
    
End Function

Sub ConvertCSVToXlsx(ByRef folderName As String)
'BY:                Cody B. Buehler
'LAST UPDATED:      2/03/2021
   
'DESCRIPTION:       This subroutine converts comma-separated values (.CSV) into .xlsx files.
    
    Dim myfile, oldfname, newfname As String
    Dim workfile

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
'   Capture name of current file
    myfile = ActiveWorkbook.Name
    
'   Set folder name to work through
'    folderName = "P:\Avid\Stage\Import\2020-12\"
    
'   Loop through all CSV filres in folder
    workfile = Dir(folderName & "*.csv")
    Do While workfile <> ""
'       Open CSV file
        Workbooks.Open Filename:=folderName & workfile
'       Capture name of old CSV file
        oldfname = ActiveWorkbook.FullName
'       Convert to XLSX
        newfname = folderName & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & ".xlsx"
        ActiveWorkbook.SaveAs Filename:=newfname, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
'       Delete old CSV file
        Kill oldfname
        Windows(myfile).Activate
        workfile = Dir()
    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Sub ConvertDEDToXlsx(ByRef folderName As String)
'BY:                Cody B. Buehler
'LAST UPDATED:      2/02/2021
   
'DESCRIPTION:       This subroutine converts CASE .DED extracts into .xlsx files.

    Dim myfile, oldfname, newfname As String
    Dim workfile
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
'   Capture name of current file
    myfile = ActiveWorkbook.Name
    
'   Loop through all DED files in folder.
    workfile = Dir(folderName & "*.DED")
    Do While workfile <> ""
'       Open DED file
        Workbooks.OpenText Filename:=folderName & workfile, _
            DataType:=xlDelimited, _
            comma:=True, _
            TextQualifier:=xlTextQualifierDoubleQuote
'       Capture name of old DED file.
        oldfname = ActiveWorkbook.FullName
'       Convert to XLSX.
        newfname = folderName & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & ".xlsx"
        ActiveWorkbook.SaveAs Filename:=newfname, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
'       Delete old DED file.
        Kill oldfname
        Windows(myfile).Activate
        workfile = Dir()
    Loop
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub ConvertXlsxToCSV(ByRef folderName As String)
    
    Dim myfile As String
    Dim oldfname As String, newfname As String
    Dim workfile
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Capture name of current file
    myfile = ActiveWorkbook.Name
        
    'Add slash at end of path in case it was not included with path.
    If Right(folderName, 1) <> "\" Then
        folderName = folderName & "\"
    End If
    
    'Loop through all xlsx files in folder
    workfile = Dir(folderName & "*.xlsx")
    Do While workfile <> ""
        'Open xlsx file
        Workbooks.Open Filename:=folderName & workfile
        'Capture name of old xlsx file
        oldfname = ActiveWorkbook.FullName
        'Convert to CSV
        newfname = folderName & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".CSV"
        ActiveWorkbook.SaveAs Filename:=newfname, FileFormat:=xlCSV, CreateBackup:=False
        ActiveWorkbook.Close
        'Delete old xlsx file
        Kill oldfname
        Windows(myfile).Activate
        workfile = Dir()
    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Sub test(Optional arr As Variant)
    If IsMissing(arr) Then arr = distinct(ActiveWorkbook.Sheets("Formatting Results").Range("D2:D14"))

    Debug.Print ("----------------------Debugging Array----------------------")
    Debug.Print ("Array row count: " & UBound(arr))
    Debug.Print ("Array col count: " & UBound(arr, 2))
    
    Dim rowString As String
    rowString = ""
    
    For r = 1 To UBound(arr)
        rowString = ""
        For c = 1 To UBound(arr, 2)
            If c = 1 Then
                rowString = arr(r, c)
            Else
                rowString = rowString & "," & arr(r, c)
            End If
        Next c
        Debug.Print (rowString)
    Next r
    
    Debug.Print ("----------------------Debug Complete-----------------------")
End Sub
