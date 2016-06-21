Attribute VB_Name = "ProposalCombiner"
Sub ProposalCombiner()
' Proposal Dataset Combiner
' Written by Paul Hively, 6/20/2016
' Takes Communications proposal data, and follow-up survey data, and combines into
' a single field (joining on ID), deleting extraneous columns

' Declarations
Dim survey_cols As Variant
Dim addl_cols As Variant
Dim concat_cols As Variant
Dim col_order As Variant
Dim final_colnames As Variant

' ========== Variables - feel free to edit ==========
' List the column names used in the survey whose data should appear as-is
survey_cols = Array("V9", "Q1", "Q2", "Q3", "Q5", "Q9")
' List of column names from the follow-up survey whose data should appear as-is
addl_cols = Array("QID5", "QID22_TEXT")
' List of column names to concatenate into Purpose
concat_cols = Array( _
    "Q8_1", "Q8_2", "Q8_3", "Q8_5", "Q8_6", "Q8_7", "Q8_8", "Q8_9", "Q8_10", _
    "Q8_11", "Q8_12", "Q8_13", "Q8_14_TEXT" _
)
' Order in which columns should appear; use "NEW.COL" to insert a blank column
col_order = Array( _
    "V9", "NEW.COL", "QID22_TEXT", "NEW.COL", "NEW.COL", _
    "Q1", "Q2", "Q3", "Purpose", _
    "QID5", "Purpose", "Q5", _
    "NEW.COL", "NEW.COL", "NEW.COL" _
)
' Ordered header names to use in the final data; should match up to col_order
final_colnames = Array( _
    "Date of Request", "Date of Mtg", "Date Promised", "Date Completed", "Writer", _
    "Requested By", "Prospect Name", "Entity ID ", "Purpose", _
    "Design Assistance Needed", "Center Ask", "Ask Amount/Range", _
    "Final Review By", "Final Draft Saved to Team Fldr (X)", "Notes" _
)

' ==========  Do not edit below this line  ==========

' Clear Results tab
Sheets("Results").Cells.Clear

' Copy survey columns to Results tab as-is
' Loop through the column names in survey_cols
For Each colname In survey_cols
    ' Try matching column with current column name
    On Error GoTo BadColName
        Sheets("Paste Survey Data").Cells.Find(colname, , xlValues, xlWhole).EntireColumn.Copy
    On Error GoTo 0
    ' Insert to column A of Results tab
    Sheets("Results").Range("A1").Insert Shift:=xlShiftToRight
Next colname

' Variables and constants for concatenation
Dim col As Range
Dim row As Range
Dim nrow As Integer
    nrow = Sheets("Paste Survey Data").UsedRange.Rows.Count
Const sep As String = ", "
' Insert empty column
' .Activate is not best practice but we expect << 1000 rows of data so the slowdown won't be noticeable
Sheets("Results").Activate
Range("A:A").Insert
' Create concatenated Centers column
For Each colname In concat_cols
    ' Grab the data associated with the current column name
    On Error GoTo BadColName
        Sheets("Paste Survey Data").Activate
        Set col = Cells.Find(colname, , xlValues, xlWhole).EntireColumn.Range(Cells(1, 1), Cells(nrow, 1))
    On Error GoTo 0
    Debug.Print col.Rows.Count
    ' Concatenate each non-blank entry to purpose
    Sheets("Results").Activate
    For Each row In col
        ' If there is text entered
        If row > 0 Then
            ' If there is nothing in the cell then initialize
            If Range("A" & row.row).Value = 0 Then
                Range("A" & row.row).Value = row
            ' If there is already something in the cell, concatenate
            Else
                Range("A" & row.row).Value = Range("A" & row.row).Value & sep & row
            End If
        End If
    Next row
Next colname
' Change header to Purpose
Sheets("Results").Range("A1").Value = "Purpose"

' Append additional fields from Follow-Up Data
For Each colname In addl_cols
    On Error GoTo BadColName
    On Error GoTo 0
Next colname

' Reorder columns as needed
For Each colname In col_order
    On Error GoTo BadColName
    On Error GoTo 0
Next colname

' Check for double headers (rows 1 and 2) and rename row 1 headers


' Done!
Exit Sub

' ====== Error handling ======
' Bad column name
BadColName:
    MsgBox "Warning: No column named " & colname
Resume Next

End Sub
