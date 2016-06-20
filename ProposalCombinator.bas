Attribute VB_Name = "ProposalCombinator"
Sub Combinator()
' Proposal Dataset Combinator
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


' Create concatenated Centers column


' Append additional fields from Follow-Up Data


' Check for double headers (rows 1 and 2)


' Reorder and rename columns as needed


' Done!
End Sub
