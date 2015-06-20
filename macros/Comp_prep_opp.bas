Attribute VB_Name = "Comp_prep_opp"
Sub Comp_Prep_Op()
    
    'Row counting
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim Lrow2 As Long
   
    'Column counting
    Dim LastcolOcc As Long
    Dim LastcolRec As Long
    
    'Ranges
    Dim OccRange As Range
    Dim RecRange As Range
    Dim TotOcc As Range
    Dim TotRec As Range
    Dim TotOccPrep As Range
    Dim TotRecPrep As Range
    
    
    'Changing calc mode
    Dim CalcMode As Long
    Dim ViewMode As Long
    
    'Setting worksheets to return to A1
    Dim WS As Worksheet, flg As Boolean

    
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = False
    End With
    
    'Unlock workbook structure and make all sheets visible
    Dim S As Object
    Dim pWord3 As String
    pWord3 = InputBox("Please Enter the password")
    If pWord3 = "" Then Exit Sub
    
    ShtName = "Workbook as a whole"
    On Error GoTo errorTrap1
    ActiveWorkbook.Unprotect Password:=pWord3
    
    For Each WS In Worksheets
        On Error GoTo errorTrap1
        WS.Unprotect Password:=pWord3
        WS.Visible = True
    Next
    
    
    MsgBox "All sheets unprotected and visible."
    
'    'Clears previous data prep
'
'
'
'        With Sheets("Rec_Prep")
'
'            Sheets("Rec_Prep").Activate
'
'            Rows("3:" & Rows.Count).ClearContents
'
'        End With
'
'        With Sheets("Records")
'
'            Sheets("Records").Activate
'
'            Rows("3:" & Rows.Count).ClearContents
'
'        End With
    
    
    'Counts rows and returns message box with number
    With Sheets("Data")
    
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        Lrow = .UsedRange.Rows.Count
        MsgBox "Rows are: " & Lrow - 1
    
    End With
    
'            'Clears out all but the top two rows of rec_prep sheet
'        With Sheets("Rec_Prep")
'
'            Sheets("Rec_Prep").Activate
'
'            Rows("3:" & Rows.Count).ClearContents
'
'        End With
'
'        With Sheets("Records")
'
'            Sheets("Records").Activate
'
'            Rows("3:" & Rows.Count).ClearContents
'
'    End With
    
    'Copies formulas in Occasion down to match data entered
    With Sheets("Occasion")

        Sheets("Occasion").Activate


        LastcolOcc = .UsedRange.Columns.Count

        'Copies formulas down
        Range("B2").Select
        Set OccRange = Range(ActiveCell, ActiveCell.Offset(0, LastcolOcc - 2))
        OccRange.Copy
        Range("B3").Select
        Range(ActiveCell, ActiveCell.Offset(Lrow - 3, LastcolOcc - 2)).Select
        Sheets("Occasion").Paste

        'Copies Occasion column formulas down
        Range("A3").Select
        ActiveCell.Copy
        Range(ActiveCell, ActiveCell.Offset(Lrow - 2, 0)).Select
        Sheets("Occasion").Paste

        'Copies occasion numbers to Records sheet
        Columns("A:A").Copy

        Sheets("Records").Activate
        Columns("A:A").PasteSpecial xlPasteValues


        'Pastes values only to Occ_Prep sheet
        Sheets("Occasion").Activate
        Range("A1").Select
        Set TotOcc = Range(ActiveCell, ActiveCell.Offset(Lrow - 1, LastcolOcc - 1))
        TotOcc.Copy
        Sheets("Occ_Prep").Activate
        Range("A1").Select
        Set TotOccPrep = Range(ActiveCell, ActiveCell.Offset(Lrow - 1, LastcolOcc - 1))
        TotOccPrep.PasteSpecial xlPasteValues




    End With


        
    
    
    
    'Copies formulas in records down to match data entered
    With Sheets("Records")

        Sheets("Records").Activate
        
        LastcolRec = .UsedRange.Columns.Count


        Range("B2").Select
        Set RecRange = Range(ActiveCell, ActiveCell.Offset(0, LastcolRec - 1))
        RecRange.Copy
        Range("B3").Select
        Range(ActiveCell, ActiveCell.Offset(Lrow - 3, LastcolRec - 1)).Select
        Sheets("Records").Paste
        
        
        'Copies values only to RecPrep sheet
        Sheets("Records").Activate
        Range("A1").Select
        Set TotRec = Range(ActiveCell, ActiveCell.Offset(Lrow - 1, LastcolRec - 1))
        TotRec.Copy
        Sheets("Rec_Prep").Activate
        Range("A1").Select
        Set TotRecPrep = Range(ActiveCell, ActiveCell.Offset(Lrow - 1, LastcolRec - 1))
        TotRecPrep.PasteSpecial xlPasteValues

    End With
    
    With Sheets("Occ_Prep")

        'Cleans the occ_prep sheet ready for upload (Column and value can be changed)
        Sheets("Occ_Prep").Activate

                'Set the first and last row to loop through
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow2 = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow2, "B")


                If Not IsError(.Value) Then

                    If .Value = "" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow2

    End With
    MsgBox "Data prepared for upload"
    'Return cells to A1
    For Each WS In ActiveWorkbook.Sheets
        WS.Activate
        WS.[a1].Select
    Next WS
        
    ActiveWorkbook.Worksheets(1).Activate
    
    Exit Sub
    
errorTrap1:
    MsgBox "Process failed: please check your password and try again"
    Exit Sub
    
        

    
    
End Sub

