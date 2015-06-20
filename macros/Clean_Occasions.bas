Attribute VB_Name = "Clean_Occasions"
Sub CleanOcc()

    'Row counting
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim Lrow2 As Long

    With Sheets("Occ_Prep")

        'Cleans the occ_prep sheet ready for upload (Column and value can be changed)
        Sheets("Occ_Prep").Activate

        'Set the first and last row to loop through
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow2 = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow2, "M")


                If Not IsError(.Value) Then

                    If .Value = "0" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow2

    End With



End Sub
