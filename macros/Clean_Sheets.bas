Attribute VB_Name = "Clean_Sheets"
Sub cleanSheet()

    'Clears previous data prep



        With Sheets("Rec_Prep")

            Sheets("Rec_Prep").Activate

            Rows("3:" & Rows.Count).ClearContents

        End With

        With Sheets("Records")

            Sheets("Records").Activate

            Rows("3:" & Rows.Count).ClearContents

        End With
        
        With Sheets("Occasion")

            Sheets("Occasion").Activate

            Rows("3:" & Rows.Count).ClearContents

        End With
        
        With Sheets("Occ_Prep")

            Sheets("Occ_Prep").Activate

            Rows("3:" & Rows.Count).ClearContents

        End With
        
        With Sheets("Data")

            Sheets("Data").Activate

            Rows("3:" & Rows.Count).ClearContents

        End With

End Sub
