Attribute VB_Name = "LockWorkBook"
Sub LockWB()


    
    'Unlock workbook structure and make all sheets visible
    Dim WS As Worksheet
    
    'Dim S As Object
    Dim pWord5 As String
    pWord5 = "19sigBurd79"
    If pWord5 = "" Then Exit Sub
    'MsgBox "Password Set"

    
    For Each WS In Worksheets
        
        On Error GoTo errorTrap1
        WS.Protect Password:=pWord5
        'WS.Visible = True
    Next WS
    
    
    Worksheets("Occasion").Visible = xlSheetVeryHidden
    Worksheets("Records").Visible = xlSheetVeryHidden
    Worksheets("Occ_Prep").Visible = xlSheetVeryHidden
    Worksheets("Rec_Prep").Visible = xlSheetVeryHidden
    Worksheets("Lists").Visible = xlSheetVeryHidden
'    S(Data).Visible = True
'    S(Occasion).Visible = False
'
    
    
    ShtName = "Workbook as a whole"
    On Error GoTo errorTrap2
    ActiveWorkbook.Protect Password:=pWord5
    
            Application.ScreenUpdating = False
        


    MsgBox "Workbook locked"
    'Return cells to A1
        For Each WS In ActiveWorkbook.Sheets
            WS.Activate
            WS.[a1].Select
        Next WS
        
    ActiveWorkbook.Sheets("Welcome").Activate
    
    Exit Sub
    
errorTrap1:
    MsgBox "Workbook already locked"
    Exit Sub

errorTrap2:
    MsgBox "Shit"
    Exit Sub
    
End Sub
