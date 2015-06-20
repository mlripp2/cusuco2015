Attribute VB_Name = "sheetReveal"
Sub sReveal()

    'Unlock workbook structure and make all sheets visible
    Dim S As Object
    Dim pWord3 As String
    pWord3 = InputBox("Please Enter the password")
    If pWord3 = "" Then Exit Sub
    
    ShtName = "Workbook as a whole"
    
    ActiveWorkbook.Unprotect Password:=pWord3
    
    For Each WS In Worksheets
        
        WS.Unprotect Password:=pWord3
        WS.Visible = True
    Next
    
    
    MsgBox "All sheets unprotected and visible."


End Sub
