Private Sub ViewHide_Click()

Dim cols As String

'Select columns
cols = "A:Z"

    If ViewHide.Value = True Then

        Sheets("sheet_name").Visible = xlSheetVeryHidden
        Sheets("sheet_name").Visible = xlSheetVeryHidden
        
        Application.ActiveSheet.Columns(cols).Hidden = True
        Application.ActiveSheet.Columns(cols).Hidden = True
        
    Else
    
        Sheets("sheet_name").Visible = xlSheetVisible
        Sheets("sheet_name").Visible = xlSheetVisible
        
        Application.ActiveSheet.Columns(cols).Hidden = False
        Application.ActiveSheet.Columns(cols).Hidden = False
        
    End If
End Sub
