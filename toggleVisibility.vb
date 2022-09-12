Private Sub ViewHide_Click()

Dim cols As String

cols = "F:H"

    If ViewHide.Value = True Then

        
        Sheets("02_Applications_&_systems").Visible = xlSheetVeryHidden
        Sheets("03_Pivot_PBC").Visible = xlSheetVeryHidden
        
        Application.ActiveSheet.Columns(cols).Hidden = True
        Application.ActiveSheet.Columns(cols).Hidden = True
        
    Else
    
        Sheets("02_Applications_&_systems").Visible = xlSheetVisible
        Sheets("03_Pivot_PBC").Visible = xlSheetVisible
        
        Application.ActiveSheet.Columns(cols).Hidden = False
        Application.ActiveSheet.Columns(cols).Hidden = False
        
    End If
End Sub
