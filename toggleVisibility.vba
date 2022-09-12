Private Sub ViewHide_Click()

Dim cols As String
Dim sheets As Worksheet
Dim wsCollection as Sheets

Set wsCollection = Worksheets
Set ws = Sheets("name")

'Select columns
cols = "A:Z"

    If ViewHide.Value = True Then

        Sheets(ws).Visible = xlSheetVeryHidden
        Sheets(ws).Visible = xlSheetVeryHidden
        
        Application.ActiveSheet.Columns(cols).Hidden = True
        Application.ActiveSheet.Columns(cols).Hidden = True
        
    Else
    
        Sheets(ws).Visible = xlSheetVisible
        Sheets(ws).Visible = xlSheetVisible
        
        Application.ActiveSheet.Columns(cols).Hidden = False
        Application.ActiveSheet.Columns(cols).Hidden = False
        
    End If
End Sub
