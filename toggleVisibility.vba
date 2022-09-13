'Author: Tim Lui
'Ownership Avega S.à.r.l.
'Be careful: The script has no error handling in place

Option Explicit

Private Sub ToggleButton1_Click()

    'Declare variables
    Dim cols As String
    Dim i As Integer
    Dim sheetNames() As String
    
    'Declare the names
    sheetNames = Split("Sheet2,Sheet3", ",")
    
    'Select columns
    cols = "A:B"

    If ToggleButton1.Value = True Then
        
        'Hide the selected sheets
        For i = LBound(sheetNames) To UBound(sheetNames)
        
            Sheets(sheetNames(i)).Visible = xlSheetVeryHidden
            
            If cols <> vbNullString Then
                Application.ActiveSheet.Columns(cols).Hidden = True
            End If
        Next i
            
    Else
        
        'Unhide the selected sheets
        For i = LBound(sheetNames) To UBound(sheetNames)
        
            Sheets(sheetNames(i)).Visible = xlSheetVisible
            
            If cols <> vbNullString Then
                Application.ActiveSheet.Columns(cols).Hidden = False
            End If
        Next i
    End If
End Sub