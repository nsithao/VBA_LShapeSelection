Attribute VB_Name = "Module1"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub LShape()
Attribute LShape.VB_Description = "L-shape"
Attribute LShape.VB_ProcData.VB_Invoke_Func = "X\n14"
'
' autoSelectRowCol Macro
' when triggered, auto select the whole range start from the selected cell to the first cell in the same row
'
    ' str = ActiveCell.Address(False, False)  ' => G16
    ' MsgBox Split(ActiveCell.Address(1, 0), "$")(0) & "_" & Split(ActiveCell.Address(1, 0), "$")(1)
    ' => G_16
        
    'Range("A15:H15,H14,H13,H12,H11,H10,H9,H8,H7,H6,H5,H4,H3,H2,H1").Select '=> L shape selection
    'Range("A15:H15,H15:H1").Select  '=> L shape selection
    
    'Dim abc As String
    'abc = Sheets(1).Range("A1").Value 'contain value such as B1, AB31
    'Range (abc & ":" & abc1) => build range from string
    
    Dim selectedCell As String
    Dim rowNumber As Integer
    Dim colLetter As String
    
    selectedCell = ActiveCell.Address(False, False)      ' get G16
    rowNumber = Split(ActiveCell.Address(1, 0), "$")(1)  ' get 16
    colLetter = Split(ActiveCell.Address(1, 0), "$")(0)  ' get G
    
    Dim rowBeginCell As String
    Dim colBeginCell As String
    
    rowBeginCell = "A" & rowNumber
    colBeginCell = colLetter & "1"
    
    Dim finalString As String
    
    finalString = "" & rowBeginCell & ":" & selectedCell & "," & selectedCell & ":" & colBeginCell & ""
    
    Range(finalString).Select ' produce a L shape selection to refer row and col in crosshair style
End Sub
