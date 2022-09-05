Attribute VB_Name = "Module1"
Sub explain_01()

    ' import
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets(1)
    
    
    ' init
    targetSheet.Cells.Clear
    

    ' variable
    startNum = 1
    endNum = 10
    devideNum = 2
    surplus = 0
    outputRowNum = 1

    MsgBox "admin", , "hallo"
    
    ' loop
    For nowNum = startNum To endNum
     surplus = nowNum Mod devideNum
        'if
        If surplus = 0 Then
        targetSheet.Range("A" & outputRowNum).Value = nowNum
        outputRowNum = outputRowNum + 1
     End If
    Next
    
    
End Sub
