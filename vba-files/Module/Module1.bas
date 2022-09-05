Attribute VB_Name = "Module1"
Sub explain_01()

    '対象指定
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets(1)
    
    
    ' 初期化
    targetSheet.Cells.Clear
    

    '変数初期値
    startNum = 1
    endNum = 10
    devideNum = 2
    surplus = 0
    outputRowNum = 1

    MsgBox "admin",,"hallo" 
    
    
    '繰り返し
    For nowNum = startNum To endNum
     surplus = nowNum Mod devideNum
        '分岐
      If surplus = 0 Then
        targetSheet.Range("A" & outputRowNum).Value = nowNum
        outputRowNum = outputRowNum + 1
     End If
    Next
    
    
End Sub
