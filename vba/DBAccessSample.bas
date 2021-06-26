Attribute VB_Name = "DBAccessSample"
Option Explicit

Sub ボタン1_Click()

    '高速化（描画中止）
    Application.ScreenUpdating = False


    Application.StatusBar = "検索中..."

    
    Dim sql As String
    sql = ActiveSheet.Cells(1, 1).Value
    
    If Trim(sql) = "" Then
        MsgBox "SQLを入力してください。", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    'クリア処理
    ActiveSheet.Range("A10:X1000").Value = ""
    
    
    Dim resultSet As Collection
    Set resultSet = selectDB(sql, "ユーザ名@ホスト名")
    If (resultSet Is Nothing) = False Then

        If resultSet.Count > 0 Then
        
            Dim iRow As Integer
            iRow = 10
            Dim record As Variant
            For Each record In resultSet
                Dim iColumn As Integer
                For iColumn = 0 To UBound(record)
                    ActiveSheet.Cells(iRow, iColumn + 1).Value = record(iColumn)
                Next iColumn
                
                
                iRow = iRow + 1
            Next
        Else
            MsgBox "data is not found.", vbOKOnly + vbExclamation
        End If
    End If
    
    Application.StatusBar = ""
    

End Sub

