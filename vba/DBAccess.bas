Attribute VB_Name = "DBAccess"
Option Explicit
'#####################################################
'# 実行には「Parameter」シートが必要。
'#  本番サーバ、テストサーバへの接続切替。SQLログの切り替えのため
'# 参照設定に追加
'#  Microsoft ActiveX Data Objects 2.8 Library
'#####################################################
Public Function connectDB(schema As String) As ADODB.Connection
    Dim conn As New ADODB.Connection
    Dim connectionString As String

    If schema = "ユーザ名@TNS名" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '本番サーバへ接続する

            'Oracle接続例
            'Data Sourceにtnsnames.ora定義済みのTNS名を設定する場合
            connectionString = "Provider=OraOLEDB.Oracle;Data Source=TNS名;User ID=ユーザ名;Password=パスワード;"


            'Oracle接続例
            'tnsnames.ora未定義の場合
            'DDDD:ホスト名、BBBB:ユーザー、CCCC:パスワード
            datasource = ""
            datasource = datasource & "(DESCRIPTION ="
            datasource = datasource & "    (ADDRESS_LIST ="
            datasource = datasource & "      (ADDRESS = (PROTOCOL = TCP)(HOST = ホスト名)(PORT = 1521))"
            datasource = datasource & "    )"
            datasource = datasource & "    (CONNECT_DATA ="
            datasource = datasource & "      (SID = orcl)"
            datasource = datasource & "    )"
            datasource = datasource & "  )"
            connectionString = "Provider=OraOLEDB.Oracle;Data Source=" & datasource & ";User ID=ユーザ名;Password=パスワード;"


        Else
            'テストサーバへ接続する。



        End If

    ElseIf schema = "ユーザ名@ホスト名" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '本番サーバへ接続する。

            'MYSQL接続例
            connectionString = "Driver={MySQL ODBC 8.0 ANSI Driver}; SERVER=ホスト名; DATABASE=データベース名; USER=ユーザ名; PASSWORD=パスワード;"

        Else
            'テストサーバへ接続する。

        End If

    ElseIf schema = "accdbファイル名" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '本番サーバへ接続する。

            'ACCDB(Accessファイル)接続例
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\TEMP\hoge.accdb;Mode=ReadWrite"

        Else
            'テストサーバへ接続する。

        End If


    Else
        'error
        connectionString = ""
    End If
    
    conn.connectionString = connectionString
    conn.Open
    conn.BeginTrans
    
    Set connectDB = conn
End Function

'#####################################################
'# データベース接続を切断
'#  コネクションが無いエラー時にも起動するように
'#  on error区を追加している
'#####################################################
Public Sub closeDB(conn As ADODB.Connection, isCommit As Boolean)
On Error Resume Next
    If isCommit Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    conn.Close: Set conn = Nothing
    
End Sub

'#####################################################
'# データベースからデータを取得し、コレクションに格納
'#  結果セットはコレクションに入れたあとに解放する
'#####################################################
Public Function selectDB(sql As String, schema As String) As Collection
On Error GoTo errPrg:

    outputLog (sql)

    'データベース接続
    Dim conn As ADODB.Connection
    Set conn = connectDB(schema)
    
    'SQLの実行と値の取得
    'SQL定義
    Dim rs As ADODB.Recordset
    
    '結果セットの取得
    Set rs = conn.Execute(sql)
    'If rs Is Nothing Then
    '    MsgBox "error"
    'End If
    
    Dim collectionResultSet As New Collection
    
    Do Until rs.EOF
        Dim record() As Variant
        ReDim record(rs.Fields.Count - 1)
        Dim i As Integer
        For i = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields(i)) = False Then
                record(i) = rs.Fields(i)
            End If
        Next i
        
        collectionResultSet.Add (record)
        
        rs.MoveNext
    Loop
    
    '結果セットのクローズ
    rs.Close: Set rs = Nothing
    
    'データベース切断
    Call closeDB(conn, True)
    
    Set selectDB = collectionResultSet
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical
       
    'データベース切断
    Call closeDB(conn, False)
End Function

'#####################################################
'# UPDATE文/INSERT文/DELETE文用
'#
'#####################################################
Public Function updateDB(sql As String, schema As String)
On Error GoTo errPrg:
    outputLog (sql)
    
    updateDB = True

    'データベース接続
    Dim conn As ADODB.Connection
    Set conn = connectDB(schema)
    
    'トランザクション開始
    conn.Execute (sql)
    
    'データベース切断
    Call closeDB(conn, True)
    
    Exit Function
    
errPrg:
    updateDB = False

    'エラー表示
    MsgBox Err.Description, vbOKOnly + vbCritical
       
    'データベース切断
    Call closeDB(conn, False)
End Function

'#####################################################
'# トランザクション対応版
'# UPDATE文/INSERT文/DELETE文用
'#####################################################
Public Function updateDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Boolean
On Error GoTo errPrg:
    outputLog (sql)

    updateDBOnTrans = True
    conn.Execute (sql)
    Exit Function
    
errPrg:
    If message Then
        'エラー表示
        If Mid(Err.Description, 1, 9) <> "ORA-00001" Then
            MsgBox Err.Description, vbOKOnly + vbCritical
        End If
    End If
    
    updateDBOnTrans = False
    
End Function

'#####################################################
'# トランザクション対応版
'# 行数カウント
'#####################################################
Public Function selectCountDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Long
On Error GoTo errPrg:
    
    outputLog (sql)
    
    selectCountDBOnTrans = -1

    'SQLの実行と値の取得
    'SQL定義
    Dim rs As ADODB.Recordset
    
    '結果セットの取得
    Set rs = conn.Execute(sql)
    
    Do Until rs.EOF
        selectCountDBOnTrans = rs.Fields(0)
        rs.MoveNext
    Loop
    
    '結果セットのクローズ
    rs.Close: Set rs = Nothing
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical
       
End Function

'#####################################################
'# トランザクション対応版
'# 検索
'#####################################################
Public Function selectDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Collection
On Error GoTo errPrg:

    outputLog (sql)

    'SQLの実行と値の取得
    'SQL定義
    Dim rs As ADODB.Recordset
    
    '結果セットの取得
    Set rs = conn.Execute(sql)
    'If rs Is Nothing Then
    '    MsgBox "error"
    'End If
    
    Dim collectionResultSet As New Collection
    
    Do Until rs.EOF
        Dim record() As Variant
        ReDim record(rs.Fields.Count - 1)
        Dim i As Integer
        For i = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields(i)) = False Then
                record(i) = rs.Fields(i)
            End If
        Next i
        
        collectionResultSet.Add (record)
        
        rs.MoveNext
    Loop
    
    '結果セットのクローズ
    rs.Close: Set rs = Nothing
    
    Set selectDBOnTrans = collectionResultSet
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Function

'#####################################################
'# ログ出力
'#
'#####################################################
Sub outputLog(sql As String)
    Worksheets("Parameter").Cells(1, 1).Value = sql
End Sub




