Attribute VB_Name = "DBAccess"
Option Explicit
'#####################################################
'# ���s�ɂ́uParameter�v�V�[�g���K�v�B
'#  �{�ԃT�[�o�A�e�X�g�T�[�o�ւ̐ڑ��ؑցBSQL���O�̐؂�ւ��̂���
'# �Q�Ɛݒ�ɒǉ�
'#  Microsoft ActiveX Data Objects 2.8 Library
'#####################################################
Public Function connectDB(schema As String) As ADODB.Connection
    Dim conn As New ADODB.Connection
    Dim connectionString As String

    If schema = "���[�U��@TNS��" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '�{�ԃT�[�o�֐ڑ�����

            'Oracle�ڑ���
            'Data Source��tnsnames.ora��`�ς݂�TNS����ݒ肷��ꍇ
            connectionString = "Provider=OraOLEDB.Oracle;Data Source=TNS��;User ID=���[�U��;Password=�p�X���[�h;"


            'Oracle�ڑ���
            'tnsnames.ora����`�̏ꍇ
            'DDDD:�z�X�g���ABBBB:���[�U�[�ACCCC:�p�X���[�h
            datasource = ""
            datasource = datasource & "(DESCRIPTION ="
            datasource = datasource & "    (ADDRESS_LIST ="
            datasource = datasource & "      (ADDRESS = (PROTOCOL = TCP)(HOST = �z�X�g��)(PORT = 1521))"
            datasource = datasource & "    )"
            datasource = datasource & "    (CONNECT_DATA ="
            datasource = datasource & "      (SID = orcl)"
            datasource = datasource & "    )"
            datasource = datasource & "  )"
            connectionString = "Provider=OraOLEDB.Oracle;Data Source=" & datasource & ";User ID=���[�U��;Password=�p�X���[�h;"


        Else
            '�e�X�g�T�[�o�֐ڑ�����B



        End If

    ElseIf schema = "���[�U��@�z�X�g��" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '�{�ԃT�[�o�֐ڑ�����B

            'MYSQL�ڑ���
            connectionString = "Driver={MySQL ODBC 8.0 ANSI Driver}; SERVER=�z�X�g��; DATABASE=�f�[�^�x�[�X��; USER=���[�U��; PASSWORD=�p�X���[�h;"

        Else
            '�e�X�g�T�[�o�֐ڑ�����B

        End If

    ElseIf schema = "accdb�t�@�C����" Then
        If Worksheets("Parameter").Cells(3, 2).Value <> "1" Then
            '�{�ԃT�[�o�֐ڑ�����B

            'ACCDB(Access�t�@�C��)�ڑ���
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\TEMP\hoge.accdb;Mode=ReadWrite"

        Else
            '�e�X�g�T�[�o�֐ڑ�����B

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
'# �f�[�^�x�[�X�ڑ���ؒf
'#  �R�l�N�V�����������G���[���ɂ��N������悤��
'#  on error���ǉ����Ă���
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
'# �f�[�^�x�[�X����f�[�^���擾���A�R���N�V�����Ɋi�[
'#  ���ʃZ�b�g�̓R���N�V�����ɓ��ꂽ���Ƃɉ������
'#####################################################
Public Function selectDB(sql As String, schema As String) As Collection
On Error GoTo errPrg:

    outputLog (sql)

    '�f�[�^�x�[�X�ڑ�
    Dim conn As ADODB.Connection
    Set conn = connectDB(schema)
    
    'SQL�̎��s�ƒl�̎擾
    'SQL��`
    Dim rs As ADODB.Recordset
    
    '���ʃZ�b�g�̎擾
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
    
    '���ʃZ�b�g�̃N���[�Y
    rs.Close: Set rs = Nothing
    
    '�f�[�^�x�[�X�ؒf
    Call closeDB(conn, True)
    
    Set selectDB = collectionResultSet
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical
       
    '�f�[�^�x�[�X�ؒf
    Call closeDB(conn, False)
End Function

'#####################################################
'# UPDATE��/INSERT��/DELETE���p
'#
'#####################################################
Public Function updateDB(sql As String, schema As String)
On Error GoTo errPrg:
    outputLog (sql)
    
    updateDB = True

    '�f�[�^�x�[�X�ڑ�
    Dim conn As ADODB.Connection
    Set conn = connectDB(schema)
    
    '�g�����U�N�V�����J�n
    conn.Execute (sql)
    
    '�f�[�^�x�[�X�ؒf
    Call closeDB(conn, True)
    
    Exit Function
    
errPrg:
    updateDB = False

    '�G���[�\��
    MsgBox Err.Description, vbOKOnly + vbCritical
       
    '�f�[�^�x�[�X�ؒf
    Call closeDB(conn, False)
End Function

'#####################################################
'# �g�����U�N�V�����Ή���
'# UPDATE��/INSERT��/DELETE���p
'#####################################################
Public Function updateDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Boolean
On Error GoTo errPrg:
    outputLog (sql)

    updateDBOnTrans = True
    conn.Execute (sql)
    Exit Function
    
errPrg:
    If message Then
        '�G���[�\��
        If Mid(Err.Description, 1, 9) <> "ORA-00001" Then
            MsgBox Err.Description, vbOKOnly + vbCritical
        End If
    End If
    
    updateDBOnTrans = False
    
End Function

'#####################################################
'# �g�����U�N�V�����Ή���
'# �s���J�E���g
'#####################################################
Public Function selectCountDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Long
On Error GoTo errPrg:
    
    outputLog (sql)
    
    selectCountDBOnTrans = -1

    'SQL�̎��s�ƒl�̎擾
    'SQL��`
    Dim rs As ADODB.Recordset
    
    '���ʃZ�b�g�̎擾
    Set rs = conn.Execute(sql)
    
    Do Until rs.EOF
        selectCountDBOnTrans = rs.Fields(0)
        rs.MoveNext
    Loop
    
    '���ʃZ�b�g�̃N���[�Y
    rs.Close: Set rs = Nothing
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical
       
End Function

'#####################################################
'# �g�����U�N�V�����Ή���
'# ����
'#####################################################
Public Function selectDBOnTrans(conn As ADODB.Connection, sql As String, message As Boolean) As Collection
On Error GoTo errPrg:

    outputLog (sql)

    'SQL�̎��s�ƒl�̎擾
    'SQL��`
    Dim rs As ADODB.Recordset
    
    '���ʃZ�b�g�̎擾
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
    
    '���ʃZ�b�g�̃N���[�Y
    rs.Close: Set rs = Nothing
    
    Set selectDBOnTrans = collectionResultSet
    
    Exit Function
    
errPrg:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Function

'#####################################################
'# ���O�o��
'#
'#####################################################
Sub outputLog(sql As String)
    Worksheets("Parameter").Cells(1, 1).Value = sql
End Sub




