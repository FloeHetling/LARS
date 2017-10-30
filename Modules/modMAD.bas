Attribute VB_Name = "modMAD"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''' Manage Audit Data (MAD) module ''''''''''''''''''''''''
        ''''''''''''''''''' ������ ������ � �����-������� ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit '�������� "�������������" ����������� - ����������� ���������� ����
Public isSQLChecked As Boolean
Public tmpSQLAvailable As Boolean

Public Function RegGetAuditData(ByVal AuditProp As String) As String
Dim AuditValue As String
'�������� ���������� �� ������� � �������� �� � ������� ������ ������ � ��������
'���� ������ ��� - ��� � �����
Call fReadValue("HKLM", "Software\LARS", AuditProp, "S", "��� ������", AuditValue)
RegGetAuditData = AuditValue
End Function

Public Function RegPutAuditData(ByVal AuditProp As String, ByVal auditdata As String)
'������ ������� ���������� ������� Call, ���������� ��� ��������� - ��������������
'���� ������ ������ � ����� ������ ������ ������
Call fWriteValue("HKLM", "Software\LARS", AuditProp, "S", auditdata)
End Function

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

Public Function SQLExecute(ByVal SQLRequestString As String, SQLMode As laSQLMode, Optional ByVal ParameterToRead As String) As Variant
On Error GoTo SQL_error
Dim SQLResponse As Variant
Dim SQL As New ADODB.Connection
    Dim SQLData As New ADODB.Recordset
    Dim SQLRequest As String, SQLAPRequest As String
    SQL.Open SQLConnString
    WriteToLog "�������� ������� SQLAuditData SQLExecute. ������ ����������:" & vbCrLf & SQLRequestString
    SQLData.Open SQLRequestString, SQL, adOpenKeyset
        If SQLMode = laRX Then
            SQLExecute = SQLData.Fields(ParameterToRead).Value
        End If
    SQL.Close
    Exit Function
'frmWriteAuditData.tDeb.Text = frmWriteAuditData.tDeb.Text & vbCrLf &
SQL_error:
Dim SQLErrNumber As Long, SQLErrDescription As String
SQLErrNumber = Err.Number
SQLErrDescription = Err.description
    If SQLErrNumber <> 0 Then
        WriteToLog " "
        WriteToLog "������ SQL ������� �� ������:"
        WriteToLog "������ SQL " & SQLErrNumber & ":" & vbCrLf & SQLErrDescription
    End If
SQLExecute = SQLErrNumber
End Function

Public Function isSQLAvailable() As Boolean
Dim sqlCheckIfAvailable As Long
If isSQLChecked = False Then
    sqlCheckIfAvailable = SQLExecute("SELECT * FROM dbo.larspc", laRX)
        If sqlCheckIfAvailable = -2147467259 Then
            tmpSQLAvailable = False
            WriteToLog " "
            WriteToLog "������ �������� SQL �������: ������ SQL �� ��������"
            WriteToLog " "
        Else
            tmpSQLAvailable = True
            WriteToLog " "
            WriteToLog "������ �������� SQL �������: �������� ����������"
            WriteToLog " "
        End If
    isSQLChecked = True
End If
isSQLAvailable = tmpSQLAvailable
End Function
