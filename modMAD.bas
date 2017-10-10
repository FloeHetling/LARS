Attribute VB_Name = "modMAD"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''' Manage Audit Data (MAD) module ''''''''''''''''''''''''
        ''''''''''''''''''' ������ ������ � �����-������� ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit '�������� "�������������" ����������� - ����������� ���������� ����

Public Function PopulateAuditData()
MsgBox "�������: ��������� �� ��������. ��������� � ���������� /edit ��� �������������� �����-����������", vbInformation, LARSver
MsgBox "TODO: ������ ������������ ����������!!!", vbExclamation, LARSver
End Function

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

Public Function cbExists(SearchString As String, ComboBoxForCheck As ComboBox) As Boolean
Dim cItem As Integer
                For cItem = 0 To ComboBoxForCheck.ListCount Step 1
                     If SearchString = ComboBoxForCheck.List(cItem) Then
                     cbExists = True
                     Exit Function
                     End If
                Next cItem
                cbExists = False
End Function

Public Function SQLExecute(ByVal SQLRequestString As String, SQLMode As laSQLMode, Optional ByVal ParameterToRead As String) As Variant
On Error GoTo SQL_error
Dim SQLResponse As Variant
Dim SQL As New ADODB.Connection
    Dim SQLData As New ADODB.Recordset
    Dim SQLRequest As String, SQLAPRequest As String
    DoEvents
    SQL.Open _
        "Provider = SQLNCLI11.1;" & _
        "Data Source=WS0006\SQLEXPRESS;" & _
        "Initial Catalog=AIDA;" & _
        "User ID=sa;" & _
        "Connect Timeout=2;" & _
        "Password=happyness;"
    Debug.Print "�������� ������� SQLAuditData SQLExecute. ������ ����������:" & vbCrLf & SQLRequestString
    SQLData.Open SQLRequestString, SQL, adOpenKeyset
        If SQLMode = laRX Then
            SQLExecute = SQLData.Fields(ParameterToRead).Value
        End If
    SQL.Close
    Exit Function

SQL_error:
Debug.Print "������ SQL " & Err.Number & ":" & vbCrLf & Err.Description
SQLExecute = Err.Number
End Function
