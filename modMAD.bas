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

Public Function SQLGetAuditData(ByVal AuditProp As String, ByVal auditdata As String)

End Function
