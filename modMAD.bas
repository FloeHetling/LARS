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

Public Function SQLGetAuditData(ByVal AuditProp As String, ByVal auditdata As String)

End Function
