Attribute VB_Name = "modStartup"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' ��������� ������ "�����" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'���������� ����������
    Public LARSver As String
    Public InfoBoxes() As String
      
'���������� ���������
''����������
    Public Const Lime = 12648384
    Public Const Sand = 12648447

'���������� �������
''���������� ����� AuditData, �������� ��� thisPC
    Public thisPC As New AuditData


Dim CLIArg As String

Sub Main()
'���������� � ���������� ���������� �������� � ������ ��
LARSver = App.ProductName & ", ������ " & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.CompanyName
'���������, ������� �� ������ ���������
'���� �� - ��������� ����� ������
    If App.PrevInstance = True Then
        Exit Sub
        End
    End If

'������� ������ ��������� �� ����� �������� � ���������� �� � ��������� ������
Dim Ctrl As Control
Dim ibIndex As Integer
Dim ibName As String
ibIndex = 0

    For Each Ctrl In frmWriteAuditData.Controls
        If InStr(1, Ctrl.Tag, "infobox") <> 0 Then
            ReDim Preserve InfoBoxes(ibIndex)
            ibName = Replace(Ctrl.Tag, "infobox,", "")
            InfoBoxes(ibIndex) = ibName
            ibIndex = ibIndex + 1
        End If
    Next

'���������� ��������� ���������� ������ � ���������� � ������ ��
CLIArg = Command$
    Select Case CLIArg
        
        Case "/edit"
        frmWriteAuditData.Show
        
        Case Else
        Call PopulateAuditData
                
    End Select
        
End Sub
