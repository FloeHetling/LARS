VERSION 5.00
Begin VB.Form frmWriteAuditData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ЛАРС: Редактор"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9525
   Icon            =   "frmWriteAuditData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   355
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tDebugPrint 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   59
      Top             =   5400
      Width           =   9375
   End
   Begin VB.ComboBox OffOLP 
      Height          =   315
      Left            =   2040
      TabIndex        =   58
      Tag             =   "infobox,OfficeLicenseModel"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.ComboBox WinOLP 
      Height          =   315
      Left            =   2040
      TabIndex        =   57
      Tag             =   "infobox,WindowsLicenseModel"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox SQLCheck 
      BorderStyle     =   0  'Нет
      Height          =   4335
      Left            =   4800
      Picture         =   "frmWriteAuditData.frx":27A2
      ScaleHeight     =   4335
      ScaleWidth      =   4575
      TabIndex        =   56
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox SqlBlock 
      BorderStyle     =   0  'Нет
      Height          =   4335
      Left            =   4800
      Picture         =   "frmWriteAuditData.frx":41B8E
      ScaleHeight     =   4335
      ScaleWidth      =   4575
      TabIndex        =   55
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame frInfo 
      Caption         =   "Информация из базы"
      Height          =   4695
      Index           =   1
      Left            =   4680
      TabIndex        =   24
      Top             =   120
      Width           =   4815
      Begin VB.Frame frWindows 
         Caption         =   "Windows"
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   4575
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   5
            Left            =   4080
            TabIndex        =   51
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   4
            Left            =   4080
            TabIndex        =   50
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   3
            Left            =   4080
            TabIndex        =   49
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   42
            Tag             =   "SQLbox,WindowsOLPSerial"
            Text            =   "Text6"
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   4
            Left            =   1800
            TabIndex        =   41
            Tag             =   "SQLbox,WindowsLicenseModel"
            Text            =   "Text5"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   3
            Left            =   1800
            TabIndex        =   40
            Tag             =   "SQLbox,WindowsVersion"
            Text            =   "Text4"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Windows:"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер OLP:"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame frOffice 
         Caption         =   "Office"
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   4575
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   8
            Left            =   4080
            TabIndex        =   54
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   7
            Left            =   4080
            TabIndex        =   53
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   6
            Left            =   4080
            TabIndex        =   52
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   8
            Left            =   1800
            TabIndex        =   35
            Tag             =   "SQLbox,OfficeOLPSerial"
            Text            =   "Text9"
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   34
            Tag             =   "SQLbox,OfficeLicenseModel"
            Text            =   "Text8"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   33
            Tag             =   "SQLbox,OfficeVersion"
            Text            =   "Text7"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Office:"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер OLP:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame frCommon 
         Caption         =   "Общая"
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   2
            Left            =   4080
            TabIndex        =   48
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   1
            Left            =   4080
            TabIndex        =   47
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "C"
            Height          =   315
            Index           =   0
            Left            =   4080
            TabIndex        =   46
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   28
            Tag             =   "SQLbox,WSSerial"
            Text            =   "Text3"
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   27
            Tag             =   "SQLbox,WSName"
            Text            =   "Text2"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox sqlinfo 
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   26
            Tag             =   "SQLbox,Company"
            Text            =   "Text1"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Организация:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Имя ПК:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер с наклейки:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
      End
   End
   Begin VB.Frame frInfo 
      Caption         =   "Информация из реестра"
      Height          =   4695
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.Frame frCommon 
         Caption         =   "Общая"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4215
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   20
            Tag             =   "infobox,Company"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   19
            Tag             =   "infobox,WSName"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   18
            Tag             =   "infobox,WSSerial"
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер с наклейки:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Имя ПК:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Организация:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frOffice 
         Caption         =   "Office"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   4215
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   13
            Tag             =   "infobox,OfficeVersion"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   8
            Left            =   1800
            TabIndex        =   12
            Tag             =   "infobox,OfficeOLPSerial"
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер OLP:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Office:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frWindows 
         Caption         =   "Windows"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   4215
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   3
            Left            =   1800
            TabIndex        =   7
            Tag             =   "infobox,WindowsVersion"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox cbinfo 
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   6
            Tag             =   "infobox,WindowsOLPSerial"
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "Номер OLP:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Windows:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox container 
      Height          =   375
      Index           =   1
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   8115
      TabIndex        =   2
      Top             =   4920
      Width           =   8175
      Begin VB.Label stDescription 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   45
         Width           =   6255
      End
   End
   Begin VB.PictureBox container 
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
      Begin VB.Label stTitle 
         Caption         =   "Готов"
         Height          =   255
         Left            =   105
         TabIndex        =   1
         Top             =   45
         Width           =   975
      End
   End
   Begin VB.Timer tDelayedReadData 
      Interval        =   20
      Left            =   8760
      Top             =   8760
   End
   Begin VB.Timer tResetColor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8400
      Top             =   8760
   End
   Begin VB.Menu cbCommands 
      Caption         =   "&Команды"
      Begin VB.Menu cmdWriteToRegistry 
         Caption         =   "Записать в реестр"
         Shortcut        =   ^S
      End
      Begin VB.Menu delim1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdRefreshFromRegistry 
         Caption         =   "Обновить из реестра"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu cmdRefreshFromSQL 
         Caption         =   "Обновить из базы"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu delim2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdExit 
         Caption         =   "Выход"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu cmdTools 
      Caption         =   "&Инструменты"
      Begin VB.Menu cmdCheckSQL 
         Caption         =   "Проверить соединение с базой"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdPopulateAuditData 
         Caption         =   "Выполнить Аудитор"
      End
      Begin VB.Menu delim3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdLaunchCLI 
         Caption         =   "Запустить командную строку"
         Shortcut        =   {F3}
      End
      Begin VB.Menu cmdWMIQL 
         Caption         =   "Запрос к WMI"
         Shortcut        =   {F12}
      End
      Begin VB.Menu delim4 
         Caption         =   "-"
      End
      Begin VB.Menu cmdReport 
         Caption         =   "Сообщить о различиях"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmWriteAuditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода
'debug.print<--> debug.print
Dim ctlInfobox As Control, SQLInfobox As Control
Dim isDataChanged As Boolean, isRunOnceDataLoaded As Boolean
Private Enum laLoadDataMode
    laLoadFromRegistry = 1
    laLoadFromSQL = 2
End Enum

Private Function LoadAuditData(Optional LoadFrom As laLoadDataMode)
Dim ctlIBValue As String, cbAuditValue As String, cbAuditValueSQL As String
tResetColor.Enabled = True
enumSQLFields = UBound(InfoBoxes) - LBound(InfoBoxes) + 1
'Заполняем классы
Status "Работаю", "Загружаю информацию из реестра", laDarkBlue
thisPC.RegLoad

If isSQLAvailable = True Then     'Здесь обращаемся к имени ПК. Оно взято в переменную
    Status "Работаю", "Загружаю информацию из SQL", laDarkBlue
    thisPCSQL.SQLLoad (HostName)    'в начальном модуле modStartup (Sub Main). Имя ПК передается методу класса SQLAuditData
    Debug.Print "Загружена информация из SQL"
End If
frmWriteAuditData.tDebugPrint.Text = ""
Debug.Print "Грузим значения в ячейки"

    For Each ctlInfobox In Me.Controls                              'Теперь выполняем для каждого инфополя из сформированного в sub_main массива
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then            'если в списке тегов есть тег инфобокса
            Dim InfoboxTag() As String
            InfoboxTag = Split(ctlInfobox.Tag, ",")
            ctlIBValue = InfoboxTag(1)                              'вычленяем из тега имя параметра
            cbAuditValue = CallByName(thisPC, ctlIBValue, VbGet)    'и вызовом класса AuditData получаем значение в этот параметр
                
                ' "infobox,"
                '
                ' защита от дубликатов в списке
                ' процедура модуля MAD cbExists проверяет, есть ли этот элемент в комбобоксе
                ' если есть - не добавляет
                '
'                If cbExists(cbAuditValue, ctlInfobox) = False Then
                    With ctlInfobox
                     .Text = cbAuditValue
                     .BackColor = laSand
'                     .ListIndex = 0
                    End With
'                End If
        End If
    Next
    
Debug.Print "Готово"
Debug.Print "Грузим значения в ячейки SQL"
    
    If isSQLAvailable = True Then 'делаем это только если стоит флажок "Сравнить с SQL"
        For Each SQLInfobox In Me.Controls
            If InStr(1, SQLInfobox.Tag, "SQLbox") <> 0 Then
                SQLInfobox.Enabled = False
                Dim SQLBoxTag() As String
                SQLBoxTag = Split(SQLInfobox.Tag, ",")
                ctlIBValue = SQLBoxTag(1)
                cbAuditValueSQL = CallByName(thisPCSQL, ctlIBValue, VbGet)
                        If Not cbAuditValueSQL = "sql_err_nodata" Then
                            SQLInfobox.Text = cbAuditValueSQL
                        Else
                            SQLInfobox.Text = ""
                        End If
            End If
        Next
    End If

Debug.Print "Готово"

isRunOnceDataLoaded = True
isDataChanged = False
    '' отлов ошибки с пустым ответом сервера
    '' и предложение внести в базу данные с форнмы
'    If enumSQLFields = 0 Then
'            If MsgBox("В БД не найдено никаких сведений о " & HostName & "!" & vbCrLf & _
'                        "Желаете добавить сведения с текущей формы как новую запись в БД?", _
'                        vbQuestion & vbYesNo, LARSver) = vbYes Then
'
'                        'проверка пустых полей
'                            Dim cbiCount As Integer, NullFieldWarning As Boolean
'                            For cbiCount = 0 To cbinfo().UBound
'                                If cbinfo(cbiCount).Text = "Нет данных" Or _
'                                cbinfo(cbiCount).Text = "" Then _
'                                NullFieldWarning = True Else _
'                                NullFieldWarning = False
'                            Next
'                        If NullFieldWarning = True Then
'                            If MsgBox("Одно или несколько полей на форме не заполнены. Продолжить?", _
'                            vbQuestion & vbYesNo, LARSver) = vbYes Then _
'                            Call SaveAuditData(laWriteToSQL)
'                        Else
'                            Call SaveAuditData(laWriteToSQL)
'                        End If
'            End If
'    End If
End Function

Private Function SaveAuditData(ByVal WriteMode As laWriteMode)
Dim ctlIBVariable As String
Dim ctlIBValue As String
tResetColor.Enabled = True
    For Each ctlInfobox In Me.Controls
        '
        'для всех элементов формы с тегом infobox
        'мы конвертим тег в свойство класса AuditData
        'затем, вызываем класс по имени экземпляра
        'и помещаем в его переменные соответствующую инфу из
        'инфобокса, который в данный момент участвует в цикле
        '
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            Dim InfoboxTag() As String
            InfoboxTag = Split(ctlInfobox.Tag, ",")
            ctlIBVariable = InfoboxTag(1)
            ctlInfobox.BackColor = laLightGreen
            ctlIBValue = ctlInfobox.Text
          '
          ' ctlIBValue = ctlInfobox.List(ctlInfobox.ListIndex) Этого здесь нахрен не надо
          '
            If WriteMode = laWriteToRegistry Then CallByName thisPC, ctlIBVariable, VbLet, ctlIBValue
            If WriteMode = laWriteToSQL Then CallByName thisPCSQL, ctlIBVariable, VbLet, ctlIBValue
        End If
    Next
    
    'обработав все элементы с тегом infobox и заполнив все переменные класса AuditData
    'запускаем внутреннюю процедуру класса, записывающую данные в реестр Windo
        Select Case WriteMode
            Case laWriteEverywhere
                thisPC.RegSave
                thisPCSQL.SQLSave (HostName)
            Case laWriteToRegistry
                thisPC.RegSave
            Case laWriteToSQL
                thisPCSQL.SQLSave (HostName)
        End Select

isDataChanged = False
End Function

Private Sub cbinfo_Change(Index As Integer)
isDataChanged = True
End Sub

Private Sub cbInfo_Click(Index As Integer)
If cbinfo(Index).BackColor = laLightRed Then
    With cbinfo(Index)
    .BackColor = vbWhite
    .Tag = Replace(.Tag, ",noreset", "")
    End With
End If
End Sub

'Private Sub cbinfo_KeyPress(index As Integer, KeyAscii As Integer)
'KeyAscii = AutoMatchCBBox(cbinfo(index), KeyAscii)
'End Sub

Private Sub cmdLoad_Click()

End Sub

Private Sub cmdSubmit_Click()
    
End Sub

Private Sub cbinfo_KeyPress(Index As Integer, KeyAscii As Integer)
isDataChanged = True
End Sub

Private Sub chkSQL_Click()
' debug.printSQLExecute("SELECT * FROM dbo.larspc", laRX) должно быть не равно -2147467259
End Sub

Private Sub cmdLaunchAIDA_Click()
Shell "\\zdc5\work\Administrator\AIDA\aida64.exe", vbNormalFocus
End Sub

Private Sub cmdCopy_Click(Index As Integer)
If Index = 4 And sqlinfo(Index).Text <> "" Then WinOLP.Text = sqlinfo(Index).Text
If Index = 7 And sqlinfo(Index).Text <> "" Then OffOLP.Text = sqlinfo(Index).Text
If sqlinfo(Index).Text <> "" Then cbinfo(Index).Text = sqlinfo(Index).Text
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLaunchCLI_Click()
Shell "cmd.exe", vbNormalFocus
End Sub

Private Sub cmdPopulateAuditData_Click()
Call PopulateAuditData
End Sub

Private Sub cmdReport_Click()
frmReport.Show
End Sub

Private Sub cmdWMIQL_Click()
frmWMIQL.Show
End Sub


Private Sub cmdWriteToRegistry_Click()
Status "Занят", "Запись в реестр Windows", laDarkGreen
'Dim cbInfoCount As Integer
'cbInfoCount = 0
'        For cbInfoCount = 0 To cbinfo().UBound
'            cbinfo(cbInfoCount).Enabled = False
'        Next
Call SaveAuditData(laWriteToRegistry)
'cbInfoCount = 0
'        For cbInfoCount = 0 To cbinfo().UBound
'            cbinfo(cbInfoCount).Enabled = True
'        Next
Status "Готов", "Синхронизация завершена", laBlack
End Sub

Private Sub Form_Load()
SQLCheck.Visible = True
SqlBlock.Visible = False
isRunOnceDataLoaded = False

With WinOLP
    .AddItem "BOX"
    .AddItem "OLP"
    .AddItem "OEM"
    
End With

With OffOLP
    .AddItem "BOX"
    .AddItem "OLP"
    .AddItem "360"
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If isDataChanged = True Then
    If MsgBox("Есть несохраненные изменения реестра" & vbCrLf & "Вы точно хотите выйти?", vbQuestion & vbYesNo, LARSver) = vbYes Then
        End
    Else
        Cancel = 1
    End If
End If
'If isSQLSyncCompleted = False Then
'    if msgbox("Хотите актуализировать записи в SQL по этому ПК?")
'
''' Это вообще не приоритет...

End Sub

Private Sub tDelayedReadData_Timer()
If CLIArg <> "" Then
    Call LoadAuditData
    tDelayedReadData.Enabled = False
    Status "Готов", "Загружены данные из реестра Windows", laBlack
    Debug.Print "Проверка SQL в отложенной записи данных"
    'Проверка доступности SQL
        If isSQLAvailable = False Then
        SQLCheck.Visible = False
        SqlBlock.Visible = True
        Else
        SQLCheck.Visible = False
        SqlBlock.Visible = False
        End If
    Debug.Print "Завершена"
End If
End Sub

Private Sub tDelayedWriteData_Timer()

End Sub

'Private Sub tDelayedWriteData_Timer()
'Select Case chkSQL.value
'        Case 0
'            Call SaveAuditData(laWriteToRegistry)
'        Case 1
'            Call SaveAuditData(laWriteToSQL)
'End Select
'tDelayedWriteData.Enabled = False
'cmdSync.Enabled = True
'
'End Sub

Private Sub tResetColor_Timer()
Dim ibColor As Integer
    For Each ctlInfobox In Me.Controls
    If (InStr(1, ctlInfobox.Tag, "infobox") <> 0) And Not (InStr(1, ctlInfobox.Tag, "noreset") <> 0) Then ctlInfobox.BackColor = vbWhite
    Next
tResetColor.Enabled = False
Status "Готов", "", laBlack
End Sub

Public Function Status(Optional ByVal StatusText As String, _
                        Optional ByVal StatusDescription As String, _
                        Optional ByRef StatusColor As laColorConstants)
With stTitle
    .Caption = StatusText
    .ForeColor = StatusColor
End With

With stDescription
    .Caption = StatusDescription
    .ForeColor = StatusColor
End With
End Function


