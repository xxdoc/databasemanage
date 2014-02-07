VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "数据库管理系统—客户查询端"
   ClientHeight    =   4272
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   7608
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4272
   ScaleWidth      =   7608
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "添加"
      Height          =   492
      Left            =   5400
      TabIndex        =   15
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修  改"
      Height          =   492
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   1212
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   5280
      TabIndex        =   9
      Text            =   "*"
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   1440
      TabIndex        =   7
      Text            =   "*"
      Top             =   1440
      Width           =   1812
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查  询"
      Height          =   492
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   1212
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   5280
      TabIndex        =   4
      Text            =   "*"
      Top             =   480
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   372
      Left            =   6480
      TabIndex        =   1
      Top             =   3600
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1440
      TabIndex        =   0
      Text            =   "*"
      Top             =   480
      Width           =   1812
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "要求输入明确账号和用户名"
      Height          =   372
      Left            =   5400
      TabIndex        =   16
      Top             =   3000
      Width           =   1212
   End
   Begin VB.Label Label8 
      Height          =   372
      Left            =   3240
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "PS：* 代表任意多个任意字符"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   2532
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "要求输入明确账号和用户名"
      Height          =   372
      Left            =   3120
      TabIndex        =   12
      Top             =   3000
      Width           =   1212
   End
   Begin VB.Label Label5 
      Caption         =   "无法修改"
      Height          =   372
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "修改日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3720
      TabIndex        =   8
      Top             =   1440
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "接入间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1296
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用 户 名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   1632
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "账  号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1308
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xlapp As Excel.Application
Public wkBook As Excel.Workbook
Public wkSheet As Excel.Worksheet

Private Sub Command1_Click()

wkBook.Save
wkBook.Close
xlapp.Quit
Shell "taskkill /im EXCEL.EXE /f", vbHide
End

End Sub

Private Sub Command2_Click()

Set xlapp1 = CreateObject("excel.application")
xlapp1.Application.EnableEvents = ture
Set wkBook1 = xlapp1.Workbooks.Open(App.Path & "\date\tmp.xls")
Set wksheet1 = wkBook1.Worksheets("sheet1")
xlapp1.DisplayAlerts = False




Dim Rows As Long
Rows = 0
Rows = wkSheet.UsedRange.Rows.Count

Dim a As Long
Dim b As Long
Dim c As Long
b = 2
For a = 2 To Rows
    If (wkSheet.Cells(a, 1) Like Text1.Text And wkSheet.Cells(a, 2) Like Text2.Text) Then
    If (wkSheet.Cells(a, 3) Like Text3.Text Or wkSheet.Cells(a, 7) Like Text3.Text) Then
    If wkSheet.Cells(a, 14) Like Text4.Text Then
        For c = 1 To 14
            wksheet1.Cells(b, c) = wkSheet.Cells(a, c)
        Next c
        b = b + 1
    End If
    End If
    End If
Next a


xlapp1.Visible = True

End Sub

Private Sub Command3_Click()





Dim Rows As Long
Rows = 0
Rows = wkSheet.UsedRange.Rows.Count
Dim d As Long
d = 0
For a = 2 To Rows
    If (wkSheet.Cells(a, 1) = Text1.Text And wkSheet.Cells(a, 2) = Text2.Text) Then
        d = 1
    End If
Next a

If d = 1 Then
wkBook.Save
wkBook.Close
xlapp.Quit

Form1.Visible = False
Form2.Visible = True
Else
MsgBox ("所输入的账号和用户名不精确")
End If

End Sub

Private Sub Command4_Click()
Dim Rows As Long
Rows = 0
Rows = wkSheet.UsedRange.Rows.Count
Dim d As Long
d = 0
For a = 2 To Rows
    If (wkSheet.Cells(a, 1) = Text1.Text And wkSheet.Cells(a, 2) = Text2.Text) Then
        d = 1
    End If
Next a

If d = 1 Then
wkBook.Save
wkBook.Close
xlapp.Quit

Form1.Visible = False
Form3.Visible = True
Else
MsgBox ("所输入的账号和用户名不精确")
End If
End Sub

Private Sub Form_Activate()

Set xlapp = CreateObject("excel.application")
xlapp.Application.EnableEvents = True
Set wkBook = xlapp.Workbooks.Open(App.Path & "\date\用户数据.xls")
Set wkSheet = wkBook.Worksheets("sheet1")
xlapp.DisplayAlerts = False
Dim Rows As Long
Rows = 0
Rows = wkSheet.UsedRange.Rows.Count

Label8.Caption = Rows




End Sub

Private Sub Form_Unload(Cancel As Integer)

wkBook.Save
wkBook.Close
xlapp.Quit
Shell "taskkill /im EXCEL.EXE /f", vbHide
End

End Sub
