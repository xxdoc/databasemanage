VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "数据库管理系统—客户添加端"
   ClientHeight    =   5532
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   7536
   Enabled         =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5532
   ScaleWidth      =   7536
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   1812
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   2052
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   1680
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1812
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   5160
      TabIndex        =   5
      Top             =   2400
      Width           =   2052
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1812
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添  加"
      Height          =   372
      Left            =   1920
      TabIndex        =   2
      Top             =   4800
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重  置"
      Height          =   372
      Left            =   4800
      TabIndex        =   1
      Top             =   4800
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   3840
      Width           =   5532
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "修改日期："
      Height          =   180
      Left            =   3840
      TabIndex        =   24
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "新 接入间："
      Height          =   180
      Left            =   360
      TabIndex        =   23
      Top             =   1080
      Width           =   996
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用 户 名 ："
      Height          =   180
      Left            =   3840
      TabIndex        =   22
      Top             =   360
      Width           =   1008
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "账   号 ："
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   360
      Width           =   924
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1440
      TabIndex        =   20
      Top             =   360
      Width           =   96
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   19
      Top             =   360
      Width           =   96
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4800
      TabIndex        =   18
      Top             =   1080
      Width           =   96
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "新 Vlan 号："
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   1116
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "对应设备端口："
      Height          =   180
      Left            =   3840
      TabIndex        =   16
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "修 改 原 因："
      Height          =   180
      Left            =   3840
      TabIndex        =   15
      Top             =   3240
      Width           =   1188
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "修改人确认人："
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "录 入 人："
      Height          =   180
      Left            =   3840
      TabIndex        =   13
      Top             =   2520
      Width           =   912
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "客户联系电话："
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "新装机地址:"
      Height          =   180
      Left            =   360
      TabIndex        =   11
      Top             =   3960
      Width           =   996
   End
   Begin VB.Label Label15 
      Height          =   372
      Left            =   3480
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xlapp As Excel.Application
Public wkBook As Excel.Workbook
Public wkSheet As Excel.Worksheet

Private Sub Command1_Click()



Dim b As Long
b = Label15.Caption
wkSheet.Cells(b, 1) = Label7.Caption
wkSheet.Cells(b, 2) = Label8.Caption
wkSheet.Cells(b, 14) = Label9.Caption
wkSheet.Cells(b, 8) = Text1.Text
wkSheet.Cells(b, 9) = Text2.Text
wkSheet.Cells(b, 13) = Text3.Text
wkSheet.Cells(b, 12) = Text4.Text
wkSheet.Cells(b, 15) = Text5.Text
wkSheet.Cells(b, 11) = Text6.Text
wkSheet.Cells(b, 10) = Text7.Text

wkBook.Save
wkBook.Close
xlapp.Quit

Form3.Visible = False
Form1.Visible = True
End Sub

Private Sub Command2_Click()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label9.Caption = Now()

End Sub

Private Sub Form_Activate()
Set xlapp = CreateObject("excel.application")
xlapp.Application.EnableEvents = True
Set wkBook = xlapp.Workbooks.Open(App.Path & "\date\用户数据.xls")
Set wkSheet = wkBook.Worksheets("sheet1")
xlapp.DisplayAlerts = False

Dim Rows As Long
Dim b As Long
b = 0
Rows = 0
Rows = wkSheet.UsedRange.Rows.Count
b = Rows + 1

Label15.Caption = b
Label7.Caption = Form1.Text1.Text
Label8.Caption = Form1.Text2.Text
Label9.Caption = Now()



End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Visible = False
Form1.Visible = True

End Sub


