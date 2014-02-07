VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   2244
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   3744
   LinkTopic       =   "Form4"
   ScaleHeight     =   2244
   ScaleWidth      =   3744
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "登  录"
      Height          =   372
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1332
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   1452
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密  码："
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   732
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作人员姓名："
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
