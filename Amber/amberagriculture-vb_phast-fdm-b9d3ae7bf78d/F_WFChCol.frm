VERSION 5.00
Begin VB.Form F_WFChCol 
   Caption         =   "F_WFChCol"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tb_TempColumn 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "5"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Tb_RHColumn 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "6"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Locations of the temperature and RH colunms in the weather file"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Temperature Column #"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "RH Column #"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "F_WFChCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()

TempColumn = Tb_TempColumn
RHColumn = Tb_RHColumn

F_WFChCol.Hide
F_WFile.Show
End Sub


