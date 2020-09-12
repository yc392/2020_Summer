VERSION 5.00
Begin VB.Form F_Error 
   Caption         =   "F_Error"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Lb_ErrorMsg 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "F_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
F_Error.Hide
End Sub
