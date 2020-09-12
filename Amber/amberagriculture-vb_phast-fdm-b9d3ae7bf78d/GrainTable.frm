VERSION 5.00
Begin VB.Form F_GrainChoice 
   Caption         =   "F_GrainChoice"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Com_EditLayer 
      Caption         =   "Edit Layer"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close Window"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Edit Existing Grain"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Add New Grain"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Visualize Grain Parameters"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Com_EditGrain 
      Caption         =   "Add"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Tb_TablePath 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "N:\Programs\VB_PHAST-FDM"
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Edit Temperature and MC by Layer"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Path for the Grain Table file"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "F_GrainChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_EditGrain_Click()

F_GrainChoice.Hide
Call UpdateF_GrainEdit
F_GrainEdit.Show
End Sub

Private Sub Com_EditLayer_Click()
F_GrainChoice.Hide
F_GrainT_MC.Show
End Sub

Private Sub Command2_Click()
F_GrainChoice.Hide
F_AdvanceSettings.Show
End Sub

Private Sub Form_Load()
TablePath = Tb_TablePath
End Sub

Private Sub Option1_Click()
' change the caption of command buttom to "add"
Com_EditGrain.Caption = "Add"
End Sub

Private Sub Option2_Click()
' change the caption of command buttom to "add"
Com_EditGrain.Caption = "Edit"
End Sub

Private Sub Option3_Click()
' change the caption of command buttom to "add"
Com_EditGrain.Caption = "Visualize"
End Sub


Private Sub Tb_TablePath_Change()
TablePath = Tb_TablePath
End Sub
