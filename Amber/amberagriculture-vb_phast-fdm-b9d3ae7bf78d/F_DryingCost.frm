VERSION 5.00
Begin VB.Form F_DryingCost 
   Caption         =   "F_DryingCost"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frm_HeaterType 
      Caption         =   "Chose the Heater Type"
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "Electrical"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Liquid Propane"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Tb_DesFinalMC 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Tb_PropCost 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Tb_ElectCost 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Tb_GrainCost 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Lb_GrainCost 
      Caption         =   "$/bu"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Desired Final MC (%, W.B.):"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Propane Cost ($/gallon):"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Electricity Cost ($/kWh):"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Grain Cost ($/tonne):"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "F_DryingCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
GrainPrice = Tb_GrainCost
ElectricityCost = Tb_ElectCost
PropaneCost = Tb_PropCost
DesiredFinMC = Tb_DesFinalMC
If Option1 = True Then
    HeaterType = True
Else
    HeaterType = False
End If

F_DryingCost.Hide
F_AdvanceSettings.Show
End Sub

Private Sub Form_Load()
Tb_GrainCost = GrainPrice
Tb_ElectCost = ElectricityCost
Tb_PropCost = PropaneCost
Tb_DesFinalMC = DesiredFinMC

If HeaterType = True Then
    Option1 = True
    Option2 = False
Else
    Option1 = False
    Option2 = True
End If
End Sub

Private Sub Tb_GrainCost_Change()
Lb_GrainCost.Caption = Tb_GrainCost / 40 & " $/bu"
End Sub
