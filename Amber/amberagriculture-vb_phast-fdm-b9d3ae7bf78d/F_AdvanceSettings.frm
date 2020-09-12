VERSION 5.00
Begin VB.Form F_AdvanceSettings 
   Caption         =   "F_AdvanceSettings"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Layer Settings"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Com_DrCost 
      Caption         =   "Drying Cost"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Com_FanBurnerSettings 
      Caption         =   "Fan-Burner Settings"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Com_DMLSettings 
      Caption         =   "DML Settings"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Com_GrainSettings 
      Caption         =   "Grain Settings"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "These Settings are for Advance Users"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "F_AdvanceSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()
F_AdvanceSettings.Hide
F_Main.Show
End Sub

Private Sub Com_DMLSettings_Click()
F_AdvanceSettings.Hide
F_DML.Show

End Sub

Private Sub Com_DrCost_Click()
F_AdvanceSettings.Hide
F_DryingCost.Show

End Sub

Private Sub Com_FanBurnerSettings_Click()
F_AdvanceSettings.Hide
F_FanBurner.Show

End Sub

Private Sub Com_GrainSettings_Click()
F_AdvanceSettings.Hide
F_GrainChoice.Show
' set array of initial grain temperature and MC at the average initial conditions for each layer
m = 0
For m = 0 To (NumberOfLayers - 1)
    GrainMC_In_WB_C(m) = Tb_AvgGrainMC
    GrainMC_In_WB_S(m) = Tb_AvgGrainMC
    GrainTemp_In_C_C(m) = Tb_AvgGrainTemp
    GrainTemp_In_C_S(m) = Tb_AvgGrainTemp
Next m

End Sub

Private Sub Command1_Click()
F_Layer.Show
F_AdvanceSettings.Hide

End Sub
