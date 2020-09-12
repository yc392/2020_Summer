VERSION 5.00
Begin VB.Form F_Layer 
   Caption         =   "F_Layer"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Co_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Co_Update 
      Caption         =   "Update"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Tb_LayerDepth 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Visualize and Change Layer Information"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Lb_NoLayers 
      Caption         =   "Number of Layers:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Lb_BinHeight 
      Caption         =   "Bin Height (m):"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Layer Depth (m):"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "F_Layer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Co_Close_Click()
F_Layer.Hide
F_AdvanceSettings.Show

End Sub

Private Sub Co_Update_Click()
SingleLayerDepth1 = Tb_LayerDepth
SingleLayerDepth = SingleLayerDepth1 - 0.1
NumberOfLayers = 32
Do Until NumberOfLayers <= 31
    SingleLayerDepth = SingleLayerDepth + 0.1
    If BinHeight - Int(BinHeight) < 0.25 Then
       NumberOfLayers = Int(BinHeight / SingleLayerDepth)
    Else
       NumberOfLayers = Int(BinHeight / SingleLayerDepth) + 1
    End If
Loop
Lb_NoLayers.Caption = "Number of Layers: " & NumberOfLayers


End Sub

Private Sub Form_Load()
Tb_LayerDepth = SingleLayerDepth
Lb_BinHeight.Caption = "Bin Height: " & BinHeight & " m"
Lb_NoLayers.Caption = "Number of Layers: " & NumberOfLayers

End Sub
