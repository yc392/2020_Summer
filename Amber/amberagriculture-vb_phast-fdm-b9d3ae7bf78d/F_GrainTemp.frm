VERSION 5.00
Begin VB.Form F_GrainT_MC 
   Caption         =   "F_GrainT_MC"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   12
      Left            =   1320
      TabIndex        =   76
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   11
      Left            =   1320
      TabIndex        =   75
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Edit Center or Side of Bin"
      Height          =   1215
      Left            =   7440
      TabIndex        =   70
      Top             =   2520
      Width           =   2655
      Begin VB.OptionButton Op_Side 
         Caption         =   "Edit Side of Bin"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Op_Center 
         Caption         =   "Edit Center of Bin"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Edit Temperature or MC"
      Height          =   1215
      Left            =   7440
      TabIndex        =   67
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton Op_MC 
         Caption         =   "Edit MC"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Op_Temp 
         Caption         =   "Edit Temperature"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   64
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Com_Save 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   63
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   31
      Left            =   5880
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   30
      Left            =   5880
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   29
      Left            =   5880
      TabIndex        =   27
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   28
      Left            =   5880
      TabIndex        =   26
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   27
      Left            =   5880
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   26
      Left            =   5880
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   25
      Left            =   5880
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   24
      Left            =   5880
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   23
      Left            =   5880
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   22
      Left            =   5880
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   21
      Left            =   5880
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   20
      Left            =   5880
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   19
      Left            =   5880
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   18
      Left            =   5880
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   17
      Left            =   5880
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   16
      Left            =   5880
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   15
      Left            =   1320
      TabIndex        =   13
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   14
      Left            =   1320
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   13
      Left            =   1320
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   10
      Left            =   1320
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tb_Layer 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_SaveC_S 
      Caption         =   " at the Center of the Bin"
      Height          =   375
      Left            =   2400
      TabIndex        =   74
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Lb_FormHeading2 
      Caption         =   " at the Center of the Bin"
      Height          =   255
      Left            =   3960
      TabIndex        =   73
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Lb_SaveT_MC 
      Alignment       =   1  'Right Justify
      Caption         =   "Save MC"
      Height          =   255
      Left            =   960
      TabIndex        =   66
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label33 
      Caption         =   "Return to Previous Window"
      Height          =   375
      Left            =   5160
      TabIndex        =   65
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Lb_FormHeading1 
      Alignment       =   1  'Right Justify
      Caption         =   "Edit MC"
      Height          =   255
      Left            =   2280
      TabIndex        =   62
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 32"
      Height          =   255
      Index           =   31
      Left            =   4920
      TabIndex        =   61
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 31"
      Height          =   255
      Index           =   30
      Left            =   4920
      TabIndex        =   60
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 30"
      Height          =   255
      Index           =   29
      Left            =   4920
      TabIndex        =   59
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 29"
      Height          =   255
      Index           =   28
      Left            =   4920
      TabIndex        =   58
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 28"
      Height          =   255
      Index           =   27
      Left            =   4920
      TabIndex        =   57
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 27"
      Height          =   255
      Index           =   26
      Left            =   4920
      TabIndex        =   56
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 26"
      Height          =   255
      Index           =   25
      Left            =   4920
      TabIndex        =   55
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 25"
      Height          =   255
      Index           =   24
      Left            =   4920
      TabIndex        =   54
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 24"
      Height          =   255
      Index           =   23
      Left            =   4920
      TabIndex        =   53
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 23"
      Height          =   255
      Index           =   22
      Left            =   4920
      TabIndex        =   52
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 22"
      Height          =   255
      Index           =   21
      Left            =   4920
      TabIndex        =   51
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 21"
      Height          =   255
      Index           =   20
      Left            =   4920
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 20"
      Height          =   255
      Index           =   19
      Left            =   4920
      TabIndex        =   49
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 19"
      Height          =   255
      Index           =   18
      Left            =   4920
      TabIndex        =   48
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 18"
      Height          =   255
      Index           =   17
      Left            =   4920
      TabIndex        =   47
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 17"
      Height          =   255
      Index           =   16
      Left            =   4920
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 16"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   45
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 15"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   44
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 14"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   43
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 13"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   42
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 12"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 11"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   40
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 10"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   39
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 9"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   38
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 8"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 7"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 6"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   35
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 5"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   34
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 4"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 3"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 2"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb_Layer 
      Caption         =   "Layer 1"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "F_GrainT_MC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()
F_GrainT_MC.Hide
F_GrainChoice.Show
m = 0
For m = 0 To (NumberOfLayers - 1)
    Tb_Layer(m).Visible = False
    Lb_Layer(m).Visible = False
Next m
End Sub

Private Sub Com_Save_Click()
' determine which text boxes should be visible
m = 0
For m = 0 To (NumberOfLayers - 1)
    Tb_Layer(m).Visible = True
    Tb_Layer(m).Text = GrainMC_In_WB_C(m)
    Lb_Layer(m).Visible = True
Next m
        
' assign the text box information to the corresponding array of temp or MC, center or side
If Op_MC = True Then
    If Op_Center = True Then
        m = 0
        For m = 0 To (NumberOfLayers - 1)
            GrainMC_In_WB_C(m) = Tb_Layer(m)
        Next m
    End If
    If Op_Side = True Then
        m = 0
        For m = 0 To (NumberOfLayers - 1)
            GrainMC_In_WB_S(m) = Tb_Layer(m)
        Next m
    End If
End If
If Op_Temp = True Then
    If Op_Center = True Then
        m = 0
        For m = 0 To (NumberOfLayers - 1)
            GrainTemp_In_C_C(m) = Tb_Layer(m)
        Next m
    End If
    If Op_Side = True Then
        m = 0
        For m = 0 To (NumberOfLayers - 1)
            GrainTemp_In_C_S(m) = Tb_Layer(m)
        Next m
    End If
End If

End Sub

Private Sub Form_Activate()
' determine which text boxes should be visible
m = 0
For m = 0 To (NumberOfLayers - 1)
    Tb_Layer(m).Visible = True
    Tb_Layer(m).Text = GrainMC_In_WB_C(m)
    Lb_Layer(m).Visible = True
Next m

End Sub

Private Sub Form_Load()
' visualize only the tex box corresponding to the layers available in the simulation


' determine which text boxes should be visible
m = 0
For m = 0 To (NumberOfLayers - 1)
    Tb_Layer(m).Visible = True
    Tb_Layer(m).Text = GrainMC_In_WB_C(m)
    Lb_Layer(m).Visible = True
Next m





End Sub


Private Sub Op_Center_Click()
Lb_FormHeading2.Caption = " at the Center of the Bin"
Lb_SaveC_S.Caption = " at the Center of the Bin"
' update the current information in the text box
If Op_MC = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainMC_In_WB_C(m)
    Next m
End If
If Op_Temp = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainTemp_In_C_C(m)
    Next m
End If

End Sub
Private Sub Op_side_Click()
Lb_FormHeading2.Caption = " at the Side of the Bin"
Lb_SaveC_S.Caption = " at the Side of the Bin"
' update the current information in the text box
If Op_MC = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainMC_In_WB_S(m)
    Next m
End If
If Op_Temp = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainTemp_In_C_S(m)
    Next m
End If

End Sub

Private Sub Op_MC_Click()
Lb_FormHeading1.Caption = "Edit MC"
Lb_SaveT_MC.Caption = "Save MC"

' update the current information in the text box
If Op_Center = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainMC_In_WB_C(m)
    Next m
End If
If Op_Side = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainMC_In_WB_S(m)
    Next m
End If

End Sub
Private Sub Op_temp_Click()
Lb_FormHeading1.Caption = "Edit Temperature"
Lb_SaveT_MC.Caption = "Save Temperature"

' update the current information in the text box
If Op_Center = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainTemp_In_C_C(m)
    Next m
End If
If Op_Side = True Then
    m = 0
    For m = 0 To (NumberOfLayers - 1)
        Tb_Layer(m).Text = GrainTemp_In_C_S(m)
    Next m
End If

End Sub

