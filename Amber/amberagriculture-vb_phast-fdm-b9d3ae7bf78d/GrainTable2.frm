VERSION 5.00
Begin VB.Form F_GrainEdit 
   Caption         =   "F_GrainEdit"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   ScaleHeight     =   7935
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "This is a Corn Variety"
      Height          =   855
      Left            =   5040
      TabIndex        =   40
      Top             =   6000
      Width           =   2895
      Begin VB.OptionButton Opt_No 
         Caption         =   "No"
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Opt_Yes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   1440
      TabIndex        =   38
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   36
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Tb_AirfRes_B 
      Height          =   285
      Left            =   6360
      TabIndex        =   35
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Tb_AirfRes_A 
      Height          =   285
      Left            =   6360
      TabIndex        =   34
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Tb_SpHeat_B 
      Height          =   285
      Left            =   6360
      TabIndex        =   30
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Tb_SpHeat_A 
      Height          =   285
      Left            =   6360
      TabIndex        =   29
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Tb_Dens_C 
      Height          =   285
      Left            =   6360
      TabIndex        =   25
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Tb_Dens_B 
      Height          =   285
      Left            =   6360
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Tb_Dens_A 
      Height          =   285
      Left            =   6360
      TabIndex        =   23
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_R_C 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_R_B 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_R_A 
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_D_C 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_D_B 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Tb_EMC_D_A 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Tb_GrainName 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "Select Grain"
      Height          =   255
      Left            =   5040
      TabIndex        =   39
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   735
      Left            =   600
      TabIndex        =   37
      Top             =   7200
      Width           =   4455
   End
   Begin VB.Label Label21 
      Caption         =   "Par. b"
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Par. a"
      Height          =   255
      Left            =   5160
      TabIndex        =   32
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "Parameters for Airflow Resistance (Pa) (ASAE D272.3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   31
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label18 
      Caption         =   "Par. XM"
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Par. X"
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Parameter fo Specific Heat (kJ/(kg ºK) (ASAE D243.4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   26
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label15 
      Caption         =   "Par. X*M^2"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Par. X*M"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Par. X"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Parameters for Bulk Density (ASAE D241.4)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   19
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Re-wetting Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Drying Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "EMC Parameters (Modified Chung-Pfost)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label8 
      Caption         =   "Par. C"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Par. B"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Par. A"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Par. C"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Par. B"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Par. A"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Grain Name"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "F_GrainEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    GrainIndex = F_GrainEdit.Combo1.ListIndex
    Call DisplayGrainInfo(ArrGrain, GrainIndex)
End Sub

Private Sub Command1_Click()
'Dim arr() As Variant
'arr = ArrGrain
'commands to add a new grain
If F_GrainChoice.Option1 = True Then
    If ErrorF_GrainEdit() = True Then
        StringGrainInfo = Tb_GrainName + vbTab + Tb_EMC_D_A + vbTab + Tb_EMC_D_B + vbTab + Tb_EMC_D_C + vbTab + Tb_EMC_R_A + vbTab + Tb_EMC_R_B + vbTab + Tb_EMC_R_C + vbTab + Tb_Dens_A + vbTab + Tb_Dens_B + vbTab + Tb_Dens_C + vbTab + Tb_SpHeat_A + vbTab + Tb_SpHeat_B + vbTab + Tb_AirfRes_A + vbTab + Tb_AirfRes_B + vbTab + DMLFlag
        Open FileName For Append As #1
        Print #1, StringGrainInfo
        Close #1
     Else
        MsgBox ("Missing Values, Check Information and Try Again!!")
    End If
F_GrainEdit.Combo1.Clear
Call ReadGrainList(FileName)
   
End If
'commands to edit and existing grain
Dim LineToFile As String
LineToFile = ""
If F_GrainChoice.Option2 = True Then
    If ErrorF_GrainEdit() = True Then
        Open FileName For Output As #1
        For m = 0 To (LineIndex - 1)
            If m = GrainIndex Then
                ArrGrain(m, 0) = Tb_GrainName
                ArrGrain(m, 1) = Tb_EMC_D_A
                ArrGrain(m, 2) = Tb_EMC_D_B
                ArrGrain(m, 3) = Tb_EMC_D_C
                ArrGrain(m, 4) = Tb_EMC_R_A
                ArrGrain(m, 5) = Tb_EMC_R_B
                ArrGrain(m, 6) = Tb_EMC_R_C
                ArrGrain(m, 7) = Tb_Dens_A
                ArrGrain(m, 8) = Tb_Dens_B
                ArrGrain(m, 9) = Tb_Dens_C
                ArrGrain(m, 10) = Tb_SpHeat_A
                ArrGrain(m, 11) = Tb_SpHeat_B
                ArrGrain(m, 12) = Tb_AirfRes_A
                ArrGrain(m, 13) = Tb_AirfRes_B
                ArrGrain(m, 14) = DMLFlag
            End If
        
            For j = 0 To 14
                LineToFile = LineToFile + ArrGrain(m, j) + vbTab
            Next j
            Print #1, LineToFile
            LineToFile = ""
        Next m
    Close #1
    F_GrainEdit.Combo1.Clear
    Call ReadGrainList(FileName)
    Else
    MsgBox ("Missing Values, Check Information and Try Again!!")
    End If
End If

End Sub


Private Sub Command2_Click()
F_GrainEdit.Hide
F_GrainChoice.Show

End Sub

Private Sub Form_Activate()
' set the default value for DMLFlag to false
DMLFlag = 0
'set the grain list defoult selected value to "0"
If F_GrainChoice.Option1 = False Then
Combo1.ListIndex = 0
End If

End Sub

Private Sub Form_Load()
FileName = TablePath + "\graininfo.dry"
Call ReadGrainList(FileName)

Call UpdateF_GrainEdit
'set the grain list defoult selected value to "0"
If F_GrainChoice.Option1 = False Then
Combo1.ListIndex = 0
End If
End Sub

Private Sub Opt_No_Click()
'set dmlflag to compute dml for corn variety to "false"
DMLFlag = 0
End Sub

Private Sub Opt_Yes_Click()
'set dmlflag to compute dml for corn variety to "true"
DMLFlag = 1
End Sub
