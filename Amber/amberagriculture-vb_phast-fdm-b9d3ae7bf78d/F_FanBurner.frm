VERSION 5.00
Begin VB.Form F_FanBurner 
   Caption         =   "F_FanBurner"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tb_BurnerEff 
      Height          =   285
      Left            =   8040
      TabIndex        =   51
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Tb_BurnerTempInc 
      Height          =   285
      Left            =   8040
      TabIndex        =   43
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Tb_FanEfficiency 
      Height          =   285
      Left            =   1680
      TabIndex        =   35
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Tb_PackingF 
      Height          =   285
      Left            =   1680
      TabIndex        =   32
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Com_Update 
      Caption         =   "Update Settings"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Tb_AirF_NUF 
      Height          =   285
      Left            =   8040
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Tb_FanPreWarm 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Tb_AirfRes 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label36 
      Caption         =   "BTU"
      Height          =   255
      Left            =   8760
      TabIndex        =   57
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Lb_BurnerEnergy_BTU 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   56
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label34 
      Caption         =   "KWH"
      Height          =   255
      Left            =   8760
      TabIndex        =   55
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Lb_BurnerEnergy_KWH 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label33 
      Caption         =   "Burner Energy Consumption:"
      Height          =   255
      Left            =   5880
      TabIndex        =   53
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label31 
      Caption         =   "%"
      Height          =   255
      Left            =   8760
      TabIndex        =   52
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "Burner Efficiency:"
      Height          =   255
      Left            =   6720
      TabIndex        =   50
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label32 
      Caption         =   "BTU"
      Height          =   255
      Left            =   8760
      TabIndex        =   49
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Lb_HeatingEnergy_BTU 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   48
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "KWH"
      Height          =   255
      Left            =   8760
      TabIndex        =   47
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Lb_HeatingEnergy_KW 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   46
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "Estimated Required Power:"
      Height          =   255
      Left            =   6000
      TabIndex        =   45
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label27 
      Caption         =   "ºC"
      Height          =   255
      Left            =   8760
      TabIndex        =   44
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "Desired Temperature Increase:"
      Height          =   255
      Left            =   5760
      TabIndex        =   42
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label25 
      Caption         =   "Burner Settings"
      Height          =   255
      Left            =   6360
      TabIndex        =   41
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label24 
      Caption         =   "%"
      Height          =   255
      Left            =   2520
      TabIndex        =   40
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Lb_FanPower_KW 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   1680
      TabIndex        =   39
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "HP"
      Height          =   255
      Left            =   2520
      TabIndex        =   38
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Lb_FanPower_HP 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   1680
      TabIndex        =   37
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "KW"
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Fan Efficiency:"
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "(for the Shedd's curves)"
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Packing Factor:"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "ºF"
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Lb_FanPreWarm_F 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "ºC"
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Inch Water"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Lb_AirfRes_Wa 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Pa"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "cfm/bu"
      Height          =   255
      Left            =   8520
      TabIndex        =   24
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Lb_AirflowSide_cfm 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   23
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "cfm/bu"
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Lb_AirflowCenter_cfm 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "m3/min/t"
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "m3/min/t"
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Lb_AirflowSide 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Lb_AirflowCenter 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8040
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Airflow at the Side:"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Aiflow at the Center:"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "cfm"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Lb_TotalAirflow_cfm 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "m3/min"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Airflow Non-Uniformity Factor:"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Fan Power:"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Fan Prewarming:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Airflow Resistance:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Lb_TotalAirflow 
      Caption         =   "xxxx"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Total Airflow:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fan and Burner Advance Settings"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "F_FanBurner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()
F_FanBurner.Hide
F_AdvanceSettings.Show
End Sub

Private Sub Com_Update_Click()
'update change in airflow non uniformity factor
AirflowNonUnif = Tb_AirF_NUF / 2
'set the airflow rate at the center and side of the bin according to the non-uniformity factor
AirflowCenter = Int(AirflowRate * (1 - AirflowNonUnif) * 100) / 100
AirflowSide = Int(AirflowRate * (1 + AirflowNonUnif) * 100) / 100
' print information about airflow in the screen
Lb_AirflowCenter.Caption = AirflowCenter
Lb_AirflowCenter_cfm.Caption = Int(AirflowCenter / 1.11 * 100) / 100
Lb_AirflowSide.Caption = AirflowSide
Lb_AirflowSide_cfm.Caption = Int(AirflowSide / 1.11 * 100) / 100

' update value for packing factor
PackingFactor = Tb_PackingF

'compute the static pressure of the system, Pa
AirfResistance = AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor) * BinHeight
Tb_AirfRes = AirfResistance
'compute the static pressure of the system, inches of water
AirfResistance_Wa = AirfResistance * PaToInWa
Lb_AirfRes_Wa.Caption = Int(AirfResistance_Wa * 100) / 100
'compute the fan prewarming, ºC
FanPreWarming_C = Int(CompFanPreWarm(AirfResistance_Wa) * 100) / 100
'print fan prewarming information
Tb_FanPreWarm = FanPreWarming_C
Lb_FanPreWarm_F.Caption = Int(FanPreWarming_C * 1.8 * 100) / 100

'update values for the fan efficiency
FanEfficiency = Tb_FanEfficiency
' update the power requirement for the fan operation, HP
FanPower_HP = Int(CompFanPower(TotalAirflow, AirfResistance_Wa, FanEfficiency) * 100) / 100
' update the power requirement for the fan operation, KW
FanPower_KW = Int(FanPower_HP * HPtoKW * 100) / 100
'print power information
Lb_FanPower_KW.Caption = FanPower_KW
Lb_FanPower_HP.Caption = FanPower_HP

' estimate energy needed for the temperature increase due to the burner
' desired temperature increase
Tinc_C = Tb_BurnerTempInc.Text
'default air conditions needed for the estimation
RH = 70
TC = 15
' compute the estimate power required
HeatingEnergy_KW = CompBurnerPower(Tinc_C, TC, RH, TotalAirflow)
'print the information
Lb_HeatingEnergy_KW.Caption = Int(HeatingEnergy_KW * 100) / 100
Lb_HeatingEnergy_BTU.Caption = Int(KWHtoBTU * HeatingEnergy_KW * 100) / 100
' set the burner efficiency
BurnerEfficiency = Tb_BurnerEff.Text
'compute the energy required by the burner to increase the temperature of the air
BurnerEnergy_KW = HeatingEnergy_KW / (BurnerEfficiency / 100)
BurnerEnergy_BTU = BurnerEnergy_KW * KWHtoBTU
'Print the Information
Lb_BurnerEnergy_KWH.Caption = Int(BurnerEnergy_KW * 100) / 100
Lb_BurnerEnergy_BTU.Caption = Int(BurnerEnergy_BTU)


End Sub

Private Sub Form_Activate()
Lb_TotalAirflow.Caption = Int(TotalAirflow)
Lb_TotalAirflow_cfm.Caption = Int(AirflowRate_cfm * BinCapacity_bu)
'set the airflow rate at the center and side of the bin according to the non-uniformity factor
AirflowCenter = Int(AirflowRate * (1 - AirflowNonUnif) * 100) / 100
AirflowSide = Int(AirflowRate * (1 + AirflowNonUnif) * 100) / 100
' print information about airflow in the screen
Lb_AirflowCenter.Caption = AirflowCenter
Lb_AirflowCenter_cfm.Caption = Int(AirflowCenter / 1.11 * 100) / 100
Lb_AirflowSide.Caption = AirflowSide
Lb_AirflowSide_cfm.Caption = Int(AirflowSide / 1.11 * 100) / 100

'compute the static pressure of the system, Pa
AirfResistance = AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor) * BinHeight
Tb_AirfRes = AirfResistance
'compute the static pressure of the system, inches of water
AirfResistance_Wa = AirfResistance * PaToInWa
Lb_AirfRes_Wa.Caption = Int(AirfResistance_Wa * 100) / 100
'compute the fan prewarming, ºC
FanPreWarming_C = Int(CompFanPreWarm(AirfResistance_Wa) * 100) / 100
'print fan prewarming information
Tb_FanPreWarm = FanPreWarming_C
Lb_FanPreWarm_F.Caption = Int(FanPreWarming_C * 1.8 * 100) / 100

Lb_FanPower_KW.Caption = FanPower_KW
Lb_FanPower_HP.Caption = FanPower_HP

' estimate energy needed for the temperature increase due to the burner
' desired temperature increase
Tinc_C = Tb_BurnerTempInc.Text
'default air conditions needed for the estimation
RH = 70
TC = 15
' compute the estimate power required
HeatingEnergy_KW = CompBurnerPower(Tinc_C, TC, RH, TotalAirflow)
'print the information
Lb_HeatingEnergy_KW.Caption = Int(HeatingEnergy_KW * 100) / 100
Lb_HeatingEnergy_BTU.Caption = Int(KWHtoBTU * HeatingEnergy_KW * 100) / 100
'Print the Information
Lb_BurnerEnergy_KWH.Caption = Int(BurnerEnergy_KW * 100) / 100
Lb_BurnerEnergy_BTU.Caption = Int(BurnerEnergy_BTU)
End Sub

Private Sub Form_Load()
Tb_AirF_NUF = AirflowNonUnif * 2
' set the default value for  the packing factor to compute static pressure to 1
Tb_PackingF = PackingFactor
' set the default value for the fan efficiency to 50%
Tb_FanEfficiency = FanEfficiency
Tb_BurnerEff = BurnerEfficiency
'set a defoult value for the burner temp increase
Tb_BurnerTempInc.Text = Tinc_C
'default air conditions needed for the estimation
RH = 70
TC = 15
' compute the estimate power required
HeatingEnergy_KW = CompBurnerPower(Tinc_C, TC, RH, TotalAirflow)

BurnerEnergy_KW = HeatingEnergy_KW / (BurnerEfficiency / 100)
BurnerEnergy_BTU = BurnerEnergy_KW * KWHtoBTU
'Print the Information
Lb_BurnerEnergy_KWH.Caption = Int(BurnerEnergy_KW * 100) / 100
Lb_BurnerEnergy_BTU.Caption = Int(BurnerEnergy_BTU)

End Sub
