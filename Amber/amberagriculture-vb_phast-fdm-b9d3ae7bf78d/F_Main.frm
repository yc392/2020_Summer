VERSION 5.00
Begin VB.Form F_Main 
   Caption         =   "F_Main"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Com_Close 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H000000FF&
      TabIndex        =   89
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Tb_SAVHmc 
      Height          =   285
      Left            =   7560
      TabIndex        =   84
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Tb_TempIncr 
      Height          =   285
      Left            =   10920
      TabIndex        =   82
      Top             =   6480
      Width           =   495
   End
   Begin VB.Frame Fr_Summary 
      Caption         =   "Select Output"
      Height          =   1335
      Left            =   9240
      TabIndex        =   79
      Top             =   6960
      Width           =   2175
      Begin VB.OptionButton Opt_All 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Summary 
         Caption         =   "Only Summary"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox Tb_BaseName 
      Height          =   285
      Left            =   6960
      TabIndex        =   77
      Text            =   "Test"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox Tb_CrMaxEMC 
      Height          =   285
      Left            =   8280
      TabIndex        =   76
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox Tb_CrMinEMC 
      Height          =   285
      Left            =   6960
      TabIndex        =   75
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox Tb_CrMaxRH 
      Height          =   285
      Left            =   8280
      TabIndex        =   74
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox Tb_CrMinRH 
      Height          =   285
      Left            =   6960
      TabIndex        =   73
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox Tb_CrMaxTemp 
      Height          =   285
      Left            =   8280
      TabIndex        =   72
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox Tb_CrMinTemp 
      Height          =   285
      Left            =   6960
      TabIndex        =   71
      Top             =   7320
      Width           =   495
   End
   Begin VB.ComboBox Co_FMultYear 
      Height          =   315
      Left            =   10320
      TabIndex        =   61
      Text            =   "Finish"
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox Co_IMultYear 
      Height          =   315
      Left            =   9240
      TabIndex        =   60
      Text            =   "Start"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Run Simulation for:"
      Height          =   1095
      Left            =   5160
      TabIndex        =   57
      Top             =   4560
      Width           =   3015
      Begin VB.OptionButton Opt_MultYears 
         Caption         =   "Multiple Years"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Opt_SingleYear 
         Caption         =   "One Single Year"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox Co_FSDay 
      Height          =   315
      Left            =   10320
      TabIndex        =   55
      Text            =   "Day"
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox Co_FSMonth 
      Height          =   315
      Left            =   9480
      TabIndex        =   54
      Text            =   "Month"
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox Co_FSYear 
      Height          =   315
      Left            =   8520
      TabIndex        =   53
      Text            =   "Year"
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox Co_ISDay 
      Height          =   315
      Left            =   8280
      TabIndex        =   52
      Text            =   "Day"
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Co_ISMonth 
      Height          =   315
      Left            =   7440
      TabIndex        =   51
      Text            =   "Month"
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Co_ISYear 
      Height          =   315
      Left            =   6480
      TabIndex        =   50
      Text            =   "Year"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Tb_NumberLayers 
      Height          =   285
      Left            =   1920
      TabIndex        =   47
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Com_Run 
      Caption         =   "Run"
      Height          =   375
      Left            =   3120
      TabIndex        =   46
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Tb_FinalMax 
      Height          =   285
      Left            =   10680
      TabIndex        =   44
      Text            =   "16"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Tb_FinalAvg 
      Height          =   285
      Left            =   10680
      TabIndex        =   43
      Text            =   "15"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Simulation Ends Based on:"
      Height          =   1815
      Left            =   5160
      TabIndex        =   36
      Top             =   2280
      Width           =   3015
      Begin VB.OptionButton Opt_DML 
         Caption         =   "Final DML"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Opt_Temp 
         Caption         =   "Final Temperature"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Opt_MC 
         Caption         =   "Final Moisture Content"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Opt_Date 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Tb_Strat2 
      Height          =   285
      Left            =   5520
      TabIndex        =   28
      Text            =   "60"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Tb_Strat1 
      Height          =   285
      Left            =   5520
      TabIndex        =   27
      Text            =   "20"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Com_AdvanceSettings 
      Caption         =   "Advance Settings"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Com_UpdateSet 
      Caption         =   "Update Settings"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox Tb_AvgGrainTemp 
      Height          =   285
      Left            =   2280
      TabIndex        =   21
      Text            =   "20"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Tb_AvgGrainMC 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Text            =   "20"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Tb_BinHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "7"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Tb_BinDiam 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "12"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Tb_Airflow 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Text            =   "1"
      Top             =   7320
      Width           =   735
   End
   Begin VB.ComboBox Co_SelectGrain 
      Height          =   315
      ItemData        =   "F_Main.frx":0000
      Left            =   720
      List            =   "F_Main.frx":0002
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox Co_SelectStrat 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Lb_LayerDepth 
      Caption         =   "L Depth: "
      Height          =   255
      Left            =   2760
      TabIndex        =   88
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "PURDUE UNIVERSITY"
      Height          =   495
      Left            =   8760
      TabIndex        =   87
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   9960
      Picture         =   "F_Main.frx":0004
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label12 
      Caption         =   "Post Harvest Education and Research Center"
      Height          =   615
      Left            =   1080
      TabIndex        =   86
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Lb_SAVHmc 
      Caption         =   "Desired Final MC (%, W.B.)"
      Height          =   255
      Left            =   5400
      TabIndex        =   85
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Lb_TempIncr 
      Caption         =   "Temp. Increase (ºC)"
      Height          =   255
      Left            =   9360
      TabIndex        =   83
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Lb_BaseFileName 
      Caption         =   "Output Files Base Name "
      Height          =   255
      Left            =   5040
      TabIndex        =   78
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Lb_CNACriteria_EMC 
      Alignment       =   1  'Right Justify
      Caption         =   "EMC (%, W.B.)"
      Height          =   255
      Left            =   5520
      TabIndex        =   70
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Lb_CNACriteria_RH 
      Alignment       =   1  'Right Justify
      Caption         =   "RH (%)"
      Height          =   255
      Left            =   5520
      TabIndex        =   69
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Lb_CnaCriteria_Temp 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperature (ºC)"
      Height          =   255
      Left            =   5400
      TabIndex        =   68
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Lb_CNACriteria_Max 
      Caption         =   "Maximum Value"
      Height          =   255
      Left            =   8280
      TabIndex        =   67
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Lb_CNACriteria_Min 
      Caption         =   "Minumum Value"
      Height          =   255
      Left            =   6960
      TabIndex        =   66
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Lb_CNACriteria 
      Caption         =   "Set Limits for the Fan Operation"
      Height          =   255
      Left            =   6960
      TabIndex        =   65
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Lb_FMultYear 
      Caption         =   "Final"
      Height          =   255
      Left            =   10320
      TabIndex        =   64
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Lb_IMultYear 
      Caption         =   "Initial"
      Height          =   255
      Left            =   9240
      TabIndex        =   63
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Lb_MultipleYears 
      Caption         =   "Set Multiple Simulation Years"
      Height          =   255
      Left            =   9240
      TabIndex        =   62
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Lb_ISDate 
      Caption         =   "Lb_ISDate"
      Height          =   255
      Left            =   6480
      TabIndex        =   56
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Lb_EndOfSim 
      Caption         =   "xxxx"
      Height          =   375
      Left            =   6000
      TabIndex        =   49
      Top             =   8400
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Number of Layers"
      Height          =   255
      Left            =   600
      TabIndex        =   48
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Lb_ValuesCriteria 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8760
      TabIndex        =   45
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Lb_MaxFinal 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8280
      TabIndex        =   42
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Lb_AvgFinal 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8280
      TabIndex        =   41
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Lb_Strat7 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8520
      TabIndex        =   35
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Lb_Strat6 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Lb_Strat5 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   6600
      TabIndex        =   33
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Lb_Strat4 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   6960
      TabIndex        =   32
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Lb_Strat3 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Lb_Strat2 
      Alignment       =   1  'Right Justify
      Caption         =   "xxxx"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Lb_Strat1 
      Alignment       =   2  'Center
      Caption         =   "xxxx"
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Moisture Content:"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "ºC"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Temperature:"
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "%, W.B."
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Average Initial Grain Settings"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Lb_BinCap_bu 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Lb_BinCap_t 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Bin Capacity"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Lb_BinHeight_ft 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Lb_BinDiam_ft 
      Caption         =   "xxxx"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Heigth (m)"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Diameter (m)"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Set Bin Dimensions"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Lb_AirflowCFM 
      Caption         =   "xxxxx"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Set Airflow Rate (m3/min/t)"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Select Grain"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Lb_SelectStrat 
      Caption         =   "Select Drying Strategy"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "PHAST-FDM Drying Model"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Co_SelectGrain_Click()
GrainIndex = Co_SelectGrain.ListIndex
' to allow for DML computation only for corn varieties
If ArrGrain(GrainIndex, 14) = "1" Then
    Opt_DML.Visible = True
Else
    Opt_DML.Visible = False
End If
End Sub

Private Sub Com_FanBurner_Click()
End Sub

Private Sub Co_SelectStrat_Click()

Lb_Strat1.Visible = False
Lb_Strat2.Visible = False
Lb_Strat3.Visible = False
Lb_Strat4.Visible = False
Lb_Strat5.Visible = False
Lb_Strat6.Visible = False
Lb_Strat7.Visible = False
Tb_Strat1.Visible = False
Tb_Strat2.Visible = False
Co_ISYear.Visible = False
Co_ISMonth.Visible = False
Co_ISDay.Visible = False
Lb_ISDate.Visible = False
Lb_SAVHmc.Visible = False
Tb_SAVHmc.Visible = False

Lb_CNACriteria.Visible = False
Lb_CNACriteria_Min.Visible = False
Lb_CNACriteria_Max.Visible = False
Lb_CnaCriteria_Temp.Visible = False
Lb_CNACriteria_RH.Visible = False
Lb_CNACriteria_EMC.Visible = False
Tb_CrMinTemp.Visible = False
Tb_CrMaxTemp.Visible = False
Tb_CrMinRH.Visible = False
Tb_CrMaxRH.Visible = False
Tb_CrMinEMC.Visible = False
Tb_CrMaxEMC.Visible = False

Lb_TempIncr.Visible = False
Tb_TempIncr.Visible = False

Select Case Co_SelectStrat.ListIndex
    Case Is = 0
        Lb_Strat1.Visible = True
        Lb_Strat2.Visible = True
        Lb_Strat3.Visible = True
        Lb_Strat4.Visible = True
        Lb_Strat5.Visible = True
        Lb_Strat6.Visible = True
        Lb_Strat7.Visible = True
        Tb_Strat1.Visible = True
        Tb_Strat2.Visible = True

        Lb_Strat1.Caption = "Set Fix Inlet Air Conditions"
        Lb_Strat2.Caption = "Temperature (ºC)"
        Lb_Strat3.Caption = "Relative Humidity (%)"
        Lb_Strat4.Caption = "EMC drying, % w.b."
        Lb_Strat5.Caption = "EMC rewetting, % w.b."
        AirTemp_C = Tb_Strat1
        AirRH = Tb_Strat2
        Lb_Strat6.Caption = Format(Mdb_Mwb(CF_EMC_D(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")
        Lb_Strat7.Caption = Format(Mdb_Mwb(CF_EMC_R(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")
        Frame2.Visible = False
        Lb_MultipleYears.Visible = False
        Lb_IMultYear.Visible = False
        Lb_FMultYear.Visible = False
        Co_IMultYear.Visible = False
        Co_FMultYear.Visible = False
        MultYears = False
        FanStatus = True
        HeaterStatus = False
        F_WFile.Hide
        Fr_Summary.Visible = False
        
    Case Else
        F_WFile.Show
        Lb_ISDate.Visible = True
        Lb_ISDate.Caption = "Set Start Simulation Date"
        Co_ISYear.Visible = True
        Co_ISMonth.Visible = True
        Co_ISDay.Visible = True
        Frame2.Visible = True
        Lb_MultipleYears.Visible = True
        Lb_IMultYear.Visible = True
        Lb_FMultYear.Visible = True
        Co_IMultYear.Visible = True
        Co_FMultYear.Visible = True
        Fr_Summary.Visible = True

        'load initial simulation date lists
        i = 1961
        For i = 1961 To 2000
            Co_ISYear.AddItem i
        Next i
        i = 1
        For i = 1 To 12
            Co_ISMonth.AddItem i
        Next i
        i = 1
        For i = 1 To 31
            Co_ISDay.AddItem i
        Next i
        Co_ISYear.ListIndex = InitialSimYear - 1961
        Co_ISMonth.ListIndex = InitialSimMonth - 1
        Co_ISDay.ListIndex = InitialSimDay - 1
        
        'load the multiple year simulation lists
        i = 1961
        For i = 1961 To 2000
            Co_IMultYear.AddItem i
            Co_FMultYear.AddItem i
        Next i
        Co_IMultYear.ListIndex = IMultYear - 1961
        Co_FMultYear.ListIndex = FMultYear - 1961
        
        If Co_SelectStrat.ListIndex = 1 Then
             'for the CNA strategy
            Lb_CNACriteria.Visible = True
            Lb_CNACriteria_Min.Visible = True
            Lb_CNACriteria_Max.Visible = True
            Lb_CnaCriteria_Temp.Visible = True
            Lb_CNACriteria_RH.Visible = True
            Lb_CNACriteria_EMC.Visible = True
            Tb_CrMinTemp.Visible = True
            Tb_CrMaxTemp.Visible = True
            Tb_CrMinRH.Visible = True
            Tb_CrMaxRH.Visible = True
            Tb_CrMinEMC.Visible = True
            Tb_CrMaxEMC.Visible = True
            Lb_CNACriteria.Caption = "Set the Limits for the Fan Operation"
            Tb_CrMinTemp = MinTempSelect
            Tb_CrMaxTemp = MaxTempSelect
            Tb_CrMinRH = MinRHSelect
            Tb_CrMaxRH = MaxRHSelect
            Tb_CrMinEMC = MinEMCSelect
            Tb_CrMaxEMC = MaxEMCSelect
                       
        End If
        If Co_SelectStrat.ListIndex = 2 Then
             'for the Constant Heat strategy
            Lb_CNACriteria.Visible = True
            Lb_CNACriteria_Min.Visible = True
            Lb_CNACriteria_Max.Visible = True
            Lb_CnaCriteria_Temp.Visible = False
            Lb_CNACriteria_RH.Visible = True
            Lb_CNACriteria_EMC.Visible = True
            Tb_CrMinTemp.Visible = False
            Tb_CrMaxTemp.Visible = False
            Tb_CrMinRH.Visible = True
            Tb_CrMaxRH.Visible = True
            Tb_CrMinEMC.Visible = True
            Tb_CrMaxEMC.Visible = True
            Lb_TempIncr.Visible = True
            Tb_TempIncr.Visible = True
            Lb_CNACriteria.Caption = "Set the Limits for the Fan and Heater"
            Tb_CrMinRH = MinRHSelect
            Tb_CrMaxRH = MaxRHSelect
            Tb_CrMinEMC = MinEMCSelect
            Tb_CrMaxEMC = MaxEMCSelect
                                    
        End If
        
        If Co_SelectStrat.ListIndex = 3 Then
             'for the Variable Heat strategy
            Lb_CNACriteria.Visible = True
            Lb_CNACriteria_Min.Visible = True
            Lb_CNACriteria_Max.Visible = True
            Lb_CnaCriteria_Temp.Visible = False
            Lb_CNACriteria_RH.Visible = False
            Lb_CNACriteria_EMC.Visible = True
            Tb_CrMinTemp.Visible = False
            Tb_CrMaxTemp.Visible = False
            Tb_CrMinRH.Visible = False
            Tb_CrMaxRH.Visible = False
            Tb_CrMinEMC.Visible = True
            Tb_CrMaxEMC.Visible = True
            Lb_TempIncr.Visible = False
            Tb_TempIncr.Visible = False
            Lb_CNACriteria.Caption = "Set the Limits for the Fan and Heater"
            Tb_CrMinEMC = MinEMCSelect
            Tb_CrMaxEMC = MaxEMCSelect
                                    
        End If
        If Co_SelectStrat.ListIndex = 4 Then
             'for the Self Adapting Variable Heat strategy
            Lb_CNACriteria.Visible = False
            Lb_CNACriteria_Min.Visible = False
            Lb_CNACriteria_Max.Visible = False
            Lb_CnaCriteria_Temp.Visible = False
            Lb_CNACriteria_RH.Visible = False
            Lb_CNACriteria_EMC.Visible = False
            Tb_CrMinTemp.Visible = False
            Tb_CrMaxTemp.Visible = False
            Tb_CrMinRH.Visible = False
            Tb_CrMaxRH.Visible = False
            Tb_CrMinEMC.Visible = False
            Tb_CrMaxEMC.Visible = False
            Lb_TempIncr.Visible = False
            Tb_TempIncr.Visible = False
            Lb_SAVHmc.Visible = True
            Tb_SAVHmc.Visible = True
            Tb_SAVHmc = SAVH_FinalMC
        End If
        
End Select
    
End Sub

Private Sub Com_AdvanceSettings_Click()
F_Main.Hide
F_AdvanceSettings.Show
End Sub

Private Sub Com_Close_Click()
End
End Sub

Private Sub Com_Run_Click()

Lb_EndOfSim.Caption = "Running Simulation"
'to update all the settings before running the simulation
Com_UpdateSet.value = True

'set all the weather file variables to 0
TotalDaysToFile = 0
TotalDaysStartSim = 0
DaysToStart = 0
FirstLineRead = 0
TotalDaysFinishSim = 0
DaysToFinish = 0
LastLineRead = 0
CurrentYear = 0

'select drying strategy
Select Case Co_SelectStrat.ListIndex
    Case Is = 0
        Call FixInlet_Strat(AirTemp_C, AirRH)
    Case Is = 1
        Call CNA_Strat
    Case Is = 2
        Call ConstHeat_Strat
    Case Is = 3
        Call VarHeat_Strat
    Case Is = 4
        Call SAVH_Strat
End Select



Lb_EndOfSim.Caption = "Simulation Completed"
End Sub

Private Sub Com_UpdateSet_Click()
'set the basename for the output files
BaseName = Tb_BaseName

' update change in initial grain MC
AvgInGrainMC = Tb_AvgGrainMC
' update change in initial grain Temp
AvgInGrainTemp = Tb_AvgGrainTemp
' update change in grain density
AvgGrainDensity = (ComputeGrainDensity(GrainIndex, AvgInGrainMC, ArrGrain)) / 1000

'update initial MC and temperature conditions fo all the layers
i = 0
For i = 0 To 31
GrainMC_In_WB_C(i) = AvgInGrainMC
GrainMC_In_WB_S(i) = AvgInGrainMC
GrainTemp_In_C_C(i) = AvgInGrainTemp
GrainTemp_In_C_S(i) = AvgInGrainTemp
Next i

' update change in bin diameter
If Tb_BinDiam = Empty Then
    BinDiameter = 1#
Else
    If Tb_BinDiam <= 0 Then
        BinDiameter = 0.001
    Else
        BinDiameter = Tb_BinDiam
    End If
End If
Lb_BinDiam_ft.Caption = (Int(BinDiameter / 0.3048 * 100)) / 100 & " ft"


' update change in bin height
If Tb_BinHeight = Empty Then
    BinHeight = 1#
Else
    If Tb_BinHeight <= 0 Then
        BinHeight = 0.001
    Else
        BinHeight = Tb_BinHeight
    End If
End If
Lb_BinHeight_ft.Caption = (Int(BinHeight / 0.3048 * 100)) / 100 & " ft"

' compute the area of the bin
BinArea = pi * (BinDiameter * BinDiameter) / 4
' compute bin storage capacity in tonnes
BinCapacity_t = BinArea * BinHeight * AvgGrainDensity
'disply updated change in bin capacity (tonnes)
Lb_BinCap_t.Caption = (Int(BinCapacity_t * 100)) / 100 & " tonnes"
'update change in bin capacity (bushels)
BinCapacity_bu = BinCapacity_t * 40
'disply updated change in bin capacity (bushels)
Lb_BinCap_bu.Caption = (Int(BinCapacity_bu * 100)) / 100 & " bushels"


' update change in airflow rate
AirflowRate = Tb_Airflow
If AirflowRate = Empty Then AirflowRate = 1
AirflowRate_cfm = (Int(AirflowRate / 1.11 * 100)) / 100
Lb_AirflowCFM.Caption = AirflowRate_cfm & " cfm/bu"
'update change in total airflow rate
TotalAirflow = AirflowRate * BinCapacity_t

'set the air speed at the center and side of the bin
AirVel_C = TotalAirflow / BinArea / 60 * (1 - AirflowNonUnif)
AirVel_S = TotalAirflow / BinArea / 60 * (1 + AirflowNonUnif)


' update the number of layer in the bin (0.5m each layer), and adjust the layer depth if #of layers is > 32
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

Lb_LayerDepth.Caption = "L Depth: " & Format(SingleLayerDepth, "0.00") & " m"
Tb_NumberLayers = NumberOfLayers
' set grain layer depth
Call SetGrainLayerDepth(NumberOfLayers, BinHeight, SingleLayerDepth)
' update initial simulation dates
InitialSimYear = Co_ISYear.ListIndex + 1961
InitialSimMonth = Co_ISMonth.ListIndex + 1
InitialSimDay = Co_ISDay.ListIndex + 1
' update final simulation dates
FinalSimYear = Co_FSYear.ListIndex + 1961
FinalSimMonth = Co_FSMonth.ListIndex + 1
FinalSimDay = Co_FSDay.ListIndex + 1
'update multiple year dates

IMultYear = Co_IMultYear.ListIndex + 1961
If Opt_MultYears = True Then
    FMultYear = Co_FMultYear.ListIndex + 1961
Else
    FMultYear = Co_ISYear
End If

'update emc calculation based on the fixed inlet air conditios
AirTemp_C = Tb_Strat1
AirRH = Tb_Strat2
Lb_Strat6.Caption = Format(Mdb_Mwb(CF_EMC_D(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")
Lb_Strat7.Caption = Format(Mdb_Mwb(CF_EMC_R(AirTemp_C, AirRH / 100, ArrGrain, GrainIndex) / 100) * 100, "#0.##")

'end of simulation criteria
If Opt_Date = True Then
    'update number of hours for the end of the simulation
    Hours_Stop = Tb_FinalAvg
    StStopSimulation = "Stop Sim.: Date (" & FinalSimMonth & "/" & FinalSimDay & ")"
End If
If Opt_MC = True Then
    'update MC limit for ending criteria
    MCavg_Stop = Tb_FinalAvg
    MCmax_Stop = Tb_FinalMax
    StStopSimulation = "Stop Sim.: MC (" & MCavg_Stop & "-" & MCmax_Stop & ")"

End If
If Opt_Temp = True Then
    'update temp limit for ending criteria
    Tempavg_Stop = Tb_FinalAvg
    Tempmax_Stop = Tb_FinalMax
    StStopSimulation = "Stop Sim.: Temp.(" & Tempavg_Stop & "-" & Tempmax_Stop & ")"
End If
If Opt_DML = True Then
    'update DML limit for ending criteria
    DMLavg_Stop = Tb_FinalAvg
    StStopSimulation = "Stop Sim.: DML(" & dmlpavg_Stop & ")"
End If

'update selecting criteria values
If Co_SelectStrat.ListIndex >= 1 Then
    If Co_SelectStrat.ListIndex = 1 Then
        MinTempSelect = Tb_CrMinTemp
        MaxTempSelect = Tb_CrMaxTemp
        MinRHSelect = Tb_CrMinRH
        MaxRHSelect = Tb_CrMaxRH
        MinEMCSelect = Tb_CrMinEMC
        MaxEMCSelect = Tb_CrMaxEMC
    ElseIf Co_SelectStrat.ListIndex = 2 Then
        MinRHSelect = Tb_CrMinRH
        MaxRHSelect = Tb_CrMaxRH
        MinEMCSelect = Tb_CrMinEMC
        MaxEMCSelect = Tb_CrMaxEMC
    ElseIf Co_SelectStrat.ListIndex = 3 Then
        MinEMCSelect = Tb_CrMinEMC
        MaxEMCSelect = Tb_CrMaxEMC
    Else
    'update the desired final moiture content for the savh strategy
    SAVH_FinalMC = Tb_SAVHmc
    End If
End If

'update the static pressure of the system, Pa
AirfResistance = AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor) * BinHeight
'update the static pressure of the system, inches of water
AirfResistance_Wa = AirfResistance * PaToInWa

'update values for the fan power
FanPower_HP = Int(CompFanPower(TotalAirflow, AirfResistance_Wa, FanEfficiency) * 100) / 100
' update the power requirement for the fan operation, KW
FanPower_KW = Int(FanPower_HP * HPtoKW * 100) / 100
'update the temperature increse value demanded to the burner
Tinc_C = Tb_TempIncr


'set the bin dimensions for the run information string
StBin = "Height: " & BinHeight & ", Diameter: " & BinDiameter
'set the airflow for the run information string
StAirflow = AirflowRate
'set the initial grain temperature and MC for the run information string
StGrain = "Initial Temp.: " & AvgInGrainTemp & ", Initial MC: " & AvgInGrainMC
'set the initial month and day of the simulation for the run information string
StStartDate = InitialSimMonth & "/" & InitialSimDay
'set the selected grain EMC for the run information string
StGrainEMC = ArrGrain(GrainIndex, 0)
'select drying strategy settings for the run information string
Select Case Co_SelectStrat.ListIndex
    Case Is = 0
        StStrategy = "Fix Inlet (" & AirTemp_C & "ºC" & ", " & AirRH & "%)"
    Case Is = 1
        StStrategy = "CNA (EMC: " & MinEMCSelect & "-" & MaxEMCSelect & "%, Temp: " & MinTempSelect & "-" & MaxTempSelect & "ºC, RH: " & MinRHSelect & "-" & MaxRHSelect & "%)"
    Case Is = 2
        StStrategy = "Const. Heat(EMC: " & MinEMCSelect & "-" & MaxEMCSelect & "%, RH: " & MinRHSelect & "-" & MaxRHSelect & "%, Temp. Inc.: " & Tinc_C & "ºC)"
    Case Is = 3
        StStrategy = "Var. Heat(EMC: " & MinEMCSelect & "-" & MaxEMCSelect & "%)"
    Case Is = 4
        StStrategy = "SAVH (Target MC: " & SAVH_FinalMC & "%)"
End Select

StRunInfo = "This run was made with the following settings: Weather File: " & StWeatherFile & "; Strategy: " & StStrategy & "; Grain EMC: " & StGrainEMC & "; Bin Dimensions: " & StBin & "; Grain: " & StGrain & "; Airflow: " & StAirflow & " m3/min/t; Starting Date: " & StStartDate & "; " & StStopSimulation


End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Form_Load()
CurrentDir = CurDir
' load the in-bin drying strategy list
Co_SelectStrat.AddItem "Fix Inlet Conditions"
Co_SelectStrat.AddItem "Natural Air"
Co_SelectStrat.AddItem "Constant Heat"
Co_SelectStrat.AddItem "Variable Heat"
Co_SelectStrat.AddItem "Self Adapting Variable Heat"

'location of the file with the list of grains available and it parameters
FileName = CurrentDir & "\graininfo.dry"
Call ReadGrainList1(FileName)
'set initial option for the grain list to "0"
Co_SelectGrain.ListIndex = 0

Lb_Strat1.Visible = True
Lb_Strat2.Visible = True
Lb_Strat3.Visible = False
Tb_Strat1.Visible = False
Tb_Strat2.Visible = False
Co_ISYear.Visible = False
Co_ISMonth.Visible = False
Co_ISDay.Visible = False
Co_FSYear.Visible = False
Co_FSMonth.Visible = False
Co_FSDay.Visible = False
Frame2.Visible = False
Lb_MultipleYears.Visible = False
Lb_IMultYear.Visible = False
Lb_FMultYear.Visible = False
Co_IMultYear.Visible = False
Co_FMultYear.Visible = False

'set defoult values for fix inlet air conditions
Tb_Strat1 = 20
Tb_Strat2 = 70

'set initial option for the strategy list to "0"
Co_SelectStrat.ListIndex = 0

'set the defoult values for initial simulation date
InitialSimYear = 1961
InitialSimMonth = 10
InitialSimDay = 1
'set the defoult values for final simulation date
FinalSimYear = 1961
FinalSimMonth = 12
FinalSimDay = 31
'set the defoult values for the multiple years simulation
IMultYear = 1961
FMultYear = 2000

'set defoult values for the temp, RH and EMC windows selecting criteria for the CNA strategy
MinTempSelect = -50
MaxTempSelect = 50
MinRHSelect = 0
MaxRHSelect = 100
MinEMCSelect = 0
MaxEMCSelect = 50

'set average initial values for temp and MC
AvgInGrainMC = Tb_AvgGrainMC.Text
AvgInGrainTemp = Tb_AvgGrainTemp.Text

'set initial MC and temperature conditions fo all the layers
i = 0
For i = 0 To 31
GrainMC_In_WB_C(i) = AvgInGrainMC
GrainMC_In_WB_S(i) = AvgInGrainMC
GrainTemp_In_C_C(i) = AvgInGrainTemp
GrainTemp_In_C_S(i) = AvgInGrainTemp
Next i

AvgGrainDensity = (ComputeGrainDensity(GrainIndex, AvgInGrainMC, ArrGrain)) / 1000
SingleLayerDepth1 = 0.5
' compute the number of layer in the bin (0.5m each layer)
If BinHeight - Int(BinHeight) < 0.25 Then
    NumberOfLayers = Int(BinHeight / SingleLayerDepth1)
Else
    NumberOfLayers = Int(BinHeight / SingleLayerDepth1) + 1
End If


' express the airflow rate in cfm/bu
AirflowRate = Tb_Airflow.Text
If AirflowRate = Empty Then AirflowRate = 1
AirflowRate_cfm = (Int(AirflowRate / 1.11 * 100)) / 100
Lb_AirflowCFM.Caption = AirflowRate_cfm & " cfm/bu"


'set defoult value for airflow non uniformity factor
AirflowNonUnif = 0.3 / 2

' express the bin diameter in ft
If Tb_BinDiam = Empty Then
    BinDiameter = 1#
Else
    If Tb_BinDiam <= 0 Then
        BinDiameter = 0.001
    Else
        BinDiameter = Tb_BinDiam.Text
    End If
End If
Lb_BinDiam_ft.Caption = (Int(BinDiameter / 0.3048 * 100)) / 100 & " ft"
' express the bin height in ft
If Tb_BinHeight = Empty Then
    BinHeight = 1#
Else
    If Tb_BinHeight <= 0 Then
        BinHeight = 0.001
    Else
        BinHeight = Tb_BinHeight.Text
    End If
End If
Lb_BinHeight_ft.Caption = (Int(BinHeight / 0.3048 * 100)) / 100 & " ft"
' compute the area of the bin
BinArea = pi * (BinDiameter * BinDiameter) / 4
' compute bin storage capacity in tonnes
BinCapacity_t = BinArea * BinHeight * AvgGrainDensity
Lb_BinCap_t.Caption = (Int(BinCapacity_t * 100)) / 100 & " tonnes"
' compute bin storage capacity in bushels
BinCapacity_bu = BinCapacity_t * 40
Lb_BinCap_bu.Caption = (Int(BinCapacity_bu * 100)) / 100 & " bushels"

'update change in total airflow rate
TotalAirflow = AirflowRate * BinCapacity_t

'set the air speed at the center and side of the bin
AirVel_C = TotalAirflow / BinArea / 60 * (1 - AirflowNonUnif)
AirVel_S = TotalAirflow / BinArea / 60 * (1 + AirflowNonUnif)


' update the number of layer in the bin (0.5m each layer), and adjust the layer depth if #of layers is > 32
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
Lb_LayerDepth.Caption = "L Depth: " & SingleLayerDepth & " m"
Tb_NumberLayers = NumberOfLayers
' set grain layer depth
Call SetGrainLayerDepth(NumberOfLayers, BinHeight, SingleLayerDepth)

'set defoult conditions for ending simulation criteria
Lb_ValuesCriteria.Caption = "Set Final MC Values"
Lb_AvgFinal.Visible = True
Lb_MaxFinal.Visible = True
Tb_FinalAvg.Visible = True
Tb_FinalMax.Visible = True
Lb_AvgFinal.Caption = "Average Final MC (%, w.b.)"
Lb_MaxFinal.Caption = "Maximum Final MC (%, w.b.)"
MCavg_Stop = Tb_FinalAvg
MCmax_Stop = Tb_FinalMax
'set the criteria to check end of simulation
DMLCriteria = False
MCCriteria = True
TempCriteria = False

'set defoult DML limit to 0.5%
DMLavg_Stop = 0.5
'set defoult values for multipliers for dml computation
DML_Mult_Fungicide = 1
DML_Mult_Genetics = 1

'set defoult conditions for selection criteria for CNA strategy
MinTempSelect = -30
MaxTempSelect = 40
MinRHSelect = 0
MaxRHSelect = 100
MinEMCSelect = 0
MaxEMCSelect = 30

'set defoult weather file columns for temperature and RH
TempColumn = 5
RHColumn = 6

'set defoult value for fan prewarming
' set value for packing factor
PackingFactor = 1
'set defoult value for the Temperature increase for the burner
Tinc_C = 2
Tb_TempIncr = Tinc_C
'compute the static pressure of the system, Pa
AirfResistance = AirFlowResistance(GrainIndex, ArrGrain, TotalAirflow, BinArea, PackingFactor) * BinHeight
'compute the static pressure of the system, inches of water
AirfResistance_Wa = AirfResistance * PaToInWa
'compute the fan prewarming, ºC
FanPreWarming_C = CompFanPreWarm(AirfResistance_Wa)

'set defoult values for the fan efficiency
FanEfficiency = 50
' update the power requirement for the fan operation, HP
FanPower_HP = Int(CompFanPower(TotalAirflow, AirfResistance_Wa, FanEfficiency) * 100) / 100
' update the power requirement for the fan operation, KW
FanPower_KW = Int(FanPower_HP * HPtoKW * 100) / 100
'set defoult values for the burner efficiency
BurnerEfficiency = 80
'set the defoult optiopn for multiple years simulation as true
MultYears = True
'set the defolut string for the base name of the output files
BaseName = Tb_BaseName
'set the defoult value for the output summary as true
S_OutputSummary = True
' swt defoult value fo the desired final MC for the savhstrategy
SAVH_FinalMC = 15
'set the defoult values for the variables related to the cost computation of drying
GrainPrice = 93.6 '$/tonne
ElectricityCost = 0.09 '$/kwh
PropaneCost = 0.5 '$/gallon
DesiredFinMC = 15 '(% final MC)
HeaterType = True '(Liquid propane)

End Sub

Private Sub Opt_All_Click()
S_OutputSummary = False
End Sub

Private Sub Opt_Date_Click()
If Co_SelectStrat.ListIndex = 0 Then
    Co_FSYear.Visible = False
    Co_FSMonth.Visible = False
    Co_FSDay.Visible = False
    Lb_ValuesCriteria.Caption = "Set Number of Hours"
    Lb_AvgFinal.Visible = True
    Lb_MaxFinal.Visible = False
    Tb_FinalAvg.Visible = True
    Tb_FinalMax.Visible = False
    Lb_AvgFinal.Caption = "Number of Hours"
    Hours_Stop = Tb_FinalAvg
Else
    Lb_ValuesCriteria.Caption = "Set Final Simulation Date"
    Lb_AvgFinal.Visible = False
    Lb_MaxFinal.Visible = False
    Tb_FinalAvg.Visible = False
    Tb_FinalMax.Visible = False
    Co_FSYear.Visible = True
    Co_FSMonth.Visible = True
    Co_FSDay.Visible = True
    'load final simulation date lists
        i = 1961
        For i = 1961 To 1989
            Co_FSYear.AddItem i
        Next i
        i = 1
        For i = 1 To 12
            Co_FSMonth.AddItem i
        Next i
        i = 1
        For i = 1 To 31
            Co_FSDay.AddItem i
        Next i
        Co_FSYear.ListIndex = FinalSimYear - 1961
        Co_FSMonth.ListIndex = FinalSimMonth - 1
        Co_FSDay.ListIndex = FinalSimDay - 1
End If
'set the criteria to check end of simulation
DMLCriteria = False
MCCriteria = False
TempCriteria = False
DateCriteria = True
End Sub

Private Sub Opt_DML_Click()
Lb_ValuesCriteria.Caption = "Set Final DML Value"
Lb_AvgFinal.Visible = True
Lb_MaxFinal.Visible = False
Tb_FinalAvg.Visible = True
Tb_FinalMax.Visible = False
Lb_AvgFinal.Caption = "Average Final DML (%)"
DMLavg_Stop = Tb_FinalAvg
'set the criteria to check end of simulation
DMLCriteria = True
MCCriteria = False
TempCriteria = False
DateCriteria = False
Co_FSYear.Visible = False
Co_FSMonth.Visible = False
Co_FSDay.Visible = False
End Sub

Private Sub Opt_MC_Click()
Lb_ValuesCriteria.Caption = "Set Final MC Values"
Lb_AvgFinal.Visible = True
Lb_MaxFinal.Visible = True
Tb_FinalAvg.Visible = True
Tb_FinalMax.Visible = True
Lb_AvgFinal.Caption = "Average Final MC (%, w.b.)"
Lb_MaxFinal.Caption = "Maximum Final MC (%, w.b.)"
MCavg_Stop = Tb_FinalAvg
MCmax_Stop = Tb_FinalMax
'set the criteria to check end of simulation
DMLCriteria = False
MCCriteria = True
TempCriteria = False
DateCriteria = False
Co_FSYear.Visible = False
Co_FSMonth.Visible = False
Co_FSDay.Visible = False
End Sub

Private Sub Opt_MultYears_Click()
Co_FMultYear.Visible = True
Co_IMultYear.Visible = True
Lb_MultipleYears.Visible = True
Lb_IMultYear.Visible = True
Lb_FMultYear.Visible = True
IMultYear = Co_IMultYear
FMultYear = Co_FMultYear
MultYears = True
Fr_Summary.Visible = True
S_OutputSummary = True
End Sub

Private Sub Opt_SingleYear_Click()
Co_FMultYear.Visible = False
Co_IMultYear.Visible = False
Lb_MultipleYears.Visible = False
Lb_IMultYear.Visible = False
Lb_FMultYear.Visible = False
FMultYear = Co_ISYear
MultYears = False
Fr_Summary.Visible = False
S_OutputSummary = False
End Sub

Private Sub Opt_Summary_Click()
S_OutputSummary = True
End Sub

Private Sub Opt_Temp_Click()
Lb_ValuesCriteria.Caption = "Set Final Temperature Values"
Lb_AvgFinal.Visible = True
Lb_MaxFinal.Visible = True
Tb_FinalAvg.Visible = True
Tb_FinalMax.Visible = True
Lb_AvgFinal.Caption = "Average Final Temp. (ºC)"
Lb_MaxFinal.Caption = "Maximum Final Temp. (ºC)"
Tempavg_Stop = Tb_FinalAvg
Tempmax_Stop = Tb_FinalMax
'set the criteria to check end of simulation
DMLCriteria = False
MCCriteria = False
TempCriteria = True
DateCriteria = False
End Sub

