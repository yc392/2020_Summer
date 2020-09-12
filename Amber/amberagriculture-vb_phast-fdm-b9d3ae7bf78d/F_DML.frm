VERSION 5.00
Begin VB.Form F_DML 
   Caption         =   "F_DML"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Com_SetValues 
      Caption         =   "Set Values"
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Tb_Genetics 
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Tb_Fungicide 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Tb_Damage 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1170
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Susceptible Hybrid: 0.91"
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Resistant Hybrid: 1.25"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Generic Hybrid: 1"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "Genetic Multiplier:"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Corn Treated with 80 ppm Soybean Oil: 1.1"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   "Corn Treated with 20 ppm Iprodione: 1.2"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "No Fungicide Application: 1"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Hand Shelled Corn: 3%"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Heavy Harvest Damage: 40%"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Combine Harvested Corn: 30%"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "%"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Fungicide Multiplier:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Damage Multiplier:"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Set Defoult Values for DML Multipliers"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "DML Information and Settings"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "F_DML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()
F_DML.Hide
F_AdvanceSettings.Show
End Sub

Private Sub Com_SetValues_Click()
' update the values for the multipliers
GDamage = Tb_Damage
DML_Mult_Fungicide = Tb_Fungicide
DML_Mult_Genetics = Tb_Genetics
'CO2 generation is divided by 1.47 to conver form grams of CO2 generated to g of dry matter consumed
DML_GramsPerKg = (CompCO2Prod(GT, GMC, GDamage, DML_Mult_Fungicide, DML_Mult_Genetics)) / 1.47
DML_TempIncr_C = CompDML_TempInc(DML_GramsPerKg, ArrGrain, GrainIndex, GMC)
End Sub

Private Sub Form_Load()
'Set the defoult values for the multipliers
'default value for the damage multiplier = 30%
Tb_Damage = 10
'defoult value for the fungicide multiplier = 1
Tb_Fungicide = 1
'defoult value for the genetic (hybrid) multiplier = 1
Tb_Genetics = 1
End Sub
