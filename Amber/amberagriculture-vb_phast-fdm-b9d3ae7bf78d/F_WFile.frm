VERSION 5.00
Begin VB.Form F_WFile 
   Caption         =   "F_WFile"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Co_FDay 
      Height          =   315
      Left            =   7080
      TabIndex        =   20
      Text            =   "Day"
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox Co_FMonth 
      Height          =   315
      Left            =   6120
      TabIndex        =   19
      Text            =   "Month"
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox Co_FYear 
      Height          =   315
      Left            =   5160
      TabIndex        =   18
      Text            =   "Year"
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox Co_IDay 
      Height          =   315
      Left            =   7080
      TabIndex        =   17
      Text            =   "Day"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox Co_IMonth 
      Height          =   315
      Left            =   6120
      TabIndex        =   16
      Text            =   "Month"
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox Co_IYear 
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Text            =   "Year"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Com_Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   5640
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change T && RH Columns"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Weather File"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Initial File Date"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Month"
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Year"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Month"
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Day"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Day"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Final File Date"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label36 
      Caption         =   "File Information"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "F_WFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Com_Close_Click()
'update file date information
InitialFileYear = Co_IYear.ListIndex + 1961
InitialFileMonth = Co_IMonth.ListIndex + 1
InitialFileDay = Co_IDay.ListIndex + 1
FinalFileYear = Co_FYear.ListIndex + 1961
FinalFileMonth = Co_FMonth.ListIndex + 1
FinalFileDay = Co_FDay.ListIndex + 1

If File1.FileName = Empty Then
    F_Error.Show
    F_Error.Lb_ErrorMsg.Caption = "Error!!!! You Must Select a Weather File"
Else
    F_WFile.Hide
    F_Main.Show
    FileNameW = File1.Path + "\" + File1.FileName
End If
'write the name of the selected weather file for the run information string
StWeatherFile = File1.FileName
End Sub

Private Sub Command2_Click()
F_WFChCol.Show
F_WFile.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
ChDrive Drive1.Drive
End Sub

Private Sub File1_Click()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
' load the list for initial and final dates of the weather file
i = 1961
For i = 1961 To 2000
    Co_IYear.AddItem i
    Co_FYear.AddItem i
Next i
i = 1
For i = 1 To 12
    Co_IMonth.AddItem i
    Co_FMonth.AddItem i
Next i
i = 1
For i = 1 To 31
    Co_IDay.AddItem i
    Co_FDay.AddItem i
Next i


'set defoult values for the weather file according to the SAMSON weather database
Co_IYear.ListIndex = 0
Co_FYear.ListIndex = 39
Co_IMonth.ListIndex = 0
Co_FMonth.ListIndex = 11
Co_IDay.ListIndex = 0
Co_FDay.ListIndex = 30
End Sub
