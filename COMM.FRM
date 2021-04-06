VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOrbit 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtPB 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtgg 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtMoon 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtGov 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtPop 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtClim 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtHydro 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtAtmos 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtGrav 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6240
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtNotes 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"COMM.frx":0000
   End
   Begin VB.TextBox txtDiam 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtStrPrt 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtWtn 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtTL 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtCR 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Load"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
   Begin VB.TextBox txtCoord 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label16 
      Caption         =   "Orbit"
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "PB"
      Height          =   255
      Left            =   6000
      TabIndex        =   35
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "G G"
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Moon"
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Gov"
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Pop"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Clim"
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Hydro"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Atmos"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Grav"
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Diam"
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Str Prt"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "WTN"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "TL"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "CR"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Coord"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    prgdll.CalcDistance (Me.Combo1.Text)
End Sub

Private Sub Command3_Click()
    prgdll.LoadMap Me.Picture1, Me.Combo1.Text
End Sub
Private Sub Form_Load()
    prgdll.LoadSubsector Me.Combo1
End Sub

Private Sub OLE1_Updated(Code As Integer)
   
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prgdll.DetectLocation Me, X, Y
End Sub
