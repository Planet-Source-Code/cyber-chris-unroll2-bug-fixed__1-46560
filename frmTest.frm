VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Unroll 2"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture2 
      Height          =   8655
      Left            =   0
      ScaleHeight     =   573
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   725
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin MSComDlg.CommonDialog CD 
         Left            =   4440
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   8700
      Left            =   -600
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11580
   End
   Begin VB.Menu mnueff 
      Caption         =   "Effects"
      Begin VB.Menu mnuef1 
         Caption         =   "Old Unroll"
         Begin VB.Menu mnuSimpleUnroll1 
            Caption         =   "Simple Unroll"
         End
         Begin VB.Menu unroll2 
            Caption         =   "Slow Unroll"
         End
      End
      Begin VB.Menu mnuleff 
         Caption         =   "Lens effect"
         Begin VB.Menu mnuef2 
            Caption         =   "Lens effect with Picture invert"
         End
         Begin VB.Menu mnueff2 
            Caption         =   "Lens effect"
         End
      End
      Begin VB.Menu mnuef3 
         Caption         =   "Unroll with mirror effect"
      End
      Begin VB.Menu mnuef4 
         Caption         =   """Cut"" the pic"
      End
      Begin VB.Menu mnuef5 
         Caption         =   "Unroll with after effect"
      End
      Begin VB.Menu mnunpw 
         Caption         =   """Needle Printer"" effects"
         Begin VB.Menu mnuef6 
            Caption         =   "Top - Down"
         End
         Begin VB.Menu mnudup 
            Caption         =   "Down - Up"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Test Picture1, Picture2
End Sub

Private Sub mnudup_Click()
NeedlePrinterDownUp Picture1, Picture2
End Sub

Private Sub mnuef2_Click()
UnrollWithLens Picture1, Picture2
End Sub

Private Sub mnuef3_Click()
UnrollWithMirror Picture1, Picture2
End Sub

Private Sub mnuef4_Click()
CutThePic Picture1, Picture2, 30
End Sub

Private Sub mnuef5_Click()
UnrollWithAfterEffect Picture1, Picture2
End Sub

Private Sub mnuef6_Click()
NeedlePrinter Picture1, Picture2
End Sub

Private Sub mnueff2_Click()
Lens2 Picture1, Picture2
End Sub

Private Sub mnuSimpleUnroll1_Click()
SimpleUnroll Picture1, Picture2
End Sub

Private Sub unroll2_Click()
SlowUnroll Picture1, Picture2
End Sub
