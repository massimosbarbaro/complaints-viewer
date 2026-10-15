VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   5040
   ClientTop       =   3405
   ClientWidth     =   4620
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4620
   Begin VB.CommandButton cmdComplaint 
      Caption         =   "&Status"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3840
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdStatistiche 
      Caption         =   "S&tatistics"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSegnalazioni 
      Caption         =   "&Segnalazioni"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReclami 
      Caption         =   "&Reclami"
      Height          =   255
      Left            =   1560
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmMain.frx":0E14
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Customer Satisfaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   4140
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1695
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Menu mnuReclami 
      Caption         =   "Reclami"
      Visible         =   0   'False
      Begin VB.Menu mnuR 
         Caption         =   "Visualizza ATC"
         Index           =   0
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuR 
         Caption         =   "Nuovo Reclamo"
         Index           =   2
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuR 
         Caption         =   "Posta"
         Index           =   4
      End
      Begin VB.Menu mnuR 
         Caption         =   "Visualizza MKG"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComplaint_Click()
   Load frmTV
   frmTV.Param = "ATC/MKG"
   frmTV.Show

End Sub

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdReclami_Click()
'    Load frmVenditori
'    frmVenditori.Show
End Sub

Private Sub cmdReclami_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PopupMenu mnuReclami, vbPopupMenuCenterAlign
End Sub


Private Sub cmdSegnalazioni_Click()
'   Load frmImportMFG
'   frmImportMFG.Show vbModal
End Sub

Private Sub cmdStatistiche_Click()
   Screen.MousePointer = vbArrowHourglass
   Load frmGrafici
   frmGrafici.Show
   Screen.MousePointer = vbNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer, res As Integer, nRs As Integer
    res = MsgBox("Esco dall'Applicazione ?", vbQuestion + vbYesNo + vbDefaultButton1, _
        "Reclami - Uscita")
    If res = vbYes Then
'
'   Check su altre forms caricate
'
        For i = Forms.Count - 1 To 0 Step -1
            If Forms(i).Name <> Me.Name Then
    '            Set f = Forms(i)
                Unload Forms(i)
            End If
        Next i
    '
    '   Controllo rs aperti !!!  (Ben Fatto)
    '
        nRs = db.Recordsets.Count
        For i = 1 To nRs
            db.Recordsets(0).Close
        Next i
    '
    '   Chiusura DB
    '
        If fDBOk Then
            db.Close
        End If
        Set db = Nothing
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuR_Click(Index As Integer)
'   Select Case Index
'      Case 0
'         Load frmTV
'         frmTV.Param = "ATC"
'         frmTV.Show
'      Case 2
'         Load frmReclamiVenditori
'      Case 4
'         Load frmPop3
'         frmPop3.Show vbModal
'      Case 5
'         Load frmTV
'         frmTV.Param = "ALL"
'         frmTV.Show
'   End Select
End Sub

