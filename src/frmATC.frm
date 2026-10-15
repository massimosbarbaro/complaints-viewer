VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmATC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ATC"
   ClientHeight    =   8265
   ClientLeft      =   870
   ClientTop       =   405
   ClientWidth     =   8040
   Icon            =   "frmATC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   8040
   Begin VB.TextBox txtQtaAccettata 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   6
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   7080
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":287A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   38
      Top             =   7635
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Reclamo Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su reclamo corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica il Reclamo Corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbDifetto 
      DataSource      =   "datDifetto"
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox cmbFamiglia 
      DataSource      =   "datDifetto"
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4635
      Width           =   2775
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   5
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   3405
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   1763
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Text            =   "frmATC.frx":2DBE
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "00530.1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdDettagli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Dettagli"
      Height          =   615
      Left            =   6600
      MousePointer    =   1  'Arrow
      Picture         =   "frmATC.frx":2DC4
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame fraSoluzione 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Soluzione"
      Height          =   1215
      Left            =   6120
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Segnalazione"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accettato"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Respinto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   540
         Width           =   975
      End
   End
   Begin VB.TextBox txtNumeroRisposta 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   930
      TabIndex        =   14
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox txtRisposta 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5400
      Width           =   7560
   End
   Begin VB.TextBox txtAzioni 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6360
      Width           =   7560
   End
   Begin VB.TextBox txtDataRisposta 
      Height          =   285
      Left            =   2055
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fraEsterno 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Esterno"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
      Begin VB.OptionButton optExt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optExt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraIQ 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Indice di Qualità"
      Height          =   615
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
      Begin VB.OptionButton optIq 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optIq 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta Accettata"
      Height          =   195
      Left            =   6480
      TabIndex        =   40
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label lblCodFamiglia 
      AutoSize        =   -1  'True
      Caption         =   "CodFamiglia"
      Height          =   195
      Left            =   2160
      TabIndex        =   37
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCodDifetto 
      AutoSize        =   -1  'True
      Caption         =   "CodDifetto"
      Height          =   195
      Left            =   1200
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblDifetto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difetto"
      Height          =   195
      Left            =   360
      TabIndex        =   35
      Top             =   4140
      Width           =   465
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia"
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   4680
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venditore"
      Height          =   195
      Left            =   6720
      TabIndex        =   33
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motivazioni"
      Height          =   195
      Left            =   360
      TabIndex        =   27
      Top             =   2160
      Width           =   795
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Analisi tecnica dei reclami"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   1080
      TabIndex        =   26
      Top             =   360
      Width           =   5700
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmATC.frx":32F6
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblReclamoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relcamo Numero"
      Height          =   195
      Left            =   4920
      TabIndex        =   23
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Risposta N"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   3360
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Azioni Correttive"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   6120
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Risposta"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Risposta"
      Height          =   195
      Left            =   2040
      TabIndex        =   19
      Top             =   3360
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta Reclamata"
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Reclamo"
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODL"
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   1440
      Width           =   330
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   4215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   7815
   End
End
Attribute VB_Name = "frmATC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private mStatus As Integer
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cStampa = 11
   cEsci = 13
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = False
            .Buttons(cEsci).Enabled = False
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = False
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = False
            .Buttons(cEsci).Enabled = True
      End Select
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
    CercaReclamoVenditore
    CercaReclamoATC
   
'    Set datElenco.Recordset = rsATC
'    datElenco.RecordSource = sqlc
'    txtODLInputTest.Text = mKey
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub Rete()
''   Dim lDiff As Long
'    Load frmRete
'    With frmRete
'        .Top = Me.Top + (Me.Height - frmRete.Height)
'        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
'        lDiff = Screen.Width * 0.98 - (.Left + .Width)
'        If lDiff < 0 Then
'            .Left = .Left + lDiff
'        End If
'        .Key = txtNumeroRisposta.Text
'        .Show vbModal
'    End With
End Sub

Private Sub cmdDettagli_Click()
    Dim lDiff As Long
    Load frmMFG
    With frmMFG
        .Top = Me.Top
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
        .Key = Me.Key
        .Show
    End With
End Sub




Private Sub SalvaRec()
   Dim sCod As Variant
   If fNuovoRecord Then
      rsATC.AddNew
      txtNumeroRisposta.Text = CalcUltimaRisposta()
      rsATC("NumeroRisposta") = txtNumeroRisposta.Text
      rsATC("ODL") = Me.Key
   Else
      rsATC.Edit
   End If
'   If Not IsNull(txtNumeroRisposta.Text) Then
'        rsATC("NumeroRisposta") = txtNumeroRisposta.Text
'   Else
'         rsATC("NumeroRisposta") = Null
'   End If
   If Not (IsNull(txtDataRisposta.Text) Or txtDataRisposta.Text = "") Then
      rsATC("DataRisposta") = txtDataRisposta.Text
   Else
       rsATC("DataRisposta") = Null
   End If
   sCod = SalvaCodiceDifetto(cmbDifetto.Text)
   lblCodDifetto.Caption = sCod
   If Not (IsNull(lblCodDifetto.Caption) Or lblCodDifetto.Caption = "") Then
       rsATC("CodDifetto") = lblCodDifetto.Caption
   Else
       rsATC("CodDifetto") = Null
   End If
   If Not (IsNull(txtQtaAccettata.Text) Or txtQtaAccettata.Text = "") Then
        rsATC("QtaAccettata") = Val(txtQtaAccettata.Text)
   Else
         rsATC("QtaAccettata") = Null
   End If
   
'     Carico i dati del codfamiglia dall'rs
   sCod = SalvaCodiceFamiglia(cmbFamiglia.Text)
   lblCodFamiglia = sCod
   If Not (IsNull(lblCodFamiglia.Caption) Or lblCodFamiglia.Caption = "") Then
       rsATC("CodFamiglia") = lblCodFamiglia
   Else
       rsATC("CodFamiglia") = Null
   End If
   If Not (IsNull(txtRisposta.Text) Or txtRisposta.Text = "") Then
       rsATC("Risposta") = txtRisposta
   Else
       rsATC("Risposta") = Null
   End If
   If Not (IsNull(txtAzioni.Text) Or txtAzioni.Text = "") Then
       rsATC("AzioniCorrettive") = txtAzioni.Text
   Else
       rsATC("AzioniCorrettive") = Null
   End If
   If optIq(0).Value Then
       rsATC("IndiceQualita") = True
   Else
       If optIq(1).Value Then
           rsATC("IndiceQualita") = False
       Else
           rsATC("IndiceQualita") = Null
       End If
   End If
   If optExt(0).Value Then
       rsATC("Esterno") = True
   Else
       If optExt(1).Value Then
           rsATC("Esterno") = False
       Else
           rsATC("Esterno") = Null
       End If
   End If
   If optSol(0).Value Then
       rsATC("Soluzione") = 0
   Else
       If optSol(1).Value Then
           rsATC("Soluzione") = 1
       Else
           If optSol(2).Value Then
              rsATC("Soluzione") = 2
           Else
   '            If optSol(3).Value Then
    '               rsATC("Soluzione") = 3
     '          Else
                   rsATC("Soluzione") = Null
       '        End If
          End If
       End If
   End If
   
   rsATC.Update
   CercaReclamoATC
'     disabilito i campi
   Campi False
   fNuovoRecord = False
End Sub

Private Sub Form_Load()
   Dim btn As Button
'  Carica i combo
   ListCodiceDifetto
   ListCodiceFamiglia
   
   Campi False
    'Nascondo tutti i pulsanti
'    cmdNew.Visible = False
'    cmdSalva.Visible = False
'    cmdAnnulla.Visible = False
'    cmdModifica.Visible = False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
     End If
   Next
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
      l = shpRisposta.Left * 2 + shpRisposta.Width + Me.Width - Me.ScaleWidth
      If Me.Width < l Then
         Me.Width = l
      End If
      l = 0
      Me.Refresh
      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Load frmVenditori
'    frmVenditori.Show
   rsATC.Close
   Set rsATC = Nothing
End Sub

Private Sub CercaReclamoVenditore()
   Dim rsVenditore As Recordset, sqlc As String, i As Integer
   
   sqlc = "SELECT ODL, DataReclamo, QtaReclamata, " & _
           "CausaReclamo, Venditore, NumeroReclamo, Stato " & _
           "FROM ReclamoVenditore " & _
           "where ODL='" & Me.Key & "'"
   Set rsVenditore = db.OpenRecordset(sqlc, dbOpenDynaset)
   If Not rsVenditore.EOF Then
      If rsVenditore("Stato") = "N" Then
         rsVenditore.Edit
         rsVenditore("Stato") = "N"
         rsVenditore.Update
         frmTV.ChangedTV = True
      End If
      For i = 0 To rsVenditore.Fields.Count - 1
         If Not IsNull(rsVenditore(i)) Then
             txtVenditore(i).Text = rsVenditore(i)
         Else
             txtVenditore(i).Text = ""
         End If
      Next i
      sStato = rsVenditore("Stato")
   Else
      Err.Raise 13, , "Errore nel data source"
   End If
   rsVenditore.Close
   Set rsVenditore = Nothing

End Sub
Private Sub CercaReclamoATC()
    Dim sqlc As String, i As Integer
   
    sqlc = "SELECT ODL, NumeroRisposta, DataRisposta, " & _
            "Risposta, AzioniCorrettive, CodFamiglia, " & _
            "CodDifetto, Esterno, StatoReclamo, IndiceQualita, " & _
            "Soluzione, QtaAccettata FROM ATC where ODL='" & Me.Key & "'"
    Set rsATC = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsATC.EOF Then
      FoundODL
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
    'For i = 0 To rsATC.Fields.Count - 1
    '    txtReclami(i).DataField = rsATC.Fields(i).Name
    'Next i
End Sub
Private Sub FoundODL()
   Dim iSoluz As Integer
    If Not IsNull(rsATC("NumeroRisposta")) Then
        txtNumeroRisposta = rsATC("NumeroRisposta")
    Else
        txtNumeroRisposta = ""
    End If
    If Not IsNull(rsATC("DataRisposta")) Then
        txtDataRisposta = rsATC("DataRisposta")
    Else
        txtDataRisposta = ""
    End If
    If Not IsNull(rsATC("CodDifetto")) Then
        lblCodDifetto = rsATC("CodDifetto")
        cmbDifetto.Text = FoundCodiceDifetto(lblCodDifetto.Caption)
    Else
        lblCodDifetto = ""
        cmbDifetto.ListIndex = -1
    End If
    If Not IsNull(rsATC("CodFamiglia")) Then
      lblCodFamiglia = rsATC("CodFamiglia")
      cmbFamiglia.Text = FoundCodiceFamiglia(lblCodFamiglia.Caption)
    Else
     lblCodFamiglia = ""
     cmbFamiglia.ListIndex = -1
    End If
    If Not IsNull(rsATC("Risposta")) Then
      txtRisposta = rsATC("Risposta")
    Else
        txtRisposta = ""
    End If
    If Not IsNull(rsATC("AzioniCorrettive")) Then
        txtAzioni = rsATC("AzioniCorrettive")
    Else
        txtAzioni = ""
    End If
    If Not IsNull(rsATC("IndiceQualita")) Then
        If rsATC("IndiceQualita") Then
            optIq(0).Value = True
        Else
            optIq(1).Value = True
        End If
    Else
          optIq(0).Value = False
          optIq(1).Value = False
    End If
    If Not IsNull(rsATC("Esterno")) Then
        If rsATC("Esterno") Then
            optExt(0).Value = True
        Else
            optExt(1).Value = True
        End If
    Else
        optExt(0).Value = False
        optExt(1).Value = False
    End If
    If Not IsNull(rsATC("Soluzione")) Then
      iSoluz = rsATC("Soluzione")
      optSol(iSoluz).Value = True
    Else
        optSol(0).Value = False
        optSol(1).Value = False
        optSol(2).Value = False
    End If
   If Not IsNull(rsATC("QtaAccettata")) Then
      txtQtaAccettata = rsATC("QtaAccettata")
    Else
        txtQtaAccettata = ""
    End If
End Sub
Private Sub ListCodiceDifetto()
   Dim rsDifetto As Recordset, sqlc As String, sMsg As String
   sqlc = "SELECT DescrizioneDifetto " & _
           "FROM Difetto "
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   cmbDifetto.Clear
   Do While Not rsDifetto.EOF
      cmbDifetto.AddItem rsDifetto(0)
      rsDifetto.MoveNext
   Loop
   rsDifetto.Close
   Set rsDifetto = Nothing
   cmbDifetto.ListIndex = -1
End Sub
Private Sub ListCodiceFamiglia()
    Dim rsFamiglia As Recordset, sqlc As String, sMsg As String
    sqlc = "SELECT DescrizioneFamiglia " & _
            "FROM Famiglia order by 1"
    Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
    cmbFamiglia.Clear
    Do While Not rsFamiglia.EOF
        cmbFamiglia.AddItem rsFamiglia(0)
        rsFamiglia.MoveNext
    Loop
    rsFamiglia.Close
    Set rsFamiglia = Nothing
    cmbFamiglia.ListIndex = -1
    
End Sub
Private Function FoundCodiceDifetto(cod As String)
   Dim rsDifetto As Recordset, sqlc As String
   sqlc = "SELECT DescrizioneDifetto " & _
           "FROM Difetto where CodDifetto='" & cod & "'"
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsDifetto.EOF Then
       FoundCodiceDifetto = rsDifetto("DescrizioneDifetto")
   Else
       FoundCodiceDifetto = ""
   End If
   rsDifetto.Close
   Set rsDifetto = Nothing
End Function
Private Function FoundCodiceFamiglia(cod As String)
   Dim rsFamiglia As Recordset, sqlc As String
   sqlc = "SELECT DescrizioneFamiglia " & _
           "FROM Famiglia where CodFamiglia='" & cod & "'"
   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsFamiglia.EOF Then
       FoundCodiceFamiglia = rsFamiglia("DescrizioneFamiglia")
   Else
       FoundCodiceFamiglia = ""
   End If
   rsFamiglia.Close
   Set rsFamiglia = Nothing
End Function
Private Function SalvaCodiceDifetto(s As String)
   Dim rsDifetto As Recordset, sqlc As String, sDescrizione As String
   Dim iPos As Integer
   iPos = InStr(s, "'")
   If iPos <> 0 Then
      s = Left$(s, iPos) & "'" & Right$(s, Len(s) - iPos)
      
   End If
   sqlc = "SELECT CodDifetto " & _
           "FROM Difetto where DescrizioneDifetto='" & Trim$(s) & "'"
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsDifetto.EOF Then
      SalvaCodiceDifetto = rsDifetto("CodDifetto")
   Else
      SalvaCodiceDifetto = ""
   End If
   rsDifetto.Close
   Set rsDifetto = Nothing
End Function

Private Function SalvaCodiceFamiglia(s As String)
   Dim rsFamiglia As Recordset, sqlc As String, sDescrizione As String
   
   sqlc = "SELECT CodFamiglia " & _
           "FROM Famiglia where DescrizioneFamiglia='" & s & "'"
   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsFamiglia.EOF Then
       SalvaCodiceFamiglia = rsFamiglia("CodFamiglia")
   Else
      SalvaCodiceFamiglia = ""
   End If
   rsFamiglia.Close
   Set rsFamiglia = Nothing
End Function

Private Sub Campi(Flag As Boolean)
    txtNumeroRisposta.Enabled = False
    txtDataRisposta.Enabled = Flag
    txtRisposta.Enabled = Flag
    txtAzioni.Enabled = Flag
    fraIQ.Enabled = Flag
    fraEsterno.Enabled = Flag
    fraSoluzione.Enabled = Flag
    cmbDifetto.Enabled = Flag
    cmbFamiglia.Enabled = Flag
    txtQtaAccettata.Enabled = Flag
    txtNumeroRisposta.Locked = Not Flag
    txtDataRisposta.Locked = Not Flag
    txtRisposta.Locked = Not Flag
    txtAzioni.Locked = Not Flag
    txtQtaAccettata.Locked = Not Flag
End Sub

Private Sub CancellaCampi()
   txtNumeroRisposta.Text = ""
   txtDataRisposta.Text = ""
   txtRisposta.Text = ""
   txtAzioni.Text = ""
   txtQtaAccettata.Text = ""
   optIq(0).Value = False
   optIq(1).Value = False
   
   optExt(0).Value = False
   optExt(1).Value = False
   
   optSol(0).Value = False
   optSol(1).Value = False
   optSol(2).Value = False
   
   cmbDifetto.ListIndex = -1
   cmbFamiglia.ListIndex = -1
End Sub


Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
   Select Case Button.Key
      Case "X"
         Unload Me
      Case "N"
         fNuovoRecord = True
         Campi True
         CancellaCampi
         Me.Status = cNew
      Case "S"
         SalvaRec
         Me.Status = cExists
      Case "A"
         If fNuovoRecord Then
            CancellaCampi
            Me.Status = cNotExists
         Else
            FoundODL
            Me.Status = cExists
         End If
         Campi False
      Case "M"
         fNuovoRecord = False
         Campi True
         Me.Status = cMod
      Case "C"
         iRes = MsgBox("Vuoi chiudere il corrente reclamo", _
               vbYesNo + vbDefaultButton2 + vbQuestion)
         If iRes = vbYes Then
            sqlc = "Update ReclamoVenditore set Stato = 'C' " & _
                     "Where ODL = '" & txtVenditore(0) & "' and " & _
                     "Stato = 'V'"
            db.Execute (sqlc)
            sqlc = "Update ATC set StatoReclamo = 'N' " & _
                  "where ODL = '" & txtVenditore(0) & "'"
            db.Execute (sqlc)
            sStato = "C"
            Me.Status = cExists
            frmTV.ChangedTV = True
         End If
      Case "P"
         fNuovoRecord = False
         Rete
         Me.Status = cExists
   End Select
End Sub
Private Function CalcUltimaRisposta()
   Dim rsT As Recordset, sqlc As String
   sqlc = "update UltimaRisposta " & _
               "set UltimaRisposta = UltimaRisposta + 1 " & _
               "Where dummy=0"
   db.Execute (sqlc)
   sqlc = "Select UltimaRisposta from UltimaRisposta " & _
               "Where dummy = 0"
   Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsT.EOF Then
      CalcUltimaRisposta = rsT(0)
   Else
      Err.Raise 555, , "Errore nel calcolo del numero risposta"
   End If
   rsT.Close
   Set rsT = Nothing
End Function

