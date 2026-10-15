VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmRicerca 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ricerche rete"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   Icon            =   "frmRicerca.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdStampa 
      Caption         =   "Stampa"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4320
      Picture         =   "frmRicerca.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   615
   End
   Begin Crystal.CrystalReport cryRete 
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox lstRete 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataA 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport cryRiepilogo 
      Left            =   4560
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ricerche per rete"
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
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3705
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmRicerca.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmRicerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdStampa_Click()
StampaMese
End Sub

Private Sub Form_Load()
   Dim sqlc As String, rsRete As Recordset

   Me.MousePointer = vbArrowHourglass
   Me.Show
   Me.Refresh
   ShowWait "Attendere !!" & vbCrLf & vbCrLf & "Caricamento ODL in corso"
   sqlc = "SELECT distinct ucase(left(MFG.Rete,1)) as rete from MFG INNER JOIN ATC ON MFG.ODL = ATC.ODL " & _
            "GROUP BY MFG.Rete, 1 " & _
            "Having ((Not (MFG.Rete) Is Null))ORDER BY 1"
            
   Set rsRete = db.OpenRecordset(sqlc, dbOpenSnapshot)
   Do While Not rsRete.EOF
   
      lstRete.AddItem rsRete(0)
      
      rsRete.MoveNext
   Loop
'   rsRete.Close
'   Set rsRete = Nothing
   lstRete.ListIndex = -1
   lstRete.SetFocus
   Me.MousePointer = vbNormal
   HideWait
   
   
      Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "mm/yyyy")
   Me.Show
   mskDataDA.SetFocus

End Sub

Private Sub lstRete_DblClick()
'Dim sTmp As String
'   Dim sSelFormula As String''
'
'
'   sTmp = lstRete.List(lstRete.ListIndex)
''   With cryRete
'      On Error Resume Next
'      .DataFiles(0) = db.Name
 '     .Destination = crptToWindow
'      .ReportFileName = App.Path & "\ricercarete.rpt" '
'      .WindowState = crptNormal
'      sSelFormula = "{qRiepilogoRete.Rete} = '" & sTmp & "'"
'      .SelectionFormula = sSelFormula
'      .Action = 1
'      'If Err.Number <> 0 Then
''          MsgBox "Errore nel produrre il report " & vbCrLf & _
'              Err.Number & " - " & Err.Description
 '         Err.Clear
 '     End If
'   End With


End Sub

Private Sub Stampa()
'Dim sTmp As String
'   Dim sSelFormula As String
'
'
'   sTmp = lstRete.List(lstRete.ListIndex)
'   With cryRete
''      On Error Resume Next
'      .DataFiles(0) = db.Name
 '     .Destination = crptToWindow
 '     .ReportFileName = App.Path & "\ricercarete.rpt"
 ''     .WindowState = crptNormal
 '     sSelFormula = "{qRiepilogoRete.Rete} = '" & sTmp & "'"
  ''    .SelectionFormula = sSelFormula
  '    .Action = 1
  ''    If Err.Number <> 0 Then
  '        MsgBox "Errore nel produrre il report " & vbCrLf & _
  '            Err.Number & " - " & Err.Description
  '        Err.Clear
   '   End If
   'End With

End Sub
Private Sub mskDataA_Validate(Cancel As Boolean)
   If mskDataA.Text <> "__/____" Then
      If Not IsDate("01/" & mskDataA.Text) Then
         MsgBox "Data A Non valida"
         Cancel = True
      End If
   Else
      mskDataA.Text = mskDataDA.Text
   End If
End Sub

Private Sub mskDataDA_Validate(Cancel As Boolean)
'  La DATADA NON può essere nulla !!
'   If mskDataDA.Text <> "__/____" Then
   If Not IsDate("01/" & mskDataDA.Text) Then
      MsgBox "Data DA Non valida"
      Cancel = True
   End If
'   End If
End Sub


Private Sub StampaMese()
   Dim dDataInizio As Date, dDataFine As Date, sSelFormula As String
   Dim sTitolo As String
   Dim sTmp As String



   sTmp = lstRete.List(lstRete.ListIndex)

   If mskDataA.Text = "__/____" Then
      mskDataA.Text = mskDataDA.Text
   End If
   dDataFine = DateAdd("d", -1, DateAdd("m", 1, "01/" & mskDataA.Text))
   dDataInizio = DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1)
   If dDataFine <= dDataInizio Then
      MsgBox "Date Errate !!"
      Exit Sub
   End If
'
'  Preparo per il titolo
'
   If mskDataDA.Text = mskDataA.Text Then
      sTitolo = "Indice di Qualità per il mese di: " & _
                  Format$(dDataInizio, "mmmm yyyy")
   Else
      If Year(dDataInizio) = Year(dDataFine) Then
         sTitolo = "Dettaglio Reclami per il periodo: " & _
                     Format$(dDataInizio, "mmmm") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      Else
         sTitolo = "Dettaglio Reclami per il periodo: " & _
                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      End If
   End If
   
   With cryRiepilogo
      On Error Resume Next
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\ricercarete.rpt"
      .Formulas(0) = "Titolo=" & Chr$(34) & sTitolo & Chr$(34)
      .WindowState = crptNormal
      sSelFormula = "{qRiepilogoRete.Rete} = '" & sTmp & "'" & _
                        "and ({qRiepilogoRete.DataReclamo})>= Date (" & _
                        Year(dDataInizio) & ", " & _
                        Month(dDataInizio) & ", " & _
                        Day(dDataInizio) & ") and " & _
                        "{qRiepilogoRete.DataReclamo} <= Date (" & _
                        Year(dDataFine) & ", " & _
                        Month(dDataFine) & ", " & _
                        Day(dDataFine) & ")"

      .SelectionFormula = sSelFormula
      .Action = 1
         If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With

End Sub

