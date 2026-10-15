VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmIQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Indice di Qualità"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   Icon            =   "frmIQ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotale 
      Caption         =   "Totale"
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdRiepilogo 
      Caption         =   "Riepilogo"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   3840
      Picture         =   "frmIQ.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Indice di qualità"
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin Crystal.CrystalReport cryIQ 
      Left            =   3720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSMask.MaskEdBox mskDataA 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport cryTot 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Crystal.CrystalReport cryRiepilogo 
      Left            =   2400
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   3375
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmIQ.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmIQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click(Index As Integer)
   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio As Date, dDataFine As Date, sSelFormula As String
   Dim sTitolo As String
   
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
         sTitolo = "Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      Else
         sTitolo = "Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      End If
   End If
   
   sqlc = "Select * from Indici order by 1"
   Set rsLimiti = db.OpenRecordset(sqlc, dbOpenSnapshot)
   i = 0
   With cryIQ
      On Error Resume Next
      .DataFiles(0) = db.Name
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\newiq.rpt"
      Do While Not rsLimiti.EOF
         .Formulas(i) = "F" & Format(i, "0") & _
                     " = if ({ATC.QtaAccettata} >= " & rsLimiti("LimInf") & _
                     " and {ATC.QtaAccettata} <= " & _
                     Format$(rsLimiti("LimSup"), "000000000") & ")" & _
                     " then 1 else 0 "
         i = i + 1
         rsLimiti.MoveNext
      Loop
      .Formulas(i) = "Titolo=" & Chr$(34) & sTitolo & Chr$(34)
      .WindowState = crptNormal
'      sSelFormula = "{}=" & sDataTemp
      .Formulas(9) = "mese=" & Chr$(34) & Mid$(dDataFine, 4, 2) & Chr$(34)
      
      sSelFormula = "(not IsNull ({ATC.QtaAccettata})) and (" & _
                        "{ReclamoVenditore.DataReclamo} >= Date (" & _
                        Year(dDataInizio) & ", " & _
                        Month(dDataInizio) & ", " & _
                        Day(dDataInizio) & ") and " & _
                        "{ReclamoVenditore.DataReclamo} <= Date (" & _
                        Year(dDataFine) & ", " & _
                        Month(dDataFine) & ", " & _
                        Day(dDataFine) & "))" & "and {ATC.IndiceQualita} and {ATC.Soluzione}='0'" & _
                        " and ({ATC.QtaAccettata}>0)"

      .SelectionFormula = sSelFormula
'      .SectionFormat
      .Action = 1
      If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With
   rsLimiti.Close
   Set rsLimiti = Nothing

End Sub

Private Sub cmdRiepilogo_Click()
   Dim dDataInizio As Date, dDataFine As Date, sSelFormula As String
   Dim sTitolo As String
   
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
         sTitolo = "Dettaglio Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      Else
         sTitolo = "Dettaglio Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      End If
   End If
   
   With cryRiepilogo
      On Error Resume Next
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\gdatefamiglie.rpt"
      .Formulas(0) = "Titolo=" & Chr$(34) & sTitolo & Chr$(34)
      .WindowState = crptNormal
      sSelFormula = "not IsNull ({qDateFamiglie.QtaAccettata}) " & _
                        "and ({qDateFamiglie.DataReclamo})>= Date (" & _
                        Year(dDataInizio) & ", " & _
                        Month(dDataInizio) & ", " & _
                        Day(dDataInizio) & ") and " & _
                        "{qDateFamiglie.DataReclamo} <= Date (" & _
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

Private Sub cmdTotale_Click(Index As Integer)
   Dim dDataInizio As Date, dDataFine As Date, sSelFormula As String
   Dim sTitolo As String

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
      sTitolo = "Riepilogo reclami per il mese di: " & _
                  Format$(dDataInizio, "mmmm yyyy")
   Else
      If Year(dDataInizio) = Year(dDataFine) Then
         sTitolo = "Riepilogo reclami pervenuti per il periodo: " & _
                     Format$(dDataInizio, "mmmm") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      Else
         sTitolo = "Riepilogo reclami pervenuti per  il periodo: " & _
                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      End If
   End If
   
   With cryTot
      On Error Resume Next
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\gricapitolazione.rpt"
      .Formulas(1) = "Titolo=" & Chr$(34) & sTitolo & Chr$(34)
      .WindowState = crptNormal
      sSelFormula = " ({qTotale.DataReclamo})>= Date (" & _
                        Year(dDataInizio) & ", " & _
                        Month(dDataInizio) & ", " & _
                        Day(dDataInizio) & ") and " & _
                        "{qTotale.DataReclamo} <= Date (" & _
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

Private Sub Form_Load()
   Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "mm/yyyy")
   Me.Show
   mskDataDA.SetFocus
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

Private Sub Stampa()

   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio As Date, dDataFine As Date, sSelFormula As String
   Dim sTitolo As String
   
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
         sTitolo = "Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      Else
         sTitolo = "Indice di Qualità per il periodo: " & _
                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
                     Format$(dDataFine, "mmmm yyyy")
      End If
   End If
   
   sqlc = "Select * from Indici order by 1"
   Set rsLimiti = db.OpenRecordset(sqlc, dbOpenSnapshot)
   i = 0
   With cryIQ
      On Error Resume Next
      .DataFiles(0) = db.Name
      .Destination = crptToWindow
'Devo inserire le vrie formuel per i vari pulsanti





'_____________________________________________________________________

      .SelectionFormula = sSelFormula
'      .SectionFormat
      .Action = 1
      If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With
   rsLimiti.Close
   Set rsLimiti = Nothing


End Sub
Private Sub IndiceQualita()

End Sub
Private Sub Riepilogo()

End Sub

