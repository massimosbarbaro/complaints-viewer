VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGrafici 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Graphics"
   ClientHeight    =   6255
   ClientLeft      =   4935
   ClientTop       =   2115
   ClientWidth     =   5325
   Icon            =   "frmGrafici.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   5325
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Visualizza"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Frame fraGrafici 
      Height          =   3015
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   3375
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Famiglia"
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Rete"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Famiglia"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Difetto"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   585
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Esterni"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Risposta"
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
      End
   End
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdRicerche 
      Caption         =   "Ricerche rete"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccess 
      Caption         =   "Access"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4560
      Picture         =   "frmGrafici.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   615
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   1560
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
      Left            =   3000
      TabIndex        =   15
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Graphics"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1950
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmGrafici.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmGrafici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIndice As Integer

Private Sub cmdAccess_Click()
Dim ac As Access.Application

'Set ac = New Access.Application
Set ac = CreateObject("Access.Application")
' put the path to your database in hereac.
'ac.OpenCurrentDatabase ("C:\Programmi\reclami\reclami.mdb")
ac.OpenCurrentDatabase ("C:\Programmi\reclami\reclami.mdb")
' by default, the OpenReport method of the
' DoCmd object will send the report to the printerac.
ac.DoCmd.OpenReport "rDateFamiglie", acViewPreview

' close the database
ac.CloseCurrentDatabase
End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdGrafico_Click()
   Dim sSelFormula As String, sGrafico As String
   Dim ac As Access.Application, sReport As String
   Dim sFile As String, sWhere As String, sTabella As String
   

'-------------------------------------------------------------------------------
' Preparo le date limite per la stampa dei grafici

   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio, dDataFine
   Dim sTitolo As String, sTest As String, sLinea As String, sGraficoFamiglia As String
   Dim sReteGrafico As String, sQtaGrafico As String, sGraficiReti As String
   
   
   
   
   If mskDataA.Text = "__/____" Then
      mskDataA.Text = mskDataDA.Text
   End If
   dDataFine = Format((DateAdd("d", -1, DateAdd("m", 1, "01/" & mskDataA.Text))), "dd/mm/yyyy")
   
   sTest = Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   dDataInizio = Format(sTest, "yyyy/mm/dd")
   'Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   If dDataFine <= dDataInizio Then
      MsgBox "Date Errate !!"
      Exit Sub
   End If
'Fine delle date
'---------------------------------------------------------------------------------
   'Apri DB con i report
   Set ac = New Access.Application

   sFile = App.Path & "\..\reclami\" & "reclami.mdb"
   ac.OpenCurrentDatabase sFile
   '("C:\Programmi\reclami\reclami.mdb")
'   MsgBox "Stiamo Elaborando il grafico"
     sWhere = " DataReclamo Between #" & dDataInizio & "# And #" & dDataFine & "#"
    Select Case iIndice
     Case 0
        sReport = "rGDifetto"
         sTabella = "SELECT Difetto.DescrizioneDifetto, Count(ATC.NumeroRisposta) " & _
                    "AS ConteggioDiNumeroRisposta, Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, " & _
                    "Year([DataBobinatrice]) AS [date] INTO [$$Difetto] " & _
                    "FROM ((Difetto INNER JOIN ATC ON Difetto.CodDifetto = ATC.CodDifetto) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL) INNER JOIN ReclamoVenditore ON ATC.ODL " & _
                    "= ReclamoVenditore.ODL " & _
                    "WHERE (((ATC.Soluzione)='0' Or (ATC.Soluzione) Is Null) " & _
                    "AND ((ATC.Esterno)=No) AND " & sWhere & _
                    ") " & _
                    "GROUP BY Difetto.DescrizioneDifetto, Year([DataBobinatrice]) " & _
                    "HAVING (((Year([DataBobinatrice]))>1998))"
    
     Case 1
        sReport = "rGFamiglia12"
        
        sLinea = "SELECT MFG.LineaProdotto, Sum(MFG.Qta) AS SommaDiQta INTO [$$Linea] " & _
                    "From MFG " & _
                    "WHERE databobinatrice Between #" & dDataInizio & "# And #" & dDataFine & "#" & _
                    "GROUP BY MFG.LineaProdotto"

        ac.DoCmd.RunSQL sLinea
        
        EseguiQuery
       
        sTabella = "SELECT Famiglia.DescrizioneFamiglia, Count(ATC.ODL) AS Numero, " & _
                    "Sum(ATC.QtaAccettata) AS Accettata, FinaleFamiglia.SommaDiSommaDiQta " & _
                    "AS Spedita, Sum(MKG.Reso) AS SommaDiReso, Sum(MKG.Accredito) AS " & _
                    "SommaDiAccredito INTO [$$Famiglia]  " & _
                    "FROM (((Famiglia INNER JOIN FinaleFamiglia ON Famiglia.CodFamiglia " & _
                    "= FinaleFamiglia.CodFamiglia) INNER JOIN ATC ON Famiglia.CodFamiglia " & _
                    "= ATC.CodFamiglia) LEFT JOIN MKG ON ATC.ODL = MKG.ODL) INNER JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL " & _
                    "WHERE " & sWhere & " AND (ATC.Esterno)=No AND (ATC.Soluzione)='0' " & _
                    "GROUP BY Famiglia.DescrizioneFamiglia, FinaleFamiglia.SommaDiSommaDiQta"

        sGraficoFamiglia = " SELECT Famiglia.DescrizioneFamiglia, Count(ATC.NumeroRisposta) AS " & _
                    "ConteggioDiNumeroRisposta, Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, " & _
                    "Year([DataBobinatrice]) AS [date] INTO [$$Grafico]" & _
                    "FROM ((Famiglia INNER JOIN ATC ON Famiglia.CodFamiglia = ATC.CodFamiglia) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL) INNER JOIN ReclamoVenditore ON ATC.ODL = " & _
                    "ReclamoVenditore.ODL " & _
                    "WHERE (ATC.Soluzione)='0' AND (ATC.Esterno)=No AND " & sWhere & _
                    " GROUP BY Famiglia.DescrizioneFamiglia, Year([DataBobinatrice]) " & _
                    "HAVING (((Year([DataBobinatrice]))>1998))"
        ac.DoCmd.RunSQL sGraficoFamiglia

     Case 2
        sReport = "rGRete"
        sReteGrafico = " SELECT Left([Rete],1) AS ReteR, Year([DataBobinatrice]) AS Anno, " & _
                        "Sum(MFG.Qta) AS SommaDiQta INTO [$$ReteGrafico] " & _
                        "From MFG " & _
                        "WHERE (MFG.Esterni) Is Null AND (MFG.DataBobinatrice) Between #" & dDataInizio & "# And #" & dDataFine & "#" & _
                        " GROUP BY Left([Rete],1), Year([DataBobinatrice])"
        
        ac.DoCmd.RunSQL sReteGrafico
          
        sQtaGrafico = "SELECT Year([DataReclamo]) AS AnnoReclamo, Left([Rete],1) AS ReteR, " & _
                        "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata INTO [$$QtaGrafico] " & _
                        "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                        "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                        "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                        " GROUP BY Year([DataReclamo]), Left([Rete],1)"

        ac.DoCmd.RunSQL sQtaGrafico
'in questo caso stabella è recgrafico
        sTabella = "SELECT Year([DataReclamo]) AS AnnoReclamo, Left([Rete],1) AS ReteR, " & _
                    "Count(ATC.NumeroRisposta) AS ConteggioDiNumeroRisposta INTO [$$RecGrafico] " & _
                    "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                    "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                    " GROUP BY Year([DataReclamo]), Left([Rete],1)"
        
        sGraficiReti = "SELECT UCase(Left$([Rete],1)) AS ReteRagg, Count(ATC.NumeroRisposta) " & _
                        "AS ConteggioDiNumeroRisposta, Year([DataBobinatrice]) AS [date], " & _
                        "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata INTO [$$GraficiReti] " & _
                        "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                        "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                        "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                        " GROUP BY UCase(Left$([Rete],1)), Year([DataBobinatrice]) " & _
                        "HAVING (((Year([DataBobinatrice]))>1998))"
        
        ac.DoCmd.RunSQL sGraficiReti
        
     Case 3
        sReport = "rGEsterno"
        sTabella = "SELECT Count(ATC.NumeroRisposta) AS ConteggioDiNumeroRisposta, " & _
                    "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, Year([DataBobinatrice]) " & _
                    "AS [date], ATC.Soluzione, MFG.Esterni, ReclamoVenditore.DataReclamo " & _
                    "INTO [$$Esterni] " & _
                    "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "LEFT JOIN MFG ON ATC.ODL = MFG.ODL " & _
                    "Where (((ATC.IndiceQualita) = No) And ((ATC.Esterno) = Yes)) And " & sWhere & _
                    "GROUP BY Year([DataBobinatrice]), ATC.Soluzione, MFG.Esterni, " & _
                    "ReclamoVenditore.DataReclamo " & _
                    "HAVING ((Year([DataBobinatrice]))>1998) "
      Case 4
       sReport = "rTotaleRiepilogo"
       sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
                    "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
                    "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
                    "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
                    "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
                    "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
                    "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
                    "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
                    "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
                    "WHERE " & sWhere

       
       
     Case 5
       sReport = "rTotaleRiepilogoFamiglie"
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
              "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
              "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
              "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
              "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
              "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
              "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
              "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
              "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
              "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
              "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
              "WHERE " & sWhere

     Case 6
       sReport = "rTotaleRiepilogoRisposta"
   
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
              "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
              "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
              "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
              "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
              "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
              "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
              "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
              "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
              "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
              "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
              "WHERE " & sWhere
   
   End Select
   
   
      If iIndice < 8 Then
         'Mando in stampa il report selezionato
'         ac.DoCmd.OpenReport sReport, acViewPreview, , sWhere
             
        ac.DoCmd.RunSQL sTabella
        ac.DoCmd.OpenReport sReport, acViewPreview
         ac.Visible = True
      
      
      
      
      Else
'         With cryGrafici
'         On Error Resume Next
'            .Destination = crptToWindow
'            .ReportFileName = App.Path & "\" & sGrafico
'            .WindowState = crptNormal
'            If iIndice = 6 Or (iIndice = 7) Then
'               sSelFormula = "{qReteRaggruppata.Soluzione}='0' and {qReteRaggruppata.IndiceQualita}=true and {qReteRaggruppata.Esterno}=false "
'            Else
'               sSelFormula = "{ATC.Soluzione}='0' and {ATC.IndiceQualita}=true  and {ATC.Esterno}=false "
'            End If
'            If iIndice < 4 Then
'            .SelectionFormula = sSelFormula
'            End If
'
'            .Action = 1
            If Err.Number <> 0 Then
               MsgBox "Errore nel produrre il report " & vbCrLf & _
                   Err.Number & " - " & Err.Description
               Err.Clear
            End If
 '        End With
      End If
  ' MsgBox "Il grafico è stato completato"
   ' close the database
  ' ac.CloseCurrentDatabase

   
End Sub
Private Sub EseguiQuery()
Dim q As QueryDef, db As Database, sqlc As String
 
 Set db = CurrentDb
    sqlc = "Drop table  Linea"
    On Error Resume Next
    db.Execute (sqlc)
    On Error GoTo 0
    sqlc = "Drop table FinaleFamiglia"
    On Error Resume Next
    db.Execute (sqlc)
    On Error GoTo 0
    sqlc = "Drop table $$Famiglia"
    On Error Resume Next
    db.Execute (sqlc)
    On Error GoTo 0
    
    
    
    
    Set q = db.QueryDefs("qQtaLinea1")
    q.Execute
    Set q = db.QueryDefs("qQtaLinea2")
    q.Execute
    Set q = db.QueryDefs("qQtaLinea3")
    q.Execute
    Set q = db.QueryDefs("qQtaLinea4")
    q.Execute
    Set q = db.QueryDefs("qFinaleFamiglia")
    q.Execute
    
    
    
End Sub
Private Sub cmdRicerche_Click()
         Load frmRicerca
         frmRicerca.Show

End Sub

Private Sub Form_Load()
iIndice = -99
   Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "mm/yyyy")
   Me.Show
'   mskDataDA.SetFocus
    optGrafico(0) = True
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

Private Sub optGrafico_Click(Index As Integer)
iIndice = Index

End Sub

