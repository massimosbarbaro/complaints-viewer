VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTV 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7050
   ClientLeft      =   405
   ClientTop       =   1170
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEsci 
      Height          =   585
      Left            =   5160
      Picture         =   "frmTV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   585
   End
   Begin MSComctlLib.ImageList imglView 
      Left            =   10320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":0620
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   0
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
            Picture         =   "frmTV.frx":0A74
            Key             =   "D"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":0EC8
            Key             =   "F"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":131C
            Key             =   "R"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":1770
            Key             =   "V"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":1BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":2018
            Key             =   "C"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":2478
            Key             =   "A"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvReclami 
      Height          =   3015
      Left            =   6000
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvReclami 
      Height          =   5655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9975
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmTV.frx":28D8
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblSaveODL 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblCursor 
      BackColor       =   &H000000C0&
      Height          =   3375
      Left            =   4680
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      ToolTipText     =   "Click & Drag to resize"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   6135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lDx As Long
Private Const MinLblCursor = 75
Private lInitW As Long, lInitH As Long, lDh As Long
Private fGetFromLoad As Boolean
Private LastNode As Node
Private mChangedTV As Boolean, mParam As String
Private sNodeType As String, sWhereStato As String


Public Property Get ChangedTV() As Boolean
   ChangedTV = mChangedTV
End Property
Public Property Let ChangedTV(vValue As Boolean)
   mChangedTV = vValue
   If mChangedTV Then
      SetTV
     mChangedTV = False
   End If
End Property
Public Property Let Param(vValue As String)
   If vValue = "ATC" Or vValue = "MKG" Or vValue = "ATC/MKG" Then
      mParam = vValue
      With Me
         .ChangedTV = True
         .Caption = "Visualizzazione dei Reclami per " & mParam
      End With
   Else
      Err.Raise 17, , "Errore nel passaggio del paramtero al TV"
   End If
End Property
Public Property Get Param() As String
   Param = mParam
End Property

Private Sub SetTV()
   Dim sqlc As String, rsTV As Recordset, i As Byte
   Dim nodX As Node, nodY As Node
   Dim sStato As String


   Set LastNode = Nothing     ' Inizializzo comunque lastnode a niente
   
   With tvReclami
      .Nodes.Clear
      .ImageList = ImageList1
'
'  Crea i Nodi Principali (Root, ATC / MKG  e Stato)
'
      i = 1
      Set nodX = .Nodes.Add(, , "R", "Stato Reclamo", i)
      
'  Nodo "ATC"
      Set nodX = .Nodes.Add("R", tvwChild, "ATC", "ATC", 5)
      
      sqlc = "Select Stato, DescrStato " & _
               "From StatoReclamo " & _
               "ORDER BY Ordine"
   
      Set rsTV = db.OpenRecordset(sqlc, dbOpenSnapshot)
'
'   Costruisco la condizione Where per gli stati (dipende dall'INI)
'
    sWhereStato = " Where rv.Stato in ("
      Do While Not rsTV.EOF
         i = i + 1
         Set nodX = .Nodes.Add("ATC", tvwChild, rsTV(0) & "A", rsTV(1), i)
         sWhereStato = sWhereStato & "'" & rsTV(0) & "',"
         rsTV.MoveNext
      Loop
      sWhereStato = Left$(sWhereStato, Len(sWhereStato) - 1) & ") "
' La costringo a nullo per evitare di farla vedere
      sWhereStato = ""
      
'  Nodo "MKG"
      Set nodX = .Nodes.Add("R", tvwChild, "MKG", "MKG", 5)
      
      rsTV.MoveFirst
      i = 1
      Do While Not rsTV.EOF
         i = i + 1
         Set nodX = .Nodes.Add("MKG", tvwChild, rsTV(0) & "M", rsTV(1), i)
         rsTV.MoveNext
      Loop
      
      rsTV.Close
      
      Select Case Me.Param
         Case "ATC"
            sqlc = "SELECT Stato & 'A' as StatoR, Venditore " & _
                     "From ReclamoVenditore as RV " & _
                     sWhereStato & _
                     "GROUP BY Stato, Venditore"
         Case "MKG"
            sqlc = "SELECT ATC.StatoReclamo  & 'M' as StatoR, rv.Venditore " & _
                     "FROM ReclamoVenditore AS rv " & _
                     "INNER JOIN ATC ON rv.ODL = ATC.ODL " & _
                     sWhereStato & _
                     "GROUP BY ATC.StatoReclamo, rv.Venditore " & _
                     "Having ((Not (ATC.StatoReclamo) Is Null)) "
         Case "ATC/MKG"
            sqlc = "SELECT 'A' as AM, ([Stato] & 'A') as StatoR, Venditore " & _
                     "From ReclamoVenditore AS rv " & _
                     sWhereStato & _
                     "GROUP BY Stato, Venditore " & _
                     "Union All " & _
                     "SELECT 'M' as AM, ATC.StatoReclamo  & 'M' as StatoR, " & _
                     "rv.Venditore " & _
                     "FROM ReclamoVenditore AS rv " & _
                     sWhereStato & _
                     "INNER JOIN ATC ON rv.ODL = ATC.ODL " & _
                     "GROUP BY ATC.StatoReclamo, rv.Venditore " & _
                     "Having ((Not (ATC.StatoReclamo) Is Null)) " & _
                     "order by 1, 2, 3"
      End Select
      
      Set rsTV = db.OpenRecordset(sqlc, dbOpenSnapshot)
'
'     Aggiunge il nodo "Venditore"
'
      Do While Not rsTV.EOF
         sStato = rsTV("StatoR")
         If Not IsNull(rsTV("Venditore")) Then
            Set nodY = .Nodes.Add(sStato, tvwChild, , rsTV("Venditore"), 6)
            End If
         rsTV.MoveNext
      Loop
      
      rsTV.Close
      Set rsTV = Nothing
      
'      .BorderStyle = 1
      
      .BorderStyle = vbFixedSingle
      .Checkboxes = False
      .FullRowSelect = True
      .HideSelection = False
      .HotTracking = True
      .Indentation = 100
      .LabelEdit = tvwManual
      .LineStyle = tvwTreeLines
      .Scroll = True
'      .SingleSel = True   Non so se mi piace !!
      'Se allo Style non metti niente il programma funziona con + e -
      .Style = tvwTreelinesPlusMinusPictureText
      
      With .Nodes("R")
         .Expanded = True
         .BackColor = vbYellow
         .Bold = True
      End With
      With .Nodes("ATC")
         .Expanded = True
         .Bold = True
'         .BackColor = vbBlue
      End With
      With .Nodes("MKG")
         .Expanded = True
         .Bold = True
'         .BackColor = vbGreen
      End With
      With .Nodes("NA")
         .Bold = True
         .ForeColor = vbRed
         .Selected = True
      End With
   End With
'
'  Imposta la Header del List View
'
   SetLVHeader
   
'   tvReclami.Nodes("NA").Selected = True
End Sub
Private Sub SetLVHeader()
   Dim clX As ColumnHeader
   With lvReclami
      .ColumnHeaders.Clear
      .View = lvwReport
      .Icons = imglView
      .SmallIcons = imglView
      Select Case Me.Param
         Case "ATC"
            Set clX = .ColumnHeaders.Add(, , "ATC ODL", 1261, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data Reclamo", 1200, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "Cliente", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Prodotto", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Qta Spedita", 1065, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Qta Reclamata", 1265, lvwColumnRight)
         Case "MKG"
            Set clX = .ColumnHeaders.Add(, , "MKG ODL", 1261, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data Risposta", 1200, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "Risposta n.", 1190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Venditore", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Cliente", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Prodotto", 1065, lvwColumnRight)
         Case "ATC/MKG"
            Set clX = .ColumnHeaders.Add(, , "ATC/MKG ODL", 1261, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data", 1200, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "Cliente", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Prodotto", 2190, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Qta Spedita", 1065, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Qta Reclamata", 1265, lvwColumnRight)
      End Select
      .Arrange = lvwAutoLeft
   End With
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   Dim nodX As Node
   With tvReclami
      If .SelectedItem <> "" Then
         Set nodX = .Nodes(.SelectedItem.Index)
      Else
         Set nodX = .Nodes(1)    '  Nuovi
      End If
   End With
   tvReclami_NodeClick nodX
End Sub

Private Sub Form_Click()
'   Dim i As Integer
'   Dim strNodes As String
'   For i = 1 To tvReclami.Nodes.Count
'   strNodes = strNodes & tvReclami.Nodes(i).Index & " " & _
'               "Key: " & tvReclami.Nodes(i).Key & " " & _
'               "Text: " & tvReclami.Nodes(i).Text & vbLf
'   Next i
'   MsgBox strNodes
End Sub


Private Sub Form_Load()
   Dim sqlc As String
   
   Screen.MousePointer = vbArrowHourglass
   
   fGetFromLoad = True
   
   lblSaveODL.Visible = False
   lDx = tvReclami.Left - Shape1.Left
   With Me
      lDh = Shape1.Top
      .Height = lDh + Shape1.Height + 2 * lDh + cmdEsci.Height + _
            (.Height - .ScaleHeight)
      .Width = Shape1.Left + Shape1.Width + 2 * (.Width - .ScaleWidth)
      lInitW = .Width
      lInitH = .Height
   End With
   fGetFromLoad = False
'
'   Me.ChangedTV = True
'
'   Me.Show
'
'  Simulo un Click sul nodo "Nuovi"
'
'   Set nodX = tvReclami.Nodes("N")
'   tvReclami_NodeClick nodX
   Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
   If Not fGetFromLoad Then
      With Me
         If .Width < lInitW Then .Width = lInitW
         If .Height < lInitH Then .Height = lInitH
         Shape1.Width = .ScaleWidth - 2 * Shape1.Left
         Shape1.Height = .ScaleHeight - (3 * lDh + cmdEsci.Height)
         lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      End With
      With tvReclami
         .Height = Shape1.Height - (2 * (tvReclami.Top - Shape1.Top))
      End With
      With lvReclami
         .Left = tvReclami.Left + tvReclami.Width + MinLblCursor
         .Top = tvReclami.Top
         .Height = tvReclami.Height
         .Width = Shape1.Width + Shape1.Left - lDx - .Left
      End With
      With cmdEsci
         .Top = Me.ScaleHeight - (.Height + lDh)
         .Left = Shape1.Left + Shape1.Width - .Width
      End With
'      With lblExplain
'         .Top = cmdEsci.Top + (cmdEsci.Height - .Height) \ 2
'         .Left = lvReclami.Left + (lvReclami.Width - .Width) \ 2
'      End With
   End If
   With lblCursor
'      .Height = tvReclami.Height * 1.01
     .Height = tvReclami.Height * 0.99
     .Width = (tvReclami.Width / 2) + (lvReclami.Width / 2) + MinLblCursor
     .Top = tvReclami.Top * 1.01
     .Left = tvReclami.Left + (tvReclami.Width / 2)
     .BackColor = Shape1.BackColor
   End With
End Sub

Private Sub lblCursor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      If X < MinLblCursor Then
         X = MinLblCursor
      End If
      If X > lblCursor.Width Then
         X = lblCursor.Width
      End If
      lvReclami.Left = lblCursor.Left + X
      lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      tvReclami.Width = (lblCursor.Left + X) - tvReclami.Left - MinLblCursor
   End If
End Sub


Private Sub lvReclami_BeforeLabelEdit(Cancel As Integer)
'   MsgBox "Non è possibile modificare i valori.", vbOKOnly + vbExclamation, ""
'   Beep
   Cancel = True
End Sub

Private Sub lvReclami_DblClick()
   If lblSaveODL.Caption <> "" Then
     Select Case sNodeType
     Case "ATC"
'      If Me.Param = "ATC" Then
'       Debug.Print
         Load frmATC
         
         frmATC.Key = lblSaveODL.Caption
         frmATC.Show
      Case "MKG"
         Load frmMKG
         frmMKG.Key = lblSaveODL.Caption
         frmMKG.Show
      Case Else
       MsgBox "Questo non è un messaggio"
      End Select
      
   Else
      MsgBox "Nessun ODL Selezionato.", _
            vbOKOnly + vbExclamation
   End If
End Sub

Private Sub lvReclami_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lblSaveODL.Caption = Trim$(Item.Text)
End Sub


Private Sub tvReclami_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Index = 1 Then
      Node.Expanded = True   ' Expand the node again.
      tvReclami_NodeClick Node
   End If
End Sub

Private Sub tvReclami_Expand(ByVal Node As MSComctlLib.Node)
'   Debug.Print "Expanded " & Node.Text
End Sub

Private Sub tvReclami_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim sqlc As String, rsLV As Recordset, iImg As Integer
   Dim itmX As ListItem, i As Integer
   Dim nodeX As Node
   
   lblSaveODL.Caption = ""    '  Elimino l'ultimo odl selezionato
'
'   Pulisco comunque la parte ListView
'
    lvReclami.ListItems.Clear

   If Node.Index = 1 Then  ' Si tratta del nodo Root
'      Set nodeX = tvReclami.Nodes("NA")
'      nodeX.Selected = True
'      tvReclami_NodeClick nodeX
   Else
      If Node.Key = "" Then      '  Nodo Foglia
         Node.Image = "A"
         If Not (LastNode Is Nothing) Then
            If Node <> LastNode Then
               LastNode.Image = "C"
            End If
         End If
         Set LastNode = Node
'*********************************
'
'  l'Sql DEVE dipendere dal parametro (ATC, MKG) !!!!!!!!!!!
'
'*********************************
         If Right$(Node.Parent.Key, 1) = "A" Then
'         If Me.Param = "ATC" Then
            sqlc = "Select rv.ODL, DataReclamo, Cliente, Prodotto, " & _
                     "Qta, QtaReclamata, " & _
                     "Stato, Soluzione " & _
                     "FROM (MFG INNER JOIN ReclamoVenditore AS rv ON MFG.ODL = rv.ODL) " & _
                     "LEFT JOIN ATC ON MFG.ODL = ATC.ODL " & _
                     "Where Stato = '" & Left$(Trim$(Node.Parent.Key), 1) & "' " & _
                     "and Venditore = '" & Trim$(Node.Text) & "' " & _
                     "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
            sNodeType = "ATC"
         Else
            'SQLC per MKG
            sqlc = "SELECT rv.ODL, ATC.DataRisposta, " & _
                     "MFG.Cliente, MFG.Prodotto, " & _
                     "Qta, QtaReclamata, " & _
                     "ATC.StatoReclamo AS Stato, Soluzione " & _
                     "FROM (ATC INNER JOIN MFG ON ATC.ODL = MFG.ODL) " & _
                     "INNER JOIN ReclamoVenditore AS rv ON ATC.ODL = rv.ODL " & _
                     "Where ATC.StatoReclamo = '" & Left$(Trim$(Node.Parent.Key), 1) & "' " & _
                     "and Venditore = '" & Trim$(Node.Text) & "' " & _
                     "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
          sNodeType = "MKG"
         End If
'                  "order by DataReclamo desc "
'  Ora è ordinato per odl in modo tale da poter digitare l'odl
'


      Else
         If Node.Key <> "ATC" And Node.Key <> "MKG" Then
            If Right$(Node.Key, 1) = "A" Then
               sqlc = "Select rv.ODL, DataReclamo, Cliente, Prodotto, " & _
                                 "Qta, QtaReclamata, Stato, ATC.Soluzione  " & _
                        "FROM (MFG INNER JOIN ReclamoVenditore AS rv ON MFG.ODL = rv.ODL) " & _
                        "LEFT JOIN ATC ON MFG.ODL = ATC.ODL " & _
                       "Where Stato = '" & Left(Trim$(Node.Key), 1) & "' " & _
                        "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
             sNodeType = "ATC"
            Else
               sqlc = "SELECT rv.ODL, ATC.DataRisposta, " & _
                        "MFG.Cliente, MFG.Prodotto, " & _
                        "Qta, QtaReclamata, " & _
                        "ATC.StatoReclamo AS Stato, Soluzione " & _
                        "FROM (ATC INNER JOIN MFG ON ATC.ODL = MFG.ODL) " & _
                        "INNER JOIN ReclamoVenditore AS rv ON ATC.ODL = rv.ODL " & _
                        "Where ATC.StatoReclamo = '" & Left(Trim$(Node.Key), 1) & "' " & _
                        "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
             sNodeType = "MKG"
            End If
         Else
            sqlc = ""
         End If
      End If

      With lvReclami
         If sqlc <> "" Then
            Set rsLV = db.OpenRecordset(sqlc, dbOpenDynaset)
            Do While Not rsLV.EOF
               Select Case rsLV("Stato")
                  Case "C"
                     iImg = 1
                  Case "N"
                     iImg = 2
                  Case "V"
                     iImg = 3
               End Select
' Carico una icona differente per le segnalazioni
   
               If rsLV("Soluzione") = 2 Then
                   iImg = 4
               End If
'**************************************************************
               Set itmX = .ListItems.Add(, , rsLV(0), , iImg)
               With itmX
                  For i = 1 To rsLV.Fields.Count - 3
                     If Not IsNull(rsLV(i)) Then
                        itmX.SubItems(i) = CStr(rsLV(i))
                     Else
                        itmX.SubItems(i) = " "
                     End If
                  Next i
               .Ghosted = False
               
               End With
               rsLV.MoveNext
            Loop
            rsLV.Close
            Set rsLV = Nothing
         End If
      End With
   End If
End Sub








