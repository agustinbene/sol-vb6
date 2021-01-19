VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000040C0&
   Caption         =   "Busqueda por componentes"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form6"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Dispositivo:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Datos"
      Height          =   6015
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8438015
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Labo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PC°"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dispositivo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "detalle"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda_componentes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda_componentes.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda_componentes.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda_componentes.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DEclaración de la Función Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Const WM_SETREDRAW As Long = &HB&




Private Sub Combo1_Click()
Combo2.Text = ""
Combo3.Text = ""

End Sub

'-----------------------------------

Private Sub Combo1_dropdown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM dispositivos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = busca.RecordCount
If W < 1 Then
Exit Sub
End If
busca.MoveFirst
    If .EOF Then
     Exit Sub
    End If
For i = 1 To W
If .EOF Then
Exit For
End If
Combo1.AddItem !nombrehw
busca.MoveNext
Next


End With
busca.Close
End Sub

Private Sub Combo2_dropdown()
Combo2.Clear
If Trim(Combo1.Text) <> "" Then
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from detalle_dispositivos,dispositivos where dispositivos.cod = detalle_dispositivos.cod_dispo and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = busca.RecordCount
If W < 1 Then
Exit Sub
End If
busca.MoveFirst
    If .EOF Then
     Exit Sub
    End If
For i = 1 To W
If .EOF Then
Exit For
End If
Combo2.AddItem !nombre
busca.MoveNext
Next
End With
busca.Close

Dim r As Integer, t As Integer
For r = 0 To Combo2.ListCount - 1
For t = 0 To Combo2.ListCount - 1
If (Combo2.List(r) = Combo2.List(t)) And t <> r Then
Combo2.RemoveItem (r)
End If
Next t
Next r


End If
End Sub

Private Sub Combo3_dropdown()
Combo3.Clear
If Trim(Combo1.Text) <> "" And Trim(Combo2.Text) <> "" Then
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from detalle_dispositivos,dispositivos,descripcion where descripcion.cod_detalle = detalle_dispositivos.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%' and detalle_dispositivos.nombre LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = busca.RecordCount
If W < 1 Then
Exit Sub
End If
busca.MoveFirst
    If .EOF Then
     Exit Sub
    End If
    
    For x = 1 To W
     If .EOF Then
     Exit Sub
    End If
    
    
    Combo3.AddItem !descrip
 
            
                busca.MoveNext

  
   Next
Dim q As Integer, e As Integer
For q = 0 To Combo3.ListCount - 1
For e = 0 To Combo3.ListCount - 1
If (Combo3.List(q) = Combo3.List(e)) And e <> q Then
Combo3.RemoveItem (q)
End If
Next e
Next q
For q = 0 To Combo3.ListCount - 1
For e = 0 To Combo3.ListCount - 1
If (Combo3.List(q) = Combo3.List(e)) And e <> q Then
Combo3.RemoveItem (q)
End If
Next e
Next q




End With
busca.Close
End If
End Sub

Private Sub Command1_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where pcs.id_lab = labos.id and pcs.idpc = descripcion.id_pc and detalle_dispositivos.cod_detalle = descripcion.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & (Combo1.Text) & "' and detalle_dispositivos.nombre LIKE'" & (Combo2.Text) & "' and descripcion.descrip LIKE'" & (Combo3.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = busca.RecordCount

With busca

If W <= 0 Then
ListView1.ListItems.Clear
Exit Sub
End If
busca.MoveFirst
ListView1.ListItems.Clear

If .EOF Then
Exit Sub
End If
For c = 1 To W
If .EOF Then
Exit For
End If
busca.Update

Set nuevo = ListView1.ListItems.Add(, , !nomb)
nuevo.SubItems(1) = !num_pc
nuevo.SubItems(2) = !nombrehw
nuevo.SubItems(3) = !nombre
nuevo.SubItems(4) = !descrip
.MoveNext
Next
End With
busca.Close

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "menu"
    Form1.Show

End Select
End Sub

'****************************************************************
' Evento al hacer clic en la columna
'----------------------------------------------------------------

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    On Error Resume Next
    
    
    With ListView1
    
        Dim i As Long
        Dim Formato As String
        Dim strData() As String
        
        Dim Columna As Long
        
        Call SendMessage(Me.hwnd, WM_SETREDRAW, 0&, 0&)
        
        
        Columna = ColumnHeader.Index - 1
        
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Tipo de dato a ordenar
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        Select Case UCase$(ColumnHeader.Tag)
    
        
        ' Fecha
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "DATE"
        
            Formato = "YYYYMMDDHhNnSs"
        
            ' Ordena alfabéticamente la columna con Fechas _
              ( es la columna que tiene en el tag el valor DATE )
        
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
            
        ' Datos de numéricos
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "NUMBER"
        
            ' Ordena alfabéticamente la columna con números _
              ( es la columna que tiene en el tag el valor NUMBER )
        
            Formato = String(30, "0") & "." & String(30, "0")
                
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
        
        Case Else
                    
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
    
    End With
    
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 1&, 0&)
    ListView1.Refresh
    
End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

