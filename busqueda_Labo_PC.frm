VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form2 
   BackColor       =   &H000040C0&
   Caption         =   "Busqueda de PC"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form2"
   ScaleHeight     =   17180.28
   ScaleMode       =   0  'User
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda con lector"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Activar lector"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Caption         =   "Imprimir codigo de barra"
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   1455
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   12
         Text            =   "1234"
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton optSize 
         BackColor       =   &H000080FF&
         Caption         =   "Pequeño "
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   972
      End
      Begin VB.OptionButton optSize 
         BackColor       =   &H000080FF&
         Caption         =   "Mediano"
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   972
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Imprimir Codigo"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Datos"
      Height          =   6735
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   9255
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   11245
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
         NumItems        =   6
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   706
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Nº PC:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1560
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
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




Private Sub Combo1_click()
If Trim(Combo2.Text) <> "" Then
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Text1.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
End If
End Sub

'-----------------------------------


Private Sub Combo1_dropdown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo1.AddItem !nomb
busca.MoveNext
Next


End With
busca.Close
End Sub


Private Sub Combo2_Click()
If Trim(Combo1.Text) <> "" Then
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Text1.Text = !id_pc
.MoveNext
Next
End With
busca.Close
End If
End Sub

Private Sub Combo2_dropdown()
Combo2.Clear
If Trim(Combo1.Text) <> "" Then
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo2.AddItem !num_pc

busca.MoveNext
Next


End With
busca.Close
End If
End Sub

Private Sub Command1_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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


Private Sub Command2_Click()
Text2.SetFocus
Frame4.Caption = "Busqueda con lector - ACTIVADA"
End Sub


Private Sub Form_Load()
orden = 1
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


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.Key
 Case Is = "labo"
    Form3.Show
    Case Is = "pc"
    Form4.Show
    Case Is = "artic"
    Form4.Show
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "menu"
    Form1.Show
End Select
End Sub



Private Sub cmdPrint_Click()
    Printer.PaintPicture Picture1, 5000, 5000
    Printer.EndDoc
End Sub

Private Sub Form_Activate()
Frame4.Caption = "Busqueda con lector - DESACTIVADA"
    optSize(1) = 1

End Sub

Private Sub optSize_Click(Index As Integer)
    Picture1.ScaleMode = 3
    
    Select Case Index
    Case 0
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
    Case 1
        Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 10
    Case 2
        Picture1.Height = Picture1.Height * (3 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 14
    End Select


    Call Text1_Change

End Sub

Private Sub Text1_Change()
    
    Call DrawBarcode(Text1, Picture1)
    
    MinWidth = 2 * Text1.Left + Text1.Width
    pw = 2 * Picture1.Left + Picture1.Width
    fw = MinWidth
    'If pw > fw Then fw = pw
   ' Form2.Width = fw

End Sub



Private Sub text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    ListView1.ListItems.Clear
    Text1.Text = Text2.Text
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select pcs.idpc, labos.nomb, pcs.num_pc, detalle_dispositivos.nombre, dispositivos.nombrehw, descripcion.descrip, descripcion.id FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.idpc LIKE'" & Val(Text2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
Text2.Text = ""
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
nuevo.SubItems(5) = !idpc
.MoveNext
Next
End With
busca.Close

    End If

End Sub

Private Sub Text2_LostFocus()
Frame4.Caption = "Busqueda con lector - DESACTIVADA"
End Sub
