VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form9 
   BackColor       =   &H000040C0&
   Caption         =   "Historial de reparaciones"
   ClientHeight    =   7215
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13545
   LinkTopic       =   "Form9"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda por profesor"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   2895
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H000080FF&
         Caption         =   "Profesor:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   2895
      Begin VB.PictureBox DTPicker2 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   2235
         TabIndex        =   15
         Top             =   1440
         Width           =   2295
      End
      Begin VB.PictureBox DTPicker1 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   2235
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "Fecha final"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "Fecha inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "Mostrar Todo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000007&
         Caption         =   "Buscar"
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00000080&
         TabIndex        =   2
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Nº PC:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Datos"
      Height          =   6495
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin MSComctlLib.ListView ListView1 
         Height          =   6135
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8438015
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Labo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PC°"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dispositivo"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Profesor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Hora"
            Object.Width           =   2646
         EndProperty
      End
   End
End
Attribute VB_Name = "Form9"
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
ListView1.ListItems.Clear
If Trim(Combo2.Text) <> "" Then
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes where labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.nomb LIKE'%" & Trim(Combo1.Text) & "' and labos.id = pcs.id_lab and pcs.num_pc LIKE'%" & Val(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close
End If
End Sub

Private Sub Combo2_Click()
ListView1.ListItems.Clear
If Trim(Combo1.Text) <> "" Then
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes where labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.nomb LIKE'%" & Trim(Combo1.Text) & "' and labos.id = pcs.id_lab and pcs.num_pc LIKE'%" & Val(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close
End If

End Sub

'-----------------------------------



Private Sub Combo9_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset

busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes WHERE labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.id = pcs.id_lab and profes.nombre LIKE'%" & Trim(Combo9.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close

End Sub

Private Sub Combo9_dropdown()
Combo9.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM profes", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo9.AddItem !nombre
busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Command1_Click()




Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset

With busca
If Trim(Combo9.Text) = "" Then
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes WHERE desc_reparaciones.fecha Between " & "# " & Format(DTPicker1.Value, "mm/dd/yyyy") & " # And # " & Format(DTPicker2.Value, "mm/dd/yyyy") & " # and labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.id = pcs.id_lab", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
Else
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes WHERE desc_reparaciones.fecha Between " & "# " & Format(DTPicker1.Value, "mm/dd/yyyy") & " # And # " & Format(DTPicker2.Value, "mm/dd/yyyy") & " # and labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.id = pcs.id_lab and profes.nombre LIKE'%" & Trim(Combo9.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
End If
'.Open "select num_factura, fecha, apel ,prec From factura_venta, clientes, articulos where fecha between#" & CDate(fech) & "# and #" & CDate(fech2) & "# and factura_venta.cod_cliente=clientes.cod_cli and articulos.prec ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

W = busca.RecordCount

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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close




End Sub





'----------------------------------

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


Dim q As Integer, e As Integer
For q = 0 To Combo2.ListCount - 1
For e = 0 To Combo2.ListCount - 1
If (Combo2.List(q) = Combo2.List(e)) And e <> q Then
Combo2.RemoveItem (q)
End If
Next e
Next q
End If
End Sub



Private Sub Command2_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes where labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.nomb LIKE'%" & Trim(Combo1.Text) & "' and labos.id = pcs.id_lab and pcs.num_pc LIKE'%" & Val(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close

End Sub

Private Sub Command3_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes where labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.id = pcs.id_lab", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(2) = !descrip
nuevo.SubItems(3) = !cant_usada
nuevo.SubItems(4) = !nombre
nuevo.SubItems(5) = !fecha
nuevo.SubItems(6) = !hora
.MoveNext
Next
End With
busca.Close

End Sub










    
       

        
   






























Private Sub DTPicker1_Change()
If DTPicker1.Year > DTPicker2.Year Then
DTPicker1.Year = Val(DTPicker2.Year)
DTPicker1.Month = Val(DTPicker2.Month)
DTPicker1.Day = Val(DTPicker2.Day)
ElseIf DTPicker1.Month > DTPicker2.Month Then
DTPicker1.Day = Val(DTPicker2.Day)
DTPicker1.Month = Val(DTPicker2.Month)
DTPicker1.Year = Val(DTPicker2.Year)
ElseIf DTPicker1.Day > DTPicker2.Day Then
DTPicker1.Month = Val(DTPicker2.Month)
DTPicker1.Day = Val(DTPicker2.Day)
DTPicker1.Year = Val(DTPicker2.Year)

End If
End Sub



Private Sub DTPicker2_Change()
If DTPicker2.Year < DTPicker1.Year Then
DTPicker2.Year = Val(DTPicker1.Year)
DTPicker2.Day = Val(DTPicker1.Day)
DTPicker2.Month = Val(DTPicker1.Month)
ElseIf DTPicker2.Month < DTPicker1.Month Then
DTPicker2.Day = Val(DTPicker1.Day)
DTPicker2.Month = Val(DTPicker1.Month)
DTPicker2.Year = Val(DTPicker1.Year)
ElseIf DTPicker2.Day < DTPicker1.Day Then
DTPicker2.Day = Val(DTPicker1.Day)
DTPicker2.Month = Val(DTPicker1.Month)
DTPicker2.Year = Val(DTPicker1.Year)
End If
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



