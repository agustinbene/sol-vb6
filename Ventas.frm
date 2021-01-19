VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form5"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cambiar Cliente"
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   7095
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Venta"
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   6480
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6015
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   10610
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   6480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5295
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   3840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "COD"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   3960
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "COD"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   4516
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Filtrar Producto por: "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda Cliente"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Apellido"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "COD"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Apellido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Telefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Direccion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Filtrar Cliente por: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo_del_cliente As Integer
Dim codigo_del_articulo As String
Dim numero_de_factura As Integer
Dim cantidad As Integer
Dim el_precio As Double
Dim i1, i2, i3 As Single
Dim hay As Integer
Dim Codproveedor As Single
Dim total As Double



Private Sub Command1_Click()
If ListView2.ListItems.Count = 0 Then
  Exit Sub
 End If
 
 Dim venta As ADODB.Recordset
  Set venta = New ADODB.Recordset
 
 
  With venta
   .Open "SELECT * FROM facturas", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
       .AddNew
       !fecha = Date
       !hora = Time
       !cod_cli = codigo_del_cliente
       
       .Update
    mi_id_venta = !cod_fact
        .Close

  For i = 1 To (ListView2.ListItems.Count)
    .Open "SELECT * FROM articulos where cod_artic like'" & ListView2.ListItems(i).SubItems(2) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
    hay = !stock
   !stock = hay - (ListView2.ListItems(i).SubItems(1))
   .Update
   .Close
   

   
     .Open "Select * from ventas ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
    .AddNew
    !cod_artic = ListView2.ListItems(i).SubItems(2)
    !cod_venta = mi_id_venta
    !cant = ListView2.ListItems(i).SubItems(1)
    !precio = ListView2.ListItems(i).SubItems(3)
    .Update
    .Close
     

   Next
   
     
MsgBox ("VENTA REALIZADA CON EXITO. Total:   $" & total)
ListView2.ListItems.Clear
   End With
   

   total = 0
 Label4.Caption = ("TOTAL: ")
   End Sub






Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub Command3_Click()
Frame1.Enabled = True
Command3.Visible = False
ListView1.ListItems.Clear
Label2.Caption = ""
Text1.SetFocus
End Sub

Private Sub Command4_Click()
Form7.Show

End Sub

Private Sub ListView3_DblClick()
codigo_del_cliente = ListView3.SelectedItem.Text
Label25.Caption = ("Cliente:  " & ListView3.SelectedItem.SubItems(1))
Frame2.Enabled = True
Frame1.Enabled = False
Command3.Visible = True
Form5.Text1.Text = Me.Text1.Text
Form7.Hide
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
ListView3.ListItems.Clear
Exit Sub
End If
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM clientes where apel LIKE'" & Trim(Text1) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount
If w < 1 Then
ListView3.ListItems.Clear
Exit Sub
End If
busca.MoveFirst
ListView3.ListItems.Clear

With busca
If .EOF Then
Exit Sub
End If

For i = 1 To w
If .EOF Then
Exit For
End If
Set nuevo = ListView3.ListItems.Add(, , !cod_cli)
nuevo.SubItems(1) = !apel
nuevo.SubItems(2) = !direc
.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
ListView1.ListItems.Clear
Exit Sub
End If
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM articulos where descrip LIKE'" & Trim(Text2) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount
If w < 1 Then
ListView1.ListItems.Clear
Exit Sub
End If
busca.MoveFirst
ListView1.ListItems.Clear

With busca
If .EOF Then
Exit Sub
End If

For i = 1 To w
If .EOF Then
Exit For
End If
Set nuevo = ListView1.ListItems.Add(, , !cod_artic)
nuevo.SubItems(1) = !descrip
nuevo.SubItems(2) = !stock
nuevo.SubItems(3) = !precio
.MoveNext
Next
End With
busca.Close




End Sub

Private Sub Form_load()


Frame2.Enabled = False
Command3.Visible = False
End Sub






Private Sub ListView1_DblClick()

i1 = ListView1.SelectedItem.Text 'cod
 i2 = ListView1.SelectedItem.SubItems(1) 'descripcion
 i3 = ListView1.SelectedItem.SubItems(2) 'stock
 el_precio = ListView1.SelectedItem.SubItems(3)
  
For n = 1 To (ListView2.ListItems.Count)
    If i1 = (ListView2.ListItems(n).SubItems(2)) Then
        cant_vendida = InputBox("Cantidad a Agregar de " & i2)
            If cant_vendida > i3 Then
             MsgBox "no puede retirar mas de lo que tiene"
                Exit Sub
            End If
        ListView2.ListItems(n).SubItems(1) = Val(ListView2.ListItems(n).SubItems(1)) + Val(cant_vendida)
        cantidad = cant_vendida
            For i = 1 To (ListView2.ListItems.Count)
                total = total + (ListView2.ListItems(i).SubItems(3) * ListView2.ListItems(i).SubItems(1))
                Label4.Caption = ("TOTAL: " & total)
            Next
        Exit Sub
    End If
Next
  
cant_vendida = InputBox("Cantidad a Comprar de " & i2)
If cant_vendida > i3 Then
MsgBox "no puede retirar mas de lo que tiene"
Exit Sub
End If

Set nuevo = ListView2.ListItems.Add(, , i2)

nuevo.SubItems(1) = cant_vendida
nuevo.SubItems(2) = i1
nuevo.SubItems(3) = el_precio
codigo_del_articulo = ListView1.SelectedItem.Text
cantidad = cant_vendida
mi_id = i1

For i = 1 To (ListView2.ListItems.Count)
total = total + (ListView2.ListItems(i).SubItems(3) * ListView2.ListItems(i).SubItems(1))
Label4.Caption = ("TOTAL: " & total)
Next



End Sub



Private Sub ListView2_DblClick()
ListView2.ListItems.Remove ListView2.SelectedItem.Index
total = 0
For i = 1 To (ListView2.ListItems.Count)
total = total + (ListView2.ListItems(i).SubItems(3) * ListView2.ListItems(i).SubItems(1))
Label4.Caption = ("TOTAL: " & total)
Next
End Sub



Private Sub Option1_Click()
Text1.SetFocus
Text1.Text = ""
End Sub

Private Sub Option2_Click()
Text1.SetFocus
Text1.Text = ""
End Sub

Private Sub Option3_Click()
Text2.SetFocus
Text2.Text = ""
End Sub

Private Sub Option4_Click()
Text2.SetFocus
Text2.Text = ""
End Sub



