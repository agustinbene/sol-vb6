VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form2"
   ScaleHeight     =   17180.28
   ScaleMode       =   0  'User
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      Height          =   6015
      Left            =   3120
      TabIndex        =   6
      Top             =   1080
      Width           =   10215
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   7
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
         ForeColor       =   -2147483648
         BackColor       =   -2147483642
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
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
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
         Caption         =   "Nº PC:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   1588
      ButtonWidth     =   1296
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Menú"
            Key             =   "menu"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cli"
                  Text            =   "Cliente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "prove"
                  Text            =   "Proveedor"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "artic"
                  Text            =   "Articulo"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cli"
                  Text            =   "Cliente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "prove"
                  Text            =   "Proveedor"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "artic"
                  Text            =   "Articulo"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "busqueda.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "busqueda.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_DropDown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount
If w < 1 Then
Exit Sub
End If
busca.MoveFirst
    If .EOF Then
     Exit Sub
    End If
For i = 1 To w
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
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount

With busca

If w <= 0 Then
ListView1.ListItems.Clear
Exit Sub
End If
busca.MoveFirst
ListView1.ListItems.Clear

If .EOF Then
Exit Sub
End If
For c = 1 To w
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

Private Sub Combo2_DropDown()
Combo2.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount
If w < 1 Then
Exit Sub
End If
busca.MoveFirst
    If .EOF Then
     Exit Sub
    End If
For i = 1 To w
If .EOF Then
Exit For
End If
Combo2.AddItem !num_pc

busca.MoveNext
Next


End With
busca.Close
End Sub

Private Sub Command1_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
w = busca.RecordCount

With busca

If w <= 0 Then
ListView1.ListItems.Clear
Exit Sub
End If
busca.MoveFirst
ListView1.ListItems.Clear

If .EOF Then
Exit Sub
End If
For c = 1 To w
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
