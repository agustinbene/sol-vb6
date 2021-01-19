VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form9 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form9"
   ClientHeight    =   7215
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13545
   LinkTopic       =   "Form9"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nº PC:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      Height          =   6015
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   10455
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   65280
         BackColor       =   -2147483647
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   4440
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
            Picture         =   "Form9.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   1588
      ButtonWidth     =   1032
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
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
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Siguiente"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   4560
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_dropdown()
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

Private Sub Combo2_dropdown()
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

Private Sub Command2_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM desc_reparaciones,pcs,labos,reparaciones,repuestos,stock,profes where labos.id = pcs.id_lab and profes.id_profe = desc_reparaciones.id_profe and pcs.idpc = desc_reparaciones.id_pc and desc_reparaciones.id_repa = reparaciones.cod_repara and reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.cod_repu = stock.cod_stock and labos.nomb LIKE'%" & Trim(Combo1.Text) & "%' and labos.id = pcs.id_lab and pcs.num_pc LIKE'%" & Val(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
