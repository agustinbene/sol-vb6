VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form15 
   BackColor       =   &H000040C0&
   Caption         =   "Bajas PC"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form15"
   ScaleHeight     =   6225
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar PC"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   2295
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
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Nº PC:"
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
      Top             =   120
      Width           =   9015
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
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
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Combo1.Text = ""
End Sub
Private Sub Combo2_Change()
Combo2.Text = ""
End Sub

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

Private Sub Combo2_dropdown()
Combo2.Clear
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
End Sub


Private Sub Command2_Click()


If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
Exit Sub
End If
If MsgBox("Esta seguro de eliminar esta Computadora!? Se eliminaram todos los campos referidas a ellas incluido el historial de reparaciones", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
.Open "select * FROM labos where labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
Label3.Caption = !ID
.Close
'Busqueda para ver si ya existe el laboratorio
.Open "delete pcs from pcs where pcs.id_lab LIKE'" & Trim(Label3.Caption) & "%' and pcs.num_pc LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("PC: " & Combo2.Text & "del laboratorio: " & Combo1.Text & " Eliminado")


End With
End Sub


