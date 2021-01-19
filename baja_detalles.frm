VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form21 
   BackColor       =   &H000040C0&
   Caption         =   "Eliminar detalles de PC"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13395
   LinkTopic       =   "Form21"
   ScaleHeight     =   7185
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Modificaciones"
      Height          =   6975
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar detalle"
         Height          =   735
         Left            =   600
         TabIndex        =   19
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Busqueda PC"
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   3015
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1080
            TabIndex        =   13
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackColor       =   &H0080C0FF&
            Caption         =   "Laboratorio:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            Caption         =   "Nº PC:"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Datos Clave"
         Height          =   3615
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   3015
         Begin VB.Label Label14 
            BackColor       =   &H0080C0FF&
            Caption         =   "ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   3120
            Width           =   2775
         End
         Begin VB.Label Label12 
            BackColor       =   &H0080C0FF&
            Caption         =   "Detalle:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackColor       =   &H0080C0FF&
            Caption         =   "Descripcion:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label7 
            BackColor       =   &H0080C0FF&
            Caption         =   "Dispositivo:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Numero PC:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080C0FF&
            Caption         =   "Nombre Labo:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   6975
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin MSComctlLib.ListView ListView1 
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   11456
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
            Text            =   "cod"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
If Trim(Combo2.Text) <> "" Then
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select labos.nomb, pcs.num_pc, detalle_dispositivos.nombre, dispositivos.nombrehw, descripcion.descrip, descripcion.id FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(5) = !id
.MoveNext
Next
End With
busca.Close
End If

End Sub

Private Sub Command1_Click()
Set delet = Nothing
If Label8.Caption = "" Or Label5.Caption = "" Or Label10.Caption = "" Or Label3.Caption = "" Then
MsgBox ("Seleccione un elemento para modificar")
Exit Sub
End If

If MsgBox("Esta seguro de modificar esta Descripcion!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If


Dim dele As ADODB.Recordset
Set delet = New ADODB.Recordset
With delet
.Open "select descripcion.id FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where pcs.id_lab = labos.id and pcs.idpc = descripcion.id_pc and detalle_dispositivos.cod_detalle = descripcion.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and descripcion.descrip LIKE'" & (Label1.Caption) & "'and dispositivos.nombrehw LIKE'" & (Label8.Caption) & "'  and labos.nomb LIKE'" & (Label5.Caption) & "' and pcs.num_pc LIKE'" & (Label3.Caption) & "' and descripcion.id LIKE'" & (Label13.Caption) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = delet.RecordCount


Do Until delet.EOF
   
         delet.Delete
  
    
      delet.MoveNext
   Loop

delet.Close

End With




Label8.Caption = ""
Label5.Caption = ""
Label10.Caption = ""
Label1.Caption = ""
Label13.Caption = ""
Label3.Caption = ""

ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select labos.nomb, pcs.num_pc, detalle_dispositivos.nombre, dispositivos.nombrehw, descripcion.descrip, descripcion.id FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(5) = !id
.MoveNext
Next
End With
busca.Close

End Sub


Private Sub ListView1_DblClick()
Label5.Caption = ListView1.selectedItem.Text 'cod
Label3.Caption = ListView1.selectedItem.SubItems(1) 'pc
Label8.Caption = ListView1.selectedItem.SubItems(2) 'dispositivo
'Text2.Text = ListView1.SelectedItem.SubItems(3) 'descripcion
Label10.Caption = ListView1.selectedItem.SubItems(3) 'descripcion
Label1.Caption = ListView1.selectedItem.SubItems(4) 'detalle
Label13.Caption = ListView1.selectedItem.SubItems(5) 'detalle
End Sub


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
If Trim(Combo1.Text) <> "" Then

Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select labos.nomb, pcs.num_pc, detalle_dispositivos.nombre, dispositivos.nombrehw, descripcion.descrip, descripcion.id FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nuevo.SubItems(5) = !id
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
