VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form12 
   BackColor       =   &H000040C0&
   Caption         =   "Modificacion de descripcion"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form12"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Modificaciones"
      Height          =   6975
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.Frame Frame5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Datos Clave"
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   3015
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080C0FF&
            Caption         =   "Nombre Labo:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Numero PC:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label7 
            BackColor       =   &H0080C0FF&
            Caption         =   "Dispositivo:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackColor       =   &H0080C0FF&
            Caption         =   "Descripcion:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Busqueda PC"
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3015
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            Caption         =   "N� PC:"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H0080C0FF&
            Caption         =   "Laboratorio:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Detalle a modificar"
         Height          =   1815
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   3015
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "MODIFICAR"
            Height          =   375
            Left            =   600
            TabIndex        =   19
            Top             =   1080
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   6975
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSComctlLib.ListView ListView1 
         Height          =   6495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Labo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PC�"
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
End
Attribute VB_Name = "Form12"
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


Private Sub Combo4_dropdown()
Combo4.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select DISTINCT descrip from detalle_dispositivos,dispositivos,descripcion where descripcion.cod_detalle = detalle_dispositivos.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & Trim(Label8.Caption) & "%' and detalle_dispositivos.nombre LIKE'" & Trim(Label10.Caption) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
    
    
    Combo4.AddItem !descrip
 
            
                busca.MoveNext

  
   Next




End With
busca.Close
End Sub



Private Sub Command1_Click()
If Label8.Caption = "" Or Label5.Caption = "" Or Label10.Caption = "" Or Label3.Caption = "" Then
MsgBox ("Seleccione un elemento para modificar")
Exit Sub
End If

If MsgBox("Esta seguro de modificar esta Descripcion!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If


Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
altas.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where pcs.id_lab = labos.id and pcs.idpc = descripcion.id_pc and detalle_dispositivos.cod_detalle = descripcion.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & (Label8.Caption) & "'  and labos.nomb LIKE'" & (Label5.Caption) & "' and pcs.num_pc LIKE'" & (Label3.Caption) & "' and detalle_dispositivos.nombre LIKE'" & (Label10.Caption) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
'altas.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where pcs.id_lab = labos.id and pcs.idpc = descripcion.id_pc and detalle_dispositivos.cod_detalle = descripcion.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & (Label8.Caption) & "'  and labos.nomb LIKE'" & (Label5.Caption) & "' and pcs.num_pc LIKE'" & (Label3.Caption) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

!descrip = UCase(Combo4.Text)
'!nombre = UCase(Text2.Text)

.Update
End With

Combo4.Text = ""
Label8.Caption = ""
Label5.Caption = ""
Label10.Caption = ""
Label3.Caption = ""


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
Frame3.Enabled = False
End Sub





Private Sub Form_Load()
Frame3.Enabled = False
End Sub

Private Sub ListView1_DblClick()
Frame3.Enabled = True
Label5.Caption = ListView1.selectedItem.Text 'cod
Label3.Caption = ListView1.selectedItem.SubItems(1) 'pc
Label8.Caption = ListView1.selectedItem.SubItems(2) 'dispositivo
'Text2.Text = ListView1.SelectedItem.SubItems(3) 'descripcion
Label10.Caption = ListView1.selectedItem.SubItems(3) 'descripcion
Combo4.Text = ListView1.selectedItem.SubItems(4) 'detalle
End Sub
