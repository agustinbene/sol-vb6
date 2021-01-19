VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form12"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Modificaciones"
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MODIFICAR"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Dispositivo:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Numero PC:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre Labo:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busqueda"
      Height          =   4935
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4080
         TabIndex        =   16
         Top             =   4440
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   4440
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8454143
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
      Begin VB.Label Label11 
         Caption         =   "Nº PC:"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form12"
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
