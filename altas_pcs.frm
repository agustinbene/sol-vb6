VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form4 
   BackColor       =   &H000080FF&
   Caption         =   "Agregar PC"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form3"
   ScaleHeight     =   7320
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos PC nueva"
      ForeColor       =   &H00000000&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13335
      Begin VB.Frame Frame6 
         BackColor       =   &H000080FF&
         Caption         =   "Copiar datos "
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   6000
         Width           =   12735
         Begin VB.ComboBox Combo8 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   2760
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox Combo7 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   480
            TabIndex        =   29
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   7200
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Copiar"
            Height          =   375
            Left            =   11520
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox Combo5 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   9840
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H000080FF&
            Caption         =   "COPIAR A ------>"
            Height          =   255
            Left            =   4440
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackColor       =   &H000080FF&
            Caption         =   "PC:"
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            BackColor       =   &H000080FF&
            Caption         =   "LAB:"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label7 
            BackColor       =   &H000080FF&
            Caption         =   "LAB:"
            Height          =   375
            Left            =   6840
            TabIndex        =   27
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            Caption         =   "PC:"
            Height          =   255
            Left            =   9480
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   4200
         TabIndex        =   17
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
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
      Begin VB.Frame Frame3 
         BackColor       =   &H000080FF&
         Caption         =   "Detallar dispositivos"
         ForeColor       =   &H00000000&
         Height          =   4335
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   3855
         Begin VB.Frame Frame4 
            BackColor       =   &H000080FF&
            Caption         =   "Detalles"
            ForeColor       =   &H00000000&
            Height          =   3375
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   3615
            Begin VB.Frame Frame5 
               BackColor       =   &H000080FF&
               Caption         =   "Descrip"
               ForeColor       =   &H00000000&
               Height          =   2415
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   3375
               Begin VB.ComboBox Combo4 
                  BackColor       =   &H0080C0FF&
                  Height          =   315
                  Left            =   600
                  TabIndex        =   5
                  Top             =   600
                  Width           =   2175
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Agregar descripcion"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   6
                  Top             =   1440
                  Width           =   3015
               End
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H0080C0FF&
               Height          =   315
               ItemData        =   "altas_pcs.frx":0000
               Left            =   840
               List            =   "altas_pcs.frx":0002
               TabIndex        =   4
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label4 
               BackColor       =   &H000080FF&
               Caption         =   "Detalle:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H0080C0FF&
            Height          =   315
            ItemData        =   "altas_pcs.frx":0004
            Left            =   960
            List            =   "altas_pcs.frx":0006
            TabIndex        =   3
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackColor       =   &H000080FF&
            Caption         =   "Tipo:"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000080FF&
         Caption         =   "Laboratorio - N°PC"
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3855
         Begin VB.ComboBox Text1 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Buscar"
            Height          =   255
            Left            =   3120
            TabIndex        =   18
            Top             =   720
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H0080C0FF&
            Height          =   315
            ItemData        =   "altas_pcs.frx":0008
            Left            =   960
            List            =   "altas_pcs.frx":0015
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H000080FF&
            Caption         =   "LAB:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H000080FF&
            Caption         =   "N°PC:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Label Label14 
         Caption         =   "cod_detallee"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label12"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   5160
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   5520
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
            Picture         =   "altas_pcs.frx":0022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "altas_pcs.frx":08FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "altas_pcs.frx":11D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "altas_pcs.frx":1AB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label15 
      Caption         =   "Fuente:"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Fuente:"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Mother:"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Codproveedor As Single
Dim cod_detallee As Single
Dim cod_dispoo As Single
Dim idd As Single
Dim cont As Single
Dim labid As Integer








Private Sub Combo1_Click()
If Trim(Text1.Text) <> "" Then
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

Combo2.AddItem !nombrehw
nombrehard = !nombrehw
busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Combo3_dropdown()
Combo3.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from detalle_dispositivos,dispositivos where dispositivos.cod = detalle_dispositivos.cod_dispo and dispositivos.nombrehw LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo3.AddItem !nombre



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
.Open "select * from detalle_dispositivos,dispositivos,descripcion where descripcion.cod_detalle = detalle_dispositivos.cod_detalle and detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & Trim(Combo2.Text) & "%' and detalle_dispositivos.nombre LIKE'" & Trim(Combo3.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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


'--------------------------------FILTRO COMBO------------------------------------

Dim q As Integer, e As Integer
    For q = 0 To Combo4.ListCount - 1
        For e = 0 To Combo4.ListCount - 1
            If (Combo4.List(q) = Combo4.List(e)) And e <> q Then
        Combo4.RemoveItem (q)
            End If
        Next e
    Next q

    For q = 0 To Combo4.ListCount - 1
        For e = 0 To Combo4.ListCount - 1
            If (Combo4.List(q) = Combo4.List(e)) And e <> q Then
        Combo4.RemoveItem (q)
            End If
        Next e
    Next q
        For q = 0 To Combo4.ListCount - 1
        For e = 0 To Combo4.ListCount - 1
            If (Combo4.List(q) = Combo4.List(e)) And e <> q Then
        Combo4.RemoveItem (q)
            End If
        Next e
    Next q
        For q = 0 To Combo4.ListCount - 1
        For e = 0 To Combo4.ListCount - 1
            If (Combo4.List(q) = Combo4.List(e)) And e <> q Then
        Combo4.RemoveItem (q)
            End If
        Next e
    Next q
'--------------------------------FILTRO COMBO------------------------------------


End Sub









Private Sub Combo5_Change()
If IsNumeric(Combo5.Text) = False Then
Combo5.Text = ""
End If
End Sub

Private Sub Combo5_dropdown()
Combo5.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo6.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo5.AddItem !num_pc

busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Combo6_dropdown()
Combo6.Clear
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
Combo6.AddItem !nomb
busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Combo7_Change()
Combo7.Text = ""
End Sub
Private Sub Combo8_Change()
Combo8.Text = ""
End Sub

Private Sub Combo7_dropdown()
Combo7.Clear
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
Combo7.AddItem !nomb
busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Combo8_click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo7.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Combo8.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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

Private Sub Combo8_DropDown()
Combo8.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo7.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo8.AddItem !num_pc

busca.MoveNext
Next


End With
busca.Close

End Sub

Private Sub Command1_Click()
If Trim(Combo7.Text) = "" Or Trim(Combo8.Text) = "" Then
Exit Sub
End If

Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas


'Busqueda para ver si ya existe el laboratorio
.Open "Select * from labos where labos.nomb LIKE'" & Trim(Combo6.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!nomb = UCase(Combo6.Text)
.Update
id_labb = !id
Else
id_labb = !id
End If
altas.Close

'completar tabla pcs

'.Open "Select * from labos, pcs where labos.nomb LIKE'" & Trim(Combo6.Text) & "%' and pcs.id_lab = labos.id and pcs.num_pc LIKE'" & Trim(ListView1.ListItems(1).SubItems(1)) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.Open "Select * from labos, pcs where labos.nomb LIKE'" & Trim(Combo6.Text) & "%' and pcs.id_lab = labos.id and pcs.num_pc LIKE'" & Trim(Combo5.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Then 'si ya existe la pc en el labo no se crea otra
altas.Close
.Open "Select * from pcs ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!num_pc = (Combo5.Text)
!id_lab = id_labb
.Update
idd = !idpc
altas.Close
Else
idd = !idpc
altas.Close
    End If
End With



For hd = 1 To ListView1.ListItems.Count
With altas
.Open "Select * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(ListView1.ListItems(hd).SubItems(2)) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W > 0 Then
codd = !cod
Label12.Caption = codd
End If
.Close
.Open "select * from detalle_dispositivos,dispositivos where detalle_dispositivos.nombre LIKE'" & Trim(ListView1.ListItems(hd).SubItems(3)) & "%' and dispositivos.cod = detalle_dispositivos.cod_dispo and dispositivos.nombrehw LIKE'" & Trim(ListView1.ListItems(hd).SubItems(2)) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
wdet = altas.RecordCount
If wdet >= 1 Then 'si ya existe detalle no se crea otro
cod_detallee = !cod_detalle
Label13.Caption = cod_detallee
End If
altas.Close


'Final grabar tabla descripcion
.Open "select * FROM descripcion,pcs,labos where descripcion.cod_dispo LIKE'" & Trim(Label12.Caption) & "%' and descripcion.cod_detalle LIKE'" & Trim(Label13.Caption) & "%' and pcs.idpc = descripcion.id_pc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Trim(Combo5.Text) & "%' and labos.nomb LIKE'" & Trim(Combo6.Text) & "%' ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
wa = altas.RecordCount
If wa >= 1 Then
altas.Close
MsgBox ("REPETIDO")
Exit Sub
Else
altas.Close
.Open "Select * from descripcion ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!cod_detalle = cod_detallee
!cod_dispo = codd
!descrip = (ListView1.ListItems(hd).SubItems(4))
!id_pc = idd
.Update
altas.Close
End If
End With
Next
End Sub



Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
Text1.Text = ""
End If
End Sub

Private Sub text1_Click()
If Trim(Combo1.Text) <> "" Then
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


Private Sub text1_dropdown()
If Trim(Combo1.Text) <> "" Then
Text1.Clear
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
Text1.AddItem !num_pc

busca.MoveNext
Next


End With
busca.Close
End If
End Sub



Private Sub Command2_Click()
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Text1) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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




















Private Sub Command3_Click() 'Completar tabla descripcion
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Or Trim(Combo3.Text) = "" Or Trim(Combo4.Text) = "" Or Trim(Text1.Text) = "" Then
MsgBox ("No deje campos en blanco")
Exit Sub
End If


Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el dispositivo
.Open "Select * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W > 0 Then
codd = !cod
Label12.Caption = codd
Else
.AddNew
'se graba el nuevo dispositivo
!nombrehw = UCase(Combo2.Text)
.Update
codd = !cod
End If
altas.Close
.Open "select * from detalle_dispositivos,dispositivos where detalle_dispositivos.nombre LIKE'" & Trim(Combo3.Text) & "%' and dispositivos.cod = detalle_dispositivos.cod_dispo and dispositivos.nombrehw LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
wdet = altas.RecordCount
If wdet >= 1 Then 'si ya existe detalle no se crea otro
cod_detallee = !cod_detalle
Label13.Caption = cod_detallee
altas.Close
Else
altas.Close
.Open "Select * from detalle_dispositivos where detalle_dispositivos.nombre LIKE'" & Trim(Combo3.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!nombre = UCase(Combo3.Text)
!cod_dispo = codd
.Update

cod_detallee = !cod_detalle
Label13.Caption = cod_detallee
altas.Close
End If
.Open "Select * from detalle_dispositivos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.Update
.Close


'Busqueda para ver si ya existe el laboratorio
.Open "Select * from labos where labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!nomb = UCase(Combo1.Text)
.Update
id_labb = !id
Else
id_labb = !id
End If
altas.Close
'completar tabla pcs
.Open "Select * from labos, pcs where labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and pcs.id_lab = labos.id and pcs.num_pc LIKE'" & Trim(Text1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Then 'si ya existe la pc en el labo no se crea otra
altas.Close
.Open "Select * from pcs ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!num_pc = (Text1.Text)
!id_lab = id_labb
.Update
idd = !idpc
altas.Close
Else
idd = !idpc
altas.Close
    End If

'Final grabar tabla descripcion
.Open "select * FROM descripcion,pcs,labos where descripcion.cod_dispo LIKE'" & Trim(Label12.Caption) & "%' and descripcion.cod_detalle LIKE'" & Trim(Label13.Caption) & "%' and pcs.idpc = descripcion.id_pc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Trim(Text1.Text) & "%' and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
wa = altas.RecordCount
If wa >= 1 Then
altas.Close
MsgBox ("REPETIDO")
Else
altas.Close
.Open "Select * from descripcion ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!cod_detalle = cod_detallee
!cod_dispo = codd
!descrip = UCase(Combo4.Text)
!id_pc = idd
.Update
altas.Close
End If
End With







'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'------------------------------Busqueda Muestreo--------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Text1) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------






End Sub
























Private Sub text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    ListView1.ListItems.Clear
Dim nuevo As ListItem
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
busca.Open "select * FROM descripcion,pcs,labos,dispositivos,detalle_dispositivos where detalle_dispositivos.cod_detalle = descripcion.cod_detalle and dispositivos.cod = descripcion.cod_dispo and labos.nomb LIKE'" & Trim(Combo1.Text) & "%' and descripcion.descrip LIKE'" & Trim(Combo4.Text) & "%' and descripcion.id_pc = pcs.idpc and labos.id = pcs.id_lab and pcs.num_pc LIKE'" & Val(Text1) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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








Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.Key
 Case Is = "cli"
    Form2.Show
    Case Is = "prove"
    Form3.Show
    Case Is = "Artic"
    Form6.Show
    End Select

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "menu"
    Form1.Show
    Case "ventas"
    Form3.Show
End Select
End Sub

