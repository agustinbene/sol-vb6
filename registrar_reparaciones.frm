VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BackColor       =   &H000040C0&
   Caption         =   "Registrar Reparacion"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form8"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Limpiar datos "
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000080FF&
      Caption         =   "Profesor a cargo"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Siguiente"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Caption         =   "Descripcion Reparacion"
      Height          =   6015
      Left            =   3840
      TabIndex        =   8
      Top             =   480
      Width           =   9255
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "Registrar Reparacion"
         Height          =   495
         Left            =   120
         MaskColor       =   &H000000FF&
         TabIndex        =   14
         Top             =   5280
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8705
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PC°"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dispositivo"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Profesor"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Lista Repuestos"
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   3375
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   8438015
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Stock"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Definir PC"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "Siguiente"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "PC:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Laboratorio:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_profee As Single
Dim codpc As Single
Dim numeropc As Single
Dim nombrelabo As String


Dim stock As Single




Private Sub Combo1_Change()
Combo1.Text = ""
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

Private Sub Combo2_Change()
Combo2.Text = ""
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




Private Sub Combo3_dropdown()
Combo3.Clear
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
Combo3.AddItem !nombre
busca.MoveNext
Next


End With
busca.Close
End Sub

Private Sub Command2_Click()
If Trim(Combo3.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el Profesor
.Open "Select * from profes where profes.nombre LIKE'" & Trim(Combo3.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!nombre = UCase(Combo3.Text)
.Update

MsgBox ("Profesor: " & Combo3.Text & " Agregado")
End If
id_profee = !id_profe
altas.Close
End With
Frame1.Enabled = True
Frame4.Enabled = False
End Sub

Private Sub Command3_Click()

Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo1.Text) & "' and pcs.num_pc LIKE'" & Trim(Combo2.Text) & "'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
nombrelabo = !nomb
numeropc = !num_pc
codpc = !idpc
Label3.Caption = ("cod: " & codpc & " Num: " & numeropc & " Labo: " & nombrelabo)
busca.MoveNext
Next
End With
Frame1.Enabled = False
Frame2.Enabled = True
End Sub



Private Sub Command4_Click()
If ListView1.ListItems.Count = 0 Then
  Exit Sub
 End If
 
 Dim repa As ADODB.Recordset
  Set repa = New ADODB.Recordset
 
 
  With repa
    

  For i = 1 To (ListView1.ListItems.Count)
    .Open "SELECT * FROM repuestos,stock where repuestos.cod_repu = stock.cod_stock and repuestos.descrip like'" & ListView1.ListItems(i).SubItems(2) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
   If ListView1.ListItems(i).SubItems(3) <= (!cant) Then
   !cant = (!cant - ListView1.ListItems(i).SubItems(3))
   Else
   MsgBox "no puede retirar mas de lo que tiene"
   .Close
   Exit Sub
   End If
   .Update
   repu_usado = !cod_repu
   .Close
          

 .Open "SELECT * FROM desc_reparaciones", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
       .AddNew
       !fecha = Date
    !hora = Time
       !id_pc = codpc
       !id_profe = id_profee
       
       .Update
       
 id_reparacion = !id_repa
        .Close
        
.Open "SELECT * FROM reparaciones", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
       .AddNew
       !cod_repara = id_reparacion
       !cant_usada = ListView1.ListItems(i).SubItems(3)
       !cod_repuesto = repu_usado
       .Update

        .Close
     

   Next
   

   End With
   
   'Busqueda y muestreo de los repuestops existentes
   ListView2.ListItems.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from stock,repuestos where repuestos.cod_repu = stock.cod_stock", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView2.ListItems.Add(, , !descrip)
nuevo.SubItems(1) = !cant
busca.MoveNext
Next

busca.Close

busca.Open "select * from stock,repuestos where repuestos.cod_repu = stock.cod_stock", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView2.ListItems.Add(, , !descrip)
nuevo.SubItems(1) = !cant
busca.MoveNext
Next
End With
busca.Close


   
ListView1.ListItems.Clear
End Sub



Private Sub Command5_Click()
ListView1.ListItems.Clear

Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Label3.Caption = ""
Frame4.Enabled = True
Frame1.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub Form_Load()
Frame2.Enabled = False
Frame3.Enabled = False
Frame1.Enabled = False
'Busqueda y muestreo de los repuestops existentes
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from stock,repuestos where repuestos.cod_repu = stock.cod_stock", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView2.ListItems.Add(, , !descrip)
nuevo.SubItems(1) = !cant
busca.MoveNext
Next
End With
busca.Close
End Sub







Private Sub Label4_Click()
Label4.Caption = ListView1.ListItems(1).SubItems(2)
End Sub

Private Sub ListView1_DblClick()
ListView1.ListItems.Remove ListView1.selectedItem.Index
End Sub

Private Sub ListView2_DblClick()
cant_vendida = InputBox("Cantidad a usar de " & repuestoo)
Text1.Text = cant_vendida
If IsNumeric(cant_vendida) = False Then
MsgBox "Solo deben ingresarse numeros enteros", vbOKOnly
Exit Sub
End If
If cant_vendida = "" Or cant_vendida < 0 Then
MsgBox "Ingrese un numero mayor a 0", vbOKOnly
Exit Sub
End If
For s = 1 To Len(Text1.Text)
If Mid(Text1, s, 1) = "." Or Mid(Text1, s, 1) = "," Then
MsgBox "Solo deben ingresarse numeros enteros", vbOKOnly
Exit Sub
End If
Next

Frame3.Enabled = True
repuestoo = ListView2.selectedItem.Text
stock = ListView2.selectedItem.SubItems(1)
  

If cant_vendida > stock Then
MsgBox "no puede retirar mas de lo que tiene"
Exit Sub
End If

Set nuevo = ListView1.ListItems.Add(, , UCase(Combo1.Text))

nuevo.SubItems(1) = Combo2.Text
nuevo.SubItems(2) = repuestoo
nuevo.SubItems(3) = cant_vendida
nuevo.SubItems(4) = UCase(Combo3.Text)



End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "menu"
    Form1.Show
End Select
End Sub

