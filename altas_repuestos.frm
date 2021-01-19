VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BackColor       =   &H000040C0&
   Caption         =   "Altas Repuestos"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form7"
   ScaleHeight     =   5910
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Agregar Repuesto"
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000008&
         Caption         =   "Agregar Repuesto"
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   1200
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   600
         TabIndex        =   3
         Top             =   2040
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   1623
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre repuesto:"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Combo1_dropdown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM repuestos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo1.AddItem !descrip
busca.MoveNext
Next


End With
busca.Close
End Sub

Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Then
Combo1.SetFocus
Exit Sub
End If

cant_vendida = InputBox("Cantidad a usar de " & UCase(Combo1.Text))
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

Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el repuesto
.Open "Select * from repuestos where repuestos.descrip LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!descrip = UCase(Combo1.Text)


.Update
codrepu = !cod_repu

.Close


'agregar stock
.Open "SELECT * FROM stock", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
   !cant = cant_vendida
   !cod_stock = codrepu
   .Update
   .Close







Else
MsgBox ("El Repuesto: '" & Combo1.Text & "' Ya Existe!")
.Close

End If
End With
Combo1.Clear
Combo1.SetFocus

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
End With
busca.Close


End Sub

Private Sub Form_Load()


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





