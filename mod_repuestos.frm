VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BackColor       =   &H000040C0&
   Caption         =   "Modificacion de repuestos y stock"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form10"
   ScaleHeight     =   5205
   ScaleMode       =   0  'User
   ScaleWidth      =   29095.02
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Modificaciones"
      Height          =   4935
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "MODIFICAR"
         Height          =   375
         Left            =   840
         MaskColor       =   &H000000FF&
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Stock"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "COD Repuesto"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   4440
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7223
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "COD"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         Caption         =   "Articulo:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   4440
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2760
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
            Picture         =   "mod_repuestos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mod_repuestos.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mod_repuestos.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mod_repuestos.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If MsgBox("Esta seguro de modificar este Articulo!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If


cant_vendida = Text2.Text
If Trim(cant_vendida) = "" Or cant_vendida < 0 Then
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
altas.Open "select * from stock,repuestos where repuestos.cod_repu = stock.cod_stock and repuestos.cod_repu LIKE'" & (Label5.Caption) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

!descrip = UCase(Text1.Text)
!cant = Text2.Text

.Update
End With

Text1.Text = ""
Text2.Text = ""
Label5.Caption = ""



'Busqueda y muestreo de los repuestops existentes
ListView1.ListItems.Clear
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
Set nuevo = ListView1.ListItems.Add(, , !cod_repu)
nuevo.SubItems(1) = !descrip
nuevo.SubItems(2) = !cant
busca.MoveNext
Next
End With
busca.Close






Frame1.Enabled = False
End Sub

Private Sub Form_Load()
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
Set nuevo = ListView1.ListItems.Add(, , !cod_repu)
nuevo.SubItems(1) = !descrip
nuevo.SubItems(2) = !cant
busca.MoveNext
Next
End With
busca.Close
End Sub



Private Sub ListView1_DblClick()
Frame1.Enabled = True
Label5.Caption = ListView1.selectedItem.Text 'cod
Text1.Text = ListView1.selectedItem.SubItems(1) 'descripcion
Text2.Text = ListView1.selectedItem.SubItems(2) 'stock
Text2.SetFocus
End Sub

Private Sub Text2_Change()
If IsNumeric(Text2) = False Then
Text2.Text = ""
End If
End Sub

Private Sub Text4_Change()
ListView1.ListItems.Clear
'Busqueda y muestreo de los repuestops existentes
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from stock,repuestos where repuestos.cod_repu = stock.cod_stock and repuestos.descrip LIKE'" & (Text4.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView1.ListItems.Add(, , !cod_repu)
nuevo.SubItems(1) = !descrip
nuevo.SubItems(2) = !cant
busca.MoveNext
Next
End With
busca.Close
End Sub
