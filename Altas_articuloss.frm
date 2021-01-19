VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form4 
   Caption         =   "Form3"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   7680
      TabIndex        =   29
      Top             =   3480
      Width           =   2175
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   4560
      TabIndex        =   22
      Top             =   5280
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4560
      TabIndex        =   13
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos PC nueva"
      Height          =   5655
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   6375
      Begin VB.ComboBox Combo13 
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox Combo12 
         Height          =   315
         Left            =   3960
         TabIndex        =   32
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   3960
         TabIndex        =   27
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   3960
         TabIndex        =   25
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   840
         TabIndex        =   20
         Top             =   3720
         Width           =   2175
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   840
         TabIndex        =   18
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Altas_articuloss.frx":0000
         Left            =   840
         List            =   "Altas_articuloss.frx":000D
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Articulo"
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "OK:"
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "monitor:"
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Soft:"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "S.O.:"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Disco:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Fuente:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Lectora:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "RAM:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "PCI:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "CPU:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Mother:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "PC:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "LAB:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   2520
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
            Picture         =   "Altas_articuloss.frx":001A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_articuloss.frx":08F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_articuloss.frx":11CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_articuloss.frx":1AA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   2
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
            ImageIndex      =   4
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
            ImageIndex      =   2
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
   Begin VB.Label Label15 
      Caption         =   "Fuente:"
      Height          =   255
      Left            =   6960
      TabIndex        =   28
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Fuente:"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Mother:"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
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



Private Sub Combo144_Click()
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM proveedores where apel LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
Codproveedor = !cod_prove
Label5.Caption = ("COD proveedor: " & Codproveedor)
End With
busca.Close
End Sub

Private Sub Combo1_DropDown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM lista", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo1.AddItem !lab
busca.MoveNext
Next


End With
busca.Close
End Sub






Private Sub Combo10_Change()
Combo10.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo10.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo11_Change()
Combo11.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo11.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo12_Change()
Combo12.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo12.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo2_Change()
Combo2.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo2.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo3_Change()
Combo3.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM cpu", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo3.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo4_Change()
Combo4.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM ram", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo4.AddItem !capacidad
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo5_Change()
Combo5.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM pci", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo5.AddItem !descripcion
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo6_Change()
Combo6.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo6.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo7_Change()
Combo7.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo7.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo8_Change()
Combo8.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo8.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Combo9_Change()
Combo9.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM mother", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo9.AddItem !modelo
busca.MoveNext
Next
End With
busca.Close
End Sub

Private Sub Command1_Click()
If Text1.Text <> "" Or Text2.Text <> "" Or Text3.Text <> "" Then
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
.Open "Select * from articulos ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!descrip = UCase(Text1)
!precio = Text2
!stock = Text3
!cod_prove = Codproveedor

.Update
End With
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
Else
MsgBox "Campos incompletos", vbExclamation, "ERROR404"
End If
End Sub

Private Sub Command2_Click()
Form4.Hide
Form3.Show
End Sub





Private Sub Text3_Change()
        If Text3.Text = "" Then
            Exit Sub
        End If
    If Not IsNumeric(Text3.Text) Then
        MsgBox "Ingrese solo numeros SIN espacios.", vbInformation
        Text3.Text = ""
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

