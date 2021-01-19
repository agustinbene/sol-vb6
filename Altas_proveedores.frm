VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos Proveedor"
      Height          =   5655
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Proveedor"
         Height          =   615
         Left            =   2040
         TabIndex        =   5
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   975
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
            Picture         =   "Altas_proveedores.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_proveedores.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_proveedores.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_proveedores.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   10
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
         NumButtons      =   3
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
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Venta"
            Key             =   "venta"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" Or Text2.Text <> "" Or Text3.Text <> "" Or Text4.Text <> "" Then
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
.Open "Select * from proveedores ", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
.AddNew
!apel = UCase(Text1)
!nombre = Text2
!tel = Text3
!direc = Text4
.Update
End With
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
Else
MsgBox "Campos incompletos", vbExclamation, "ERROR404"
End If
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
    Case Is = "artic"
    Form4.Show
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

