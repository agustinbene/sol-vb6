VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3360
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
            Picture         =   "Altas_clientes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_clientes.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_clientes.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Altas_clientes.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Clientes"
      Height          =   5655
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Cliente"
         Height          =   615
         Left            =   2040
         TabIndex        =   10
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

