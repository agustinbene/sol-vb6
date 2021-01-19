VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Sistema de organizacion de laboratorios"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1482
      ButtonWidth     =   2328
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Altas"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "labo"
                  Text            =   "Laboratorio"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pc"
                  Text            =   "PC"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repu"
                  Text            =   "Repuesto"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "altadispo"
                  Text            =   "Dispositivos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "alta_deta"
                  Text            =   "Detalles"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bajas"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajalabo"
                  Text            =   "Laboratorios"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajapc"
                  Text            =   "Computadoras"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajadescrip"
                  Text            =   "Descripciones"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajadispo"
                  Text            =   "Dispositivos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajadetalle"
                  Text            =   "Detalles"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bajarepu"
                  Text            =   "Repuestos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificaciones"
            Key             =   "venta"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "modlabos"
                  Text            =   "Laboratorios"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "modPC"
                  Text            =   "PC"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "moddispo"
                  Text            =   "Dispositivos"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "modrepuesto"
                  Text            =   "Repuestos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "modstock"
                  Text            =   "Stock"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Busqueda"
            Key             =   "busque"
            ImageIndex      =   5
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "busquepc"
                  Text            =   "PC"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dispo"
                  Text            =   "PC por dispositivo"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "buscarepa"
                  Text            =   "Reparaciones"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reparaciones"
            Key             =   "repa"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "regirepa"
                  Text            =   "Informar Reparacion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "verrepa"
                  Text            =   "Ver Reparaciones"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.Image Image1 
         Height          =   15
         Left            =   0
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   4695
      Left            =   2760
      Top             =   1800
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Form_Load()
cargarconexion
Set ADOCN = New ADODB.Connection
With ADOCN
.CursorLocation = adUseClient
.ConnectionString = cadenacnx
.Open
End With


Image2.Stretch = True
Image2.Picture = LoadPicture(App.Path & "\fondo2.jpg")



  
End Sub
Private Sub Form_Resize()
With Image2
    .Left = 0
    .Top = 0
    .Width = Me.Width / 15
    .Height = Me.Height / 15
End With
End Sub


'Cerrar todas las ventanas al cerrar form 1
Private Sub Form_Unload(Cancel As Integer)

For Each i In Forms
If i.Name <> Me.Name Then Unload i
Next
Unload Me

End Sub







Private Sub Image3_Click()

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.Key
    Case Is = "labo"
        Form3.Show
    Case Is = "pc"
        Form4.Show
    Case Is = "repu"
        Form7.Show
    Case Is = "busquepc"
        Form2.Show
    Case Is = "dispo"
        Form6.Show
    Case Is = "regirepa"
        Form8.Show
    Case Is = "buscarepa"
        Form9.Show
    Case Is = "verrepa"
        Form9.Show
    Case Is = "modrepuesto"
        Form10.Show
    Case Is = "moddispo"
        Form11.Show
    Case Is = "modPC"
        Form12.Show
    Case Is = "modlabos"
        Form13.Show
    Case Is = "altadispo"
        Form14.Show
    Case Is = "bajapc"
        Form15.Show
    Case Is = "bajalabo"
        Form16.Show
    Case Is = "bajadispo"
        Form17.Show
    Case Is = "alta_deta"
        Form18.Show
    Case Is = "bajadetalle"
        Form19.Show
    Case Is = "bajarepu"
        Form20.Show
    Case Is = "modstock"
        Form10.Show
            Case Is = "bajadescrip"
        Form21.Show
    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "venta"
        Form5.Show
    End Select
End Sub



   
      
   
          
      
  

