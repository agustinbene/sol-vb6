VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   BackColor       =   &H000040C0&
   Caption         =   "Modificacion de laboratorios"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form13"
   ScaleHeight     =   5775
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Modificaciones"
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "Software"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "ID Laboraorio"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Busqueda"
      Height          =   4935
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "COD"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Software"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         Caption         =   "Dispositivo:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   4440
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
Exit Sub
End If

If MsgBox("Esta seguro de modificar este Articulo!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If

Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
altas.Open "select * from labos where labos.id LIKE'" & (Label5.Caption) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
!nomb = UCase(Text1.Text)
!soft = UCase(Text2.Text)
.Update
End With

Text1.Text = ""

Label5.Caption = ""



'Busqueda y muestreo de los repuestops existentes
ListView1.ListItems.Clear

Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView1.ListItems.Add(, , !ID)
nuevo.SubItems(1) = !nomb
nuevo.SubItems(2) = !soft
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
.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Set nuevo = ListView1.ListItems.Add(, , !ID)
nuevo.SubItems(1) = !nomb
nuevo.SubItems(2) = !soft
busca.MoveNext
Next
End With
busca.Close
End Sub



Private Sub ListView1_DblClick()
Frame1.Enabled = True
Label5.Caption = ListView1.selectedItem.Text 'cod
Text1.Text = ListView1.selectedItem.SubItems(1) 'descripcion
Text2.Text = ListView1.selectedItem.SubItems(2) 'descripcion

End Sub
