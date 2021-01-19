VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H000040C0&
   Caption         =   "Baja detalles"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form19"
   ScaleHeight     =   4245
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Dispositivo"
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Detalle"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Dispositivo:"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
Combo1.Text = ""
End Sub
Private Sub Combo2_Change()
Combo2.Text = ""
End Sub

Private Sub Combo1_dropdown()
Combo1.Clear
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
Combo1.AddItem !nombrehw
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
.Open "select * from detalle_dispositivos,dispositivos where dispositivos.cod = detalle_dispositivos.cod_dispo and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo2.AddItem !nombre
busca.MoveNext
Next
End With
busca.Close

Dim r As Integer, t As Integer
For r = 0 To Combo2.ListCount - 1
For t = 0 To Combo2.ListCount - 1
If (Combo2.List(r) = Combo2.List(t)) And t <> r Then
Combo2.RemoveItem (r)
End If
Next t
Next r
 End Sub


Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el dispositivo
.Open "Select * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W > 0 Then
codd = !cod
Else
.AddNew
'se graba el nuevo dispositivo
!nombrehw = UCase(Combo1.Text)
.Update
codd = !cod
End If
altas.Close
'Busqueda para ver si ya existe el detalle
.Open "Select * from dispositivos,detalle_dispositivos where detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%' and detalle_dispositivos.nombre LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!nombre = UCase(Combo2.Text)
!cod_dispo = codd
.Update
MsgBox ("Detalle: " & Combo2.Text & " Agregado")
Else
MsgBox ("El Detalle: '" & Combo2.Text & "' Ya Existe!")
End If
altas.Close
End With
End Sub

Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el dispositivo
.Open "Select * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W > 0 Then
codd = !cod
altas.Close
Else
.AddNew
'se graba el nuevo dispositivo
!nombrehw = UCase(Combo1.Text)
.Update
codd = !cod
altas.Close
End If

'Busqueda para ver si ya existe el detalle
.Open "delete detalle_dispositivos.* from dispositivos,detalle_dispositivos where detalle_dispositivos.cod_dispo = dispositivos.cod and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%' and detalle_dispositivos.nombre LIKE'" & Trim(Combo2.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
If W >= 1 Then
MsgBox ("Detalle: " & Combo2.Text & " Eliminado")
Combo2.Clear
End If

End With
End Sub
