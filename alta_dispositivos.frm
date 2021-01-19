VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H000040C0&
   Caption         =   "Agregar dispositivos"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form14"
   ScaleHeight     =   4095
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Dispositivo"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Dispositivo"
         Height          =   615
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




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

Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el dispositivo
.Open "Select * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W = 0 Or .EOF Then
.AddNew
'se graba el nuevo dispositivo
!nombrehw = UCase(Combo1.Text)
.Update
MsgBox ("Dispositivo: " & Combo1.Text & " Agregado")
Else
MsgBox ("El Dispositivo: '" & Combo1.Text & "' Ya Existe!")
End If
altas.Close
End With
End Sub

Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el laboratorio
.Open "select * FROM dispositivos ,detalle_dispositivos where detalle_dispositivos.cod_dispo = dispositivos.cod and detalle_dispositivos.nombre LIKE'%" & Trim("") & "%' and dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W >= 1 Then
If MsgBox("Esta seguro de eliminar este dispositivo!? Hay Computadoras asociadas a el!!!", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If
End If
.Close
.Open "delete * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("Dispositivo: " & Combo1.Text & " Eliminado")


End With
End Sub
