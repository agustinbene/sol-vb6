VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H000040C0&
   Caption         =   "Baja Repuestos"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form20"
   ScaleHeight     =   3675
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Repuesos"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   1560
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Change()
Combo1.Text = ""
End Sub

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

Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas


.Open "Select * from repuestos,reparaciones where reparaciones.cod_repuesto = repuestos.cod_repu and repuestos.descrip LIKE'" & Trim(Combo1.Text) & "%' and reparaciones.cant_usada LIKE'%" & ("") & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then

Else

MsgBox "El repuesto: " & Combo1.Text & " esta relacionado a una o varias reparaciones y no puede ser eliminado!", vbOKOnly + vbExclamation, "No se puede eliminar"















Exit Sub
End If
altas.Close


If MsgBox("Esta seguro de eliminar este repuesto!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If

.Open "delete * from repuestos where repuestos.descrip LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("Repuesto: " & Combo1.Text & " Eliminado")
Combo1.Clear


End With
End Sub


