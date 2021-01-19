VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H000040C0&
   Caption         =   "Baja dispositivos"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form17"
   ScaleHeight     =   3630
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Dispositivos"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   3015
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
Attribute VB_Name = "Form17"
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

Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
If MsgBox("Esta seguro de eliminar este dispositivo!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
.Open "delete * from dispositivos where dispositivos.nombrehw LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("Dispositivo: " & Combo1.Text & " Eliminado")
Combo1.Clear


End With
End Sub

