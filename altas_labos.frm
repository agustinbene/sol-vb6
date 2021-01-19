VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000040C0&
   Caption         =   "Agregar laboratorios"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5700
   LinkTopic       =   "Form3"
   ScaleHeight     =   3225
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Laboratorio"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar laboratorio"
         Height          =   615
         Left            =   1680
         TabIndex        =   3
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Software:"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_dropdown()
Combo1.Clear
Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo1.AddItem !nomb
busca.MoveNext
Next


End With
busca.Close
End Sub

Private Sub Combo2_dropdown()
Combo2.Clear
If Trim(Combo1.Text) <> "" Then

Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
busca.Open "select * FROM labos", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Combo2.AddItem !soft
busca.MoveNext
Next


End With
busca.Close
End If
End Sub

Private Sub Command1_Click()
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el laboratorio
.Open "Select * from labos where labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W < 1 Or .EOF Then
.AddNew
'se graba el nuevo labo
!nomb = UCase(Combo1.Text)
!soft = UCase(Combo2.Text)
.Update
MsgBox ("Laboratorio: " & Combo1.Text & " Agregado")
Else
MsgBox ("El Laboratorio: '" & Combo1.Text & "' Ya Existe!")
End If
altas.Close
End With

End Sub







Private Sub Command2_Click()

If Combo1.Text = "" Then
Exit Sub
End If
If MsgBox("Esta seguro de eliminar este Laboratorio!?", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
'Busqueda para ver si ya existe el laboratorio
.Open "select * FROM labos ,pcs where pcs.id_lab = labos.id and PCS.num_pc LIKE'%" & Trim("") & "%' and labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
W = altas.RecordCount
If W >= 1 Then
.Close
MsgBox ("No es posible eliminar el laboratorio ya que hay PCs relacionadas a el")
Exit Sub
End If
.Close
.Open "delete * from labos where labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("Laboratorio: " & Combo1.Text & " Eliminado")


End With
End Sub


