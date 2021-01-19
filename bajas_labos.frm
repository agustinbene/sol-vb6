VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H000040C0&
   Caption         =   "Bajas Labos"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form16"
   ScaleHeight     =   4755
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Datos Laboratorio"
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   2040
         TabIndex        =   1
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Computadoras asociadas :"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form16"
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




Private Sub Combo1_Change()
Combo1.Text = ""
End Sub

Private Sub Combo1_click()

Dim busca As ADODB.Recordset
Set busca = New ADODB.Recordset
With busca
.Open "select * from pcs,labos where pcs.id_lab = labos.id and labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText
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
Label3.Caption = W

busca.MoveNext
Next


End With
busca.Close
End Sub



Private Sub Command2_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
If MsgBox("Esta seguro de eliminar este Laboratorio!? Hay " & Label3.Caption & " Computadoras asociadas a el!!!", vbYesNo, "MODIFICAR") = vbNo Then
Exit Sub
End If
Dim alta As ADODB.Recordset
Set altas = New ADODB.Recordset
With altas
.Open "delete * from labos where labos.nomb LIKE'" & Trim(Combo1.Text) & "%'", ADOCN, adOpenDynamic, adLockOptimistic, adCmdText

MsgBox ("Laboratorio: " & Combo1.Text & " Eliminado")
Combo1.Clear
Label3.Caption = ""

End With
End Sub
