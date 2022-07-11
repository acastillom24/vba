VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login_VMC 
   Caption         =   "MAPFRE PERU"
   ClientHeight    =   7248
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15336
   OleObjectBlob   =   "Login_VMC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Login_VMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If Me.TextBox1.Value = "alincastillo" And Me.TextBox2.Value = "Ac1995$24" Then
        Application.Visible = True
        Unload Me
        ElseIf Me.TextBox1.Value = "mapfreperu" And Me.TextBox2.Value = "crgper7i" Then
            Login_VMC.Hide
            Load ctdGruposForm
            ctdGruposForm.Show
        Else:
            
    End If
End Sub

Private Sub CommandButton1_Enter()
    If Me.TextBox1.Value = "alincastillo" And Me.TextBox2.Value = "Ac1995$24" Then
        Application.Visible = True
        Unload Me
        ElseIf Me.TextBox1.Value = "mapfreperu" And Me.TextBox2.Value = "crgper7i" Then
            Login_VMC.Hide
            Load ctdGruposForm
            ctdGruposForm.Show
        Else:
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub TextBox1_Enter()
    Me.Label1.Visible = False
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.TextBox1.Value = "" Then
        Me.Label1.Visible = True
    End If
End Sub

Private Sub TextBox2_Enter()
    
    Me.Label2.Visible = False

End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.TextBox1.Value = "" Then
        Me.Label2.Visible = True
    End If
End Sub
