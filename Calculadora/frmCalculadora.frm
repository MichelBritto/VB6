VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmCalculadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora"
   ClientHeight    =   5325
   ClientLeft      =   8235
   ClientTop       =   2205
   ClientWidth     =   4395
   Icon            =   "frmCalculadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin MSScriptControlCtl.ScriptControl msc 
      Left            =   510
      Top             =   5790
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Frame fraPrincipal 
      Height          =   5385
      Left            =   -30
      TabIndex        =   0
      Top             =   -60
      Width           =   4425
      Begin VB.CommandButton cmdAbreParentese 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2340
         TabIndex        =   15
         Top             =   4500
         Width           =   885
      End
      Begin VB.CommandButton cmdFechaParentese 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         TabIndex        =   16
         Top             =   4500
         Width           =   885
      End
      Begin VB.CommandButton cmdDividir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         Picture         =   "frmCalculadora.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3690
         Width           =   885
      End
      Begin VB.CommandButton cmdMultiplicar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         Picture         =   "frmCalculadora.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2910
         Width           =   885
      End
      Begin VB.CommandButton cmdSubtrair 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         Picture         =   "frmCalculadora.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2100
         Width           =   885
      End
      Begin VB.CommandButton cmdEqual 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         Picture         =   "frmCalculadora.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4500
         Width           =   1935
      End
      Begin VB.CommandButton cmdSoma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         Picture         =   "frmCalculadora.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1290
         Width           =   885
      End
      Begin VB.CommandButton cmdBack 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2310
         Picture         =   "frmCalculadora.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3690
         Width           =   885
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   19
         Top             =   3690
         Width           =   885
      End
      Begin VB.CommandButton cmd0 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1260
         TabIndex        =   10
         Top             =   3690
         Width           =   885
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1260
         TabIndex        =   8
         Top             =   2910
         Width           =   885
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2340
         TabIndex        =   9
         Top             =   2910
         Width           =   885
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   7
         Top             =   2910
         Width           =   885
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1260
         TabIndex        =   5
         Top             =   2100
         Width           =   885
      End
      Begin VB.CommandButton cmd6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2340
         TabIndex        =   6
         Top             =   2100
         Width           =   885
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   4
         Top             =   2100
         Width           =   885
      End
      Begin VB.CommandButton cmd8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1260
         TabIndex        =   2
         Top             =   1290
         Width           =   885
      End
      Begin VB.CommandButton cmd9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2340
         TabIndex        =   3
         Top             =   1290
         Width           =   885
      End
      Begin VB.CommandButton cmd7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   1
         Top             =   1290
         Width           =   885
      End
      Begin VB.TextBox txtVisor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   4185
      End
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdblResultado As Double

Private Sub cmd1_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 1
End Sub
Private Sub cmd2_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 2
End Sub
Private Sub cmd3_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 3
End Sub
Private Sub cmd4_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 4
End Sub
Private Sub cmd5_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 5
End Sub
Private Sub cmd6_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 6
End Sub
Private Sub cmd7_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 7
End Sub
Private Sub cmd8_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 8
End Sub
Private Sub cmd9_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & 9
End Sub
Private Sub cmd0_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    txtVisor.Text = txtVisor.Text & "0"
End Sub

Private Sub cmdAbreParentese_Click()
    txtVisor.BackColor = vbWhite
    If Right(txtVisor.Text, 1) = "(" Then Exit Sub
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Not TemOperacaoRepetida("(") Then
        txtVisor.Text = txtVisor.Text & "("
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & "("
    End If
End Sub

Private Sub cmdBack_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "" Then txtVisor.Text = "0"
    txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
    If txtVisor.Text = "" Then txtVisor.Text = "0"
End Sub

Private Sub cmdClear_Click()
    txtVisor.BackColor = vbWhite
    txtVisor.Text = "0"
End Sub

Private Sub calcular()
'msc = Microsoft Script Control
On Error GoTo erro
    If txtVisor.Text <> "" Then
        txtVisor.Text = msc.Eval(txtVisor.Text)
        txtVisor.BackColor = vbHighlight
    Else
        txtVisor.Text = "0"
        txtVisor.BackColor = vbWhite
    End If
    Exit Sub
erro:
    Call MsgBox("Expressão contém um ou mais erros de lógica!" & vbCrLf & vbCrLf & "Corrija e tente novamente!", vbOKOnly + vbExclamation, "Atenção!")
    txtVisor.Text = "0"
    txtVisor.BackColor = vbWhite
End Sub

Private Sub cmdDividir_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Right(txtVisor.Text, 1) = "/" Then Exit Sub
    If Not TemOperacaoRepetida("/") Then
        txtVisor.Text = txtVisor.Text & "/"
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & "/"
    End If
End Sub

Private Sub cmdEqual_Click()
    calcular
End Sub

Private Sub cmdFechaParentese_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Right(txtVisor.Text, 1) = ")" Then Exit Sub
    If Not TemOperacaoRepetida(")") Then
        txtVisor.Text = txtVisor.Text & ")"
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & ")"
    End If
End Sub

Private Sub cmdMultiplicar_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Right(txtVisor.Text, 1) = "*" Then Exit Sub
    If Not TemOperacaoRepetida("*") Then
        txtVisor.Text = txtVisor.Text & "*"
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & "*"
    End If
End Sub

Private Sub cmdSoma_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Right(txtVisor.Text, 1) = "+" Then Exit Sub
    If TemOperacaoRepetida("+") = False Then
        txtVisor.Text = txtVisor.Text & "+"
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & "+"
    End If
End Sub

Private Sub cmdSubtrair_Click()
    txtVisor.BackColor = vbWhite
    If txtVisor.Text = "0" Then txtVisor.Text = ""
    If Right(txtVisor.Text, 1) = "-" Then Exit Sub
    If Not TemOperacaoRepetida("-") Then
        txtVisor.Text = txtVisor.Text & "-"
    Else
        txtVisor.Text = Left(txtVisor.Text, Len(txtVisor.Text) - 1)
        txtVisor.Text = txtVisor.Text & "-"
    End If
End Sub

Private Function TemOperacaoRepetida(strOp As String) As Boolean
    TemOperacaoRepetida = False

    If strOp <> "(" And strOp <> ")" Then
    
        Select Case Right(txtVisor.Text, 1)
            Case "+"
                    TemOperacaoRepetida = True
            Case "-"
                TemOperacaoRepetida = True
            Case "*"
                TemOperacaoRepetida = True
            Case "/"
                TemOperacaoRepetida = True
    '        Case "("
    '            TemOperacaoRepetida = True
    '        Case ")"
    '            TemOperacaoRepetida = True
        End Select
    End If

End Function

Private Sub Form_Load()
    txtVisor.Enabled = True
End Sub
