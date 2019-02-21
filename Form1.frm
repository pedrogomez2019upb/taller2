VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "¿Pase la materia?"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      Height          =   3615
      Left            =   6960
      TabIndex        =   14
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   735
      Left            =   4320
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   4320
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar promedio"
      Height          =   735
      Left            =   4320
      TabIndex        =   11
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox nota5 
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox nota4 
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox nota3 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox nota2 
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox nota1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Por favor ingresa las notas del 1 al 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Nota 5"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Nota 4"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Nota 3"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Nota 2"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Nota1"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Promedio Semestre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n1 = Val(nota1.Text) * 0.1
n2 = Val(nota2.Text) * 0.15
n3 = Val(nota3.Text) * 0.25
n4 = Val(nota4.Text) * 0.15
n5 = Val(nota5.Text) * 0.35

p = (n1 + n2 + n3 + n4 + n5)

If p >= 4.5 Then
respuesta.Text = "El promedio obtenido es de : " & p & ". Que buen promedio tienes!"
Else
respuesta.Text = "El promedio obtenido es de : " & p
End If


End Sub

Private Sub Command2_Click()
nota1.Text = ""
nota2.Text = ""
nota3.Text = ""
nota4.Text = ""
nota5.Text = ""
respuesta.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
If p >= 3 Then
respuesta.Text = "Con este promedio pasaste la materia/curso"
Else
respuesta.Text = "Con este promedio no pasaste la materia/curso"
End If
End Sub

