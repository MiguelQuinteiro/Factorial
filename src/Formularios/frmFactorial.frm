VERSION 5.00
Begin VB.Form frmFactorial 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PROYECTO FACTORIAL   -- VERSIÓN  1.0"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFactorial 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "FACTORIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Text            =   "7"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtFactorial 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblFactorial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "El factorial es:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmFactorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : FACTORIAL
'* CONTENIDO     : CALCULA EL FACTORIAL DE UN NUMERO
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 04 DE MARZO DE 2014
'* ACTUALIZACION : 04 DE MARZO DE 2014
'****************************************************************************************
Option Explicit

Private Sub cmdFactorial_Click()
  txtFactorial = Factorial(Val(txtNumero))
End Sub

Function Factorial(ByVal n As Integer) As Double
  If n <= 1 Then
    Factorial = 1
  Else
    Factorial = (n * Factorial(n - 1))
  End If
End Function


