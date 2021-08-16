VERSION 5.00
Begin VB.Form Form19 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elije opción"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "numeropresrecuperar.frx":0000
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escoje opción"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "Recuperar presupuesto"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nuevo presupuesto"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Form19.Option1 Then
        Unload Form19
        Load Form15
        Form15.Show vbModal
    End If
    If Form19.Option2 Then
        Unload Form19
        Load Form20
        Form20.Show vbModal
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
