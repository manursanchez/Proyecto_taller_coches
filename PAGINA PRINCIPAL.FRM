VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Facturacion Hmnos Maldonado Almagro"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10065
   Icon            =   "Pagina principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   120
      Picture         =   "Pagina principal.frx":08CA
      ScaleHeight     =   5955
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   855
         Left            =   7440
         Picture         =   "Pagina principal.frx":13A3A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del programa "
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Presupuestos"
         Height          =   855
         Left            =   5040
         Picture         =   "Pagina principal.frx":13E7C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Crear o ver presupuestos "
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Facturas"
         Height          =   855
         Left            =   2640
         Picture         =   "Pagina principal.frx":142BE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Crear o ver facturas "
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clientes"
         Height          =   855
         Left            =   120
         Picture         =   "Pagina principal.frx":14700
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Dar de alta un nuevo cliente o ver los clientes en la base de datos "
         Top             =   4800
         Width           =   2055
      End
   End
   Begin VB.Menu opciones 
      Caption         =   "Opciones de menu"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de este programa"
      End
      Begin VB.Menu repara 
         Caption         =   "Reparar base de datos"
      End
      Begin VB.Menu guion 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir del programa"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acerca_Click()
    Load Form18
    Form18.Show vbModal
End Sub

Private Sub Command1_Click()
    Load Form2
    Form2.Show vbModal
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Load Form8
    Form8.Show vbModal
End Sub

Private Sub Command4_Click()
    Load Form19
    Form19.Show vbModal
End Sub

Private Sub repara_Click()
    Load Form4
    Form4.Show vbModal
End Sub

Private Sub salir_Click()
    Unload Me
End Sub
