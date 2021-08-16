VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Número de presupuesto"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "Numero presupuesto.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuestos"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce número para el nuevo presupuesto:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim consulta As String
    If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
        MsgBox "Error en la entrada de datos, no se admite valor vacio o NO numérico", vbCritical, "Error de datos"
    Else
        consulta = "select * from presupuestos where npresupuesto=""" + Text1.Text + """"
        Data1.RecordSource = consulta
        Data1.Refresh
        If Data1.Recordset.RecordCount = 1 Then
            MsgBox "Ya hay un presupuesto con ese mismo número", vbInformation, "Error en número de presupuesto"
        Else
            npres = Val(Form15.Text1.Text)
            Unload Me
            Load Form16
            Form16.Show vbModal
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

