VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizacion y tratamiento de la Base de datos de clientes"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Listado de clientes por pantalla"
      Height          =   975
      Left            =   5040
      Picture         =   "Form3.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operaciones "
      Height          =   2535
      Left            =   7200
      TabIndex        =   31
      Top             =   3120
      Width           =   3255
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar cliente actual"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modificar clientes"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir listados"
      Height          =   2415
      Left            =   7200
      TabIndex        =   26
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Listado de clientes con telefonos"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Listado de clientes con automoviles"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Listado de clientes con direcciones"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox Text13 
      DataField       =   "observaciones"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   3240
      MaxLength       =   255
      ScrollBars      =   3  'Both
      TabIndex        =   25
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox Text12 
      DataField       =   "fax"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   24
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      DataField       =   "movil"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataField       =   "tlf"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   22
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "matricula"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "vehiculo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   20
      Top             =   3600
      Width           =   5055
   End
   Begin VB.TextBox Text7 
      DataField       =   "provincia"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "localidad"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   18
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "cp"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccion"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   16
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      DataField       =   "apellidos"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   15
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   14
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      DataField       =   "dni"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      MaxLength       =   9
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Pulsa las flechas para moverte a traves de los registros"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   6120
      Width           =   8055
   End
   Begin VB.Label Label13 
      Caption         =   "NOMBRE"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "APELLIDOS"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "DIRECCION"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "CP"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "LOCALIDAD"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "PROVINCIA"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "VEHICULO"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "MATRICULA"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "TELEFONO"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "MOVIL"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "FAX"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "OBSERVACIONES"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "DNI"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Private Sub Command1_Click()
    If Form3.Option1 Then
        MsgBox "Asegúrese que tiene papel en la impresora y que está encendida", vbInformation, "Información de sistema"
        DataReport1.PrintReport
    End If
    If Form3.Option2 Then
        MsgBox "Asegúrese que tiene papel en la impresora y que está encendida", vbInformation, "Información de sistema"
        DataReport2.PrintReport
    End If
    If Form3.Option3 Then
        MsgBox "Asegúrese que tiene papel en la impresora y que está encendida", vbInformation, "Información de sistema"
        DataReport3.PrintReport
    End If
End Sub

Private Sub Command2_Click()
    'Form3.Text1.Enabled = True
    If bandera = False Then
        Form3.Text2.Enabled = True
        Form3.Text3.Enabled = True
        Form3.Text4.Enabled = True
        Form3.Text5.Enabled = True
        Form3.Text6.Enabled = True
        Form3.Text7.Enabled = True
        Form3.Text8.Enabled = True
        Form3.Text9.Enabled = True
        Form3.Text10.Enabled = True
        Form3.Text11.Enabled = True
        Form3.Text12.Enabled = True
        Form3.Text13.Enabled = True
        Form3.Command2.Caption = "Modificado"
        bandera = True
        Form3.Data1.Enabled = False
        Form3.Command4.Enabled = False
    Else
        If bandera = True Then
            Form3.Text2.Enabled = False
            Form3.Text3.Enabled = False
            Form3.Text4.Enabled = False
            Form3.Text5.Enabled = False
            Form3.Text6.Enabled = False
            Form3.Text7.Enabled = False
            Form3.Text8.Enabled = False
            Form3.Text9.Enabled = False
            Form3.Text10.Enabled = False
            Form3.Text11.Enabled = False
            Form3.Text12.Enabled = False
            Form3.Text13.Enabled = False
            Form3.Command2.Caption = "Modificar cliente"
            bandera = False
            Form3.Data1.Enabled = True
            Form3.Command4.Enabled = True
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim opc As Integer
    opc = MsgBox("¿Esta seguro de querer eliminar el cliente actual?", vbCritical + vbOKCancel)
    If opc = vbOK Then
        If Form3.Data1.Recordset.BOF = True Then
            MsgBox "¡¡¡No hay clientes para borrar!!!", vbExclamation, "Error de datos"
        Else
            Form3.Data1.Recordset.Delete
            Form3.Data1.Refresh
            MsgBox "Cliente eliminado", vbInformation
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Load Form22
    Form22.Show vbModal
End Sub

Private Sub Form_Load()
    bandera = False
End Sub

'A partir de aqui son las validaciones de los campos cuando se estan modificando

Private Sub Text2_Validate(Cancel As Boolean)
    If Form3.Text2.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If Form3.Text3.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text4_validate(Cancel As Boolean)
    If Form3.Text4.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text5_validate(Cancel As Boolean)
    If Form3.Text5.Text = "" Or Not IsNumeric(Form3.Text5.Text) Then
        MsgBox "Tiene que introducir algun valor o el valor no es numerico", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text6_validate(Cancel As Boolean)
    If Form3.Text6.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text7_validate(Cancel As Boolean)
    If Form3.Text7.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text8_validate(Cancel As Boolean)
    If Form3.Text8.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text9_validate(Cancel As Boolean)
    If Form3.Text9.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text10_validate(Cancel As Boolean)
    If Form3.Text10.Text = "" Or Not IsNumeric(Form3.Text10.Text) Then
        MsgBox "Tiene que introducir algun valor o el valor no es numérico", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text11_validate(Cancel As Boolean)
    If Form3.Text11.Text = "" Or Not IsNumeric(Form3.Text11.Text) Then
        MsgBox "Tiene que introducir algun valor o el valor no es numérico", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text12_validate(Cancel As Boolean)
    If Form3.Text12.Text = "" Or Not IsNumeric(Form3.Text12.Text) Then
        MsgBox "Tiene que introducir algun valor o el valor no es numérico", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text13_validate(Cancel As Boolean)
    If Form3.Text13.Text = "" Then
        MsgBox "Tiene que introducir algun valor en este cuadro", vbExclamation
        Cancel = True
    End If
End Sub
