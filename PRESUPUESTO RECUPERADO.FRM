VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form21 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto recuperado"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "Presupuesto recuperado.frx":0000
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos del cliente"
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   9735
      Begin VB.Data Data3 
         Caption         =   "Data3 AUXILIAR DE FACTURAS"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Aux Facturas"
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DNI"
         DataField       =   "dni"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE"
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "APELLIDOS"
         DataField       =   "apellidos"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DIRECCION"
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CP"
         DataField       =   "cp"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LOCALIDAD"
         DataField       =   "localidad"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROVINCIA"
         DataField       =   "provincia"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VEHICULO"
         DataField       =   "vehiculo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MATRICULA"
         DataField       =   "matricula"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TLF"
         DataField       =   "tlf"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOVIL"
         DataField       =   "movil"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FAX"
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   975
      Left            =   120
      Picture         =   "Presupuesto recuperado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 AUXILIAR DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   5175
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label18 
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "€UROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "€+IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Data Data2 
      Caption         =   "DATA2 TABLA TOTALESPRES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "totalespre"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar presupuesto"
      Height          =   975
      Left            =   3720
      Picture         =   "Presupuesto recuperado.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir Presupuesto"
      Height          =   975
      Left            =   7200
      Picture         =   "Presupuesto recuperado.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 PRESUPUESTOS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuestos"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "OBSENFAC"
      DataSource      =   "Data1"
      Height          =   975
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4800
      Width           =   4215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Presupuesto recuperado.frx":1590
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "Presupuesto recuperado.frx":15A4
      TabIndex        =   0
      Top             =   2520
      Width           =   9735
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C000&
      Caption         =   " NUMERO DE FACTURA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NFAC"
      DataField       =   "numfac"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C000&
      Caption         =   " FECHA DE FACTURA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   26
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA"
      DataField       =   "fecha"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   25
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Observaciones de factura:"
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   4440
      Width           =   2535
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim opc As Integer
    Dim consulta As String
    Dim consultat As String
    opc = MsgBox("¿Esta seguro de querer eliminar el presupuesto actual?", vbOKCancel, "Eliminar factura")
    If opc = vbOK Then
        'Eliminacion de datos de tabla presupuestos
        consulta = "select * from presupuestos where npresupuesto=""" + Label14.Caption + """"
        Data6.RecordSource = consulta
        Data6.Refresh
        If Data6.Recordset.RecordCount = 1 Then
            Data6.Recordset.MoveFirst
            Do While Not Data6.Recordset.EOF
                Data6.Recordset.Delete
                Data6.Recordset.MoveNext
            Loop
        Else
            MsgBox "¡NO ENCUENTRO LA FACTURA PARA PODER BORRARLA!", vbCritical, "Error en la busqueda del presupuesto"
            MsgBox "Consulte con el programador acerca de este error", vbInformation
        End If
        'Borramos de la tabla principal de totalespre
        consultat = "select * from totalespre where npresupuesto=""" + Label14.Caption + """"
        Data2.RecordSource = consultat
        Data2.Refresh
        If Data2.Recordset.RecordCount = 1 Then
            Data2.Recordset.Delete
            Data2.Refresh
        Else
            MsgBox "NO ENCUENTRO EL REGISTRO A BORRAR EN TABLA TOTALES,CONSULTA AL PROGRAMADOR ACERCA DE ESTE ERROR", vbCritical, "Error en la eliminación"
        End If
        MsgBox "El presupuesto ha sido eliminado satisfactoriamenmte", vbInformation, "Información de eliminación"
        Unload Form21
    Else
        MsgBox "Operación de eliminacion cancelada por el usuario", vbInformation, "Factura no eliminada"
    End If
End Sub



Private Sub Command4_Click()
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command1
    DataReport5.Show vbModal
    DataEnvironment2.Connection1.Close
End Sub

Private Sub Form_Load()
    'Pasamos los totales a los labels que tenemos habilitados en el formulario
    Form21.Text2.Text = teuros
    Form21.Text4.Text = teurosiva
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Eliminacion de datos de las tablas auxiliares
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
        Loop
        Data1.Recordset.MoveFirst
        Data1.Recordset.Delete
        
End Sub

