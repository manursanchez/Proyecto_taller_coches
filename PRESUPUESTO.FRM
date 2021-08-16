VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form17 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "Presupuesto.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Observaciones en presupuesto"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1560
      Picture         =   "Presupuesto.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6360
      Width           =   4095
   End
   Begin VB.Data totalpre 
      Caption         =   "TOTALPRESUPUESTOS"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7080
      TabIndex        =   39
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Quitar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   38
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   37
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   36
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 PRESUPUESTOS"
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
      RecordSource    =   "Presupuestos"
      Top             =   1080
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4-AUXLIAR DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3-TABLA AUXILIAR DE FACTURAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   840
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data1 
      Caption         =   "DATA1 CLIENTES"
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
      RecordSource    =   "Clientes"
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Observaciones"
      Height          =   1455
      Left            =   6960
      TabIndex        =   28
      Top             =   600
      Width           =   2055
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "observaciones"
         DataSource      =   "Data1"
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Totales"
      Height          =   1215
      Left            =   6960
      TabIndex        =   25
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox eurosiva 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox euros 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "€+Iva"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "€"
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presupuesto"
      Height          =   2535
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "Presupuesto.frx":0D0C
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Repuestos en presupuesto"
      Height          =   1455
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   8775
      Begin VB.CommandButton Command2 
         Caption         =   "Calcular IVA"
         Height          =   855
         Left            =   5640
         Picture         =   "Presupuesto.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   33
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   24
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Añadir"
         Enabled         =   0   'False
         Height          =   855
         Left            =   7200
         Picture         =   "Presupuesto.frx":1162
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Cantidad"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "€ + Iva"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         DataField       =   "peurosiva"
         DataSource      =   "Data2"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "€"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6615
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fax"
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "movil"
         DataField       =   "movil"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "tlf"
         DataField       =   "tlf"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matricula"
         DataField       =   "matricula"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vehiculo"
         DataField       =   "vehiculo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CP"
         DataField       =   "cp"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Provincia"
         DataField       =   "provincia"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Localidad"
         DataField       =   "localidad"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "direccion"
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apellidos"
         DataField       =   "apellidos"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DNI"
         DataField       =   "dni"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de presupuesto:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Dim guardar As Boolean
Dim imprimir As Boolean

Private Sub Command1_Click()
    Dim a As Currency
    Dim b As Currency
    Dim cant As Integer 'Va a contener la cantidad de elementos
    Dim numero As Currency
    If Form17.Text4.Text = "" Or Not IsNumeric(Form17.Text4.Text) Then
        MsgBox "El campo cantidad esta vacío o tiene valores NO NUMERICOS .Por favor compruebelo antes de continuar", vbExclamation, "Error de datos"
    Else
        cant = 0 'Inicializamos cant
        'Añadimos el cliente a la tabla auxiliar de clientes
        If bandera = False Then 'Controlamos que solo lo introduzcamos una sola vez
            Form17.Data4.Recordset.AddNew
            Form17.Data4.Recordset.Fields("numfac") = Form17.Label29.Caption
            Form17.Data4.Recordset.Fields("fecha") = Form17.Text2.Text
            Form17.Data4.Recordset.Fields("dni") = Form17.Label3.Caption
            Form17.Data4.Recordset.Fields("nombre") = Form17.Label4.Caption
            Form17.Data4.Recordset.Fields("apellidos") = Form17.Label5.Caption
            Form17.Data4.Recordset.Fields("direccion") = Form17.Label6.Caption
            Form17.Data4.Recordset.Fields("cp") = Val(Form17.Label7.Caption)
            Form17.Data4.Recordset.Fields("localidad") = Form17.Label8.Caption
            Form17.Data4.Recordset.Fields("provincia") = Form17.Label9.Caption
            Form17.Data4.Recordset.Fields("vehiculo") = Form17.Label10.Caption
            Form17.Data4.Recordset.Fields("matricula") = Form17.Label11.Caption
            Form17.Data4.Recordset.Fields("tlf") = Form17.Label12.Caption
            Form17.Data4.Recordset.Fields("fax") = Form17.Label13.Caption
            Form17.Data4.Recordset.Fields("movil") = Form17.Label14.Caption
            Form17.Data4.Recordset.Fields("observaciones") = Form17.Label24.Caption
            Form17.Data4.Recordset.Update
            Form17.Data4.Refresh
            bandera = True
        End If
        
        'Vamos a añadir elementos a la tabla auxiliar de facturas
    
        Form17.Data3.Recordset.AddNew
        Form17.Data3.Recordset.Fields("facnum") = Form17.Label29.Caption
        Form17.Data3.Recordset.Fields("concepto") = Form17.Text1.Text
        Form17.Data3.Recordset.Fields("cantidad") = Form17.Text4.Text
   
    
        'Vamos a hacer las operaciones matematicas para los totales
        'Las mostraremos primero por pantalla en sus correspondientes
        'textboxs y despues los meteremos en la tabla
        
        cant = Val(Form17.Text4.Text) 'Le asignamos a cant la cantidad de elementos
    
        'Total en Euros
        a = Form17.Text5.Text
        numero = cant * a
        teuros = teuros + numero
        euros.Text = Format(teuros, "#########.##")
        
        'Total en euros + IVA
        b = Form17.Label21.Caption
        numero = cant * b
        teurosiva = teurosiva + numero
        eurosiva.Text = Format(teurosiva, "#########.##")
    
        'Introducimos las cantidades y los precios en la tabla auxiliar
        If Text5.Text = 0 Then
            Form17.Data3.Recordset.Fields("peuros") = 0
            Form17.Data3.Recordset.Fields("peurosiva") = 0
        Else
            Form17.Data3.Recordset.Fields("peuros") = Format(cant * Form17.Text5.Text, "#####.##")
            Form17.Data3.Recordset.Fields("peurosiva") = Format(cant * Form17.Label21.Caption, "#####.##")
        End If
        Form17.Data3.Recordset.Update
        Form17.Data3.Refresh
        Form17.Text4.Text = 1
        Text1.Text = ""
        Text5.Text = ""
        Text4.Text = 1
        Label21.Caption = ""
        Command1.Enabled = False
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command7.Enabled = True
    End If
End Sub


Private Sub Command2_Click()
    Dim comodin As Double
    If Text5.Text = "" Or Not IsNumeric(Text5.Text) Or Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
        MsgBox "Imposible hacer calculos ya que hay algun campo vacio o erroneo. Compruebe los campos precio en euros y cantidad antes de continuar", vbExclamation, "Error de datos"
    Else
        If Text5.Text = 0 Then
            Label21.Caption = 0
        Else
            comodin = Text5.Text '* Text4.Text
            Label21.Caption = (comodin * 16 / 100) + comodin
            Label21.Caption = Format(Label21.Caption, "#######.##")
        End If
        Command1.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    
    guardarfactura
    guardartotales
    imprimir = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command7.Enabled = False
End Sub

Private Sub Command4_Click()
    
    If Not Form17.Data3.Recordset.BOF Then
        Form17.Command1.Enabled = False 'Desactivamos el boton añadir a factura
        Form17.Command5.Enabled = False 'Desactivamos el boton elimianr componente de factura
        Form17.Command3.Enabled = False
        Form17.Command7.Enabled = False 'Desactivamos el boton de observaciones en factura
        guardarfactura
        guardartotales
        guardar = True
        imprimir = True
        DataEnvironment2.Connection1.Open
        DataEnvironment2.Command1
        DataReport4.Show vbModal
        DataEnvironment2.Connection1.Close
    Else
        MsgBox "No puedo imprimir nada por que no se han introducido datos en la factura", vbInformation
    End If
End Sub

Private Sub Command5_Click()
    Load Form7
    Form7.Show vbModal
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Command7_Click()
    Load Form6
    Form6.Show vbModal
End Sub

Private Sub Form_Load()
    Dim consulta As String
    bandera = False
    imprimir = False
    guardar = False
    Form17.Label29.Caption = npres
    Text2.Text = Date
    consulta = "select * from clientes where dni=""" + identificacion + """"
    Data1.RecordSource = consulta
    Data1.Refresh
End Sub
Private Sub guardarfactura() 'Guarda de faxturas auxiliares a presupuestos
     If imprimir = False Then
        
        If Form17.Data4.Recordset.BOF = True Then
            MsgBox "No hay presupuesto para guardar ni para imprimir.Añada elementos a factura y guarde o imprima despues", vbInformation, "Error de datos"
        Else
            Form17.Data3.Recordset.MoveFirst
            Do While Not Form17.Data3.Recordset.EOF
                Form17.Data5.Recordset.AddNew
                Form17.Data5.Recordset.Fields("npresupuesto") = Val(Form17.Label29.Caption)
                Form17.Data5.Recordset.Fields("fecha") = Form17.Text2.Text
                Form17.Data5.Recordset.Fields("dni") = Form17.Label3.Caption
                Form17.Data5.Recordset.Fields("nombre") = Form17.Label4.Caption
                Form17.Data5.Recordset.Fields("apellidos") = Form17.Label5.Caption
                Form17.Data5.Recordset.Fields("direccion") = Form17.Label6.Caption
                Form17.Data5.Recordset.Fields("vehiculo") = Form17.Label7.Caption
                Form17.Data5.Recordset.Fields("matricula") = Form17.Label8.Caption
                Form17.Data5.Recordset.Fields("cp") = Form17.Label9.Caption
                Form17.Data5.Recordset.Fields("localidad") = Form17.Label10.Caption
                Form17.Data5.Recordset.Fields("provincia") = Form17.Label11.Caption
                Form17.Data5.Recordset.Fields("tlf") = Form17.Label12.Caption
                Form17.Data5.Recordset.Fields("movil") = Form17.Label13.Caption
                Form17.Data5.Recordset.Fields("fax") = Form17.Label14.Caption
                Form17.Data5.Recordset.Fields("observaciones") = Form17.Label24.Caption
                Form17.Data5.Recordset.Fields("concepto") = Form17.Data3.Recordset.Fields("concepto")
                Form17.Data5.Recordset.Fields("cantidad") = Form17.Data3.Recordset.Fields("cantidad")
                Form17.Data5.Recordset.Fields("peuros") = Form17.Data3.Recordset.Fields("peuros")
                Form17.Data5.Recordset.Fields("peurosiva") = Form17.Data3.Recordset.Fields("peurosiva")
                Form17.Data5.Recordset.Fields("obsenpres") = Form17.Data4.Recordset.Fields("OBSENFAC")
                Form17.Data5.UpdateRecord
                Form17.Data3.Recordset.MoveNext
            Loop
            MsgBox "Presupuesto guardado sin problemas", vbInformation
            Form17.Data5.Refresh
        End If
    End If
End Sub
Private Sub borrarauxclientes()
    If Not Form17.Data4.Recordset.BOF Then
        Form17.Data4.Recordset.MoveFirst
        Form17.Data4.Recordset.Delete
        Form17.Data4.UpdateRecord
        Form17.Data4.Refresh
    End If
End Sub
Private Sub borrarauxfacturas()
    If Not Form17.Data3.Recordset.BOF Then
        Form17.Data3.Recordset.MoveFirst
        Do While Not Form17.Data3.Recordset.EOF
            Form17.Data3.Recordset.Delete
            Form17.Data3.Recordset.MoveNext
        Loop
        Form17.Data3.Refresh
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    borrarauxclientes
    borrarauxfacturas
End Sub
Private Sub guardartotales()
    'totalpre es el data que esta conectado a la tabla totales presupuestos
    If imprimir = False Then
        Form17.totalpre.Recordset.AddNew
        Form17.totalpre.Recordset.Fields("npresupuesto") = Form17.Label29.Caption
        If euros.Text = "," Or eurosiva = "," Then
            Form17.totalpre.Recordset.Fields("teuros") = 0
            Form17.totalpre.Recordset.Fields("teurosiva") = 0
        Else
            Form17.totalpre.Recordset.Fields("teuros") = Form17.euros.Text
            Form17.totalpre.Recordset.Fields("teurosiva") = Form17.eurosiva.Text
        End If
        Form17.totalpre.Recordset.Update
        Form17.totalpre.Refresh
    End If
End Sub
