VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de factura"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Observaciones en factura"
      Height          =   1095
      Left            =   6360
      Picture         =   "Form12.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   " Pulsa aqui para introducir observaciones en la factura "
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "TOTALES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Funciones"
      Height          =   4575
      Left            =   9240
      TabIndex        =   43
      Top             =   3120
      Width           =   2055
      Begin VB.CommandButton nuevafactura 
         Caption         =   "Nueva factura"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   " Crear una nueva factura "
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton salir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Salir al panel principal del programa"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton eliminar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Eliminar algun componente que hayamos introducido por error en la factura"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton imprimirinforme 
         Caption         =   "Imprimir factura"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Sacar la factura por impresora"
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Totales"
      Height          =   1575
      Left            =   6360
      TabIndex        =   38
      Top             =   3120
      Width           =   2655
      Begin VB.TextBox eurosiva 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox euros 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label44 
         Caption         =   "€uros + IVA"
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "€uros"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Data Data4 
      Caption         =   "DATA 4 AUX DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "DATA3 AUXILIAR DE FACTURAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7800
      TabIndex        =   36
      ToolTipText     =   "Formato de fecha dia/mes/año. Ejemplo 25/08/2002"
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form12.frx":1E14
      Height          =   2895
      Left            =   240
      TabIndex        =   35
      Top             =   4920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame2 
      Caption         =   "Materiales para introducir en factura"
      Height          =   1575
      Left            =   240
      TabIndex        =   29
      Top             =   3120
      Width           =   5895
      Begin VB.CommandButton añadir 
         Caption         =   "Añadir a factura"
         Height          =   375
         Left            =   2760
         TabIndex        =   51
         ToolTipText     =   " Pulsar para añadir elementos a la factura "
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular IVA"
         Height          =   255
         Left            =   4440
         TabIndex        =   50
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   48
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "peurosiva"
         DataSource      =   "Data2"
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "€+16 % IVA"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Precio en €"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Introduce la cantidad"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Data Data1 
      Caption         =   "DATA1 CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   960
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      Begin VB.Data Data5 
         Caption         =   "DATA5 FACTURAS"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Facturas"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "matricula"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "MATRICULA"
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "vehiculo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label25 
         Caption         =   "AUTOMOVIL"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5040
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "FAX"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "movil"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "MOVIL"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "tlf"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "TELEFONO"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "observaciones"
         DataSource      =   "Data1"
         Height          =   855
         Left            =   7080
         TabIndex        =   18
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label17 
         Caption         =   "OBSERVACIONES"
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "provincia"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "PROVINCIA"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "localidad"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "LOCALIDAD"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "cp"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "CP"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "apellidos"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dni"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "DIRECCION"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "APELLIDOS"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "NOMBRE"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "DNI"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label40 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7080
      TabIndex        =   37
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "NUMERO DE FACTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Dim guardar As Boolean
Dim imprimir As Boolean
Private Sub añadir_Click()
    Dim a As Currency
    Dim b As Currency
    Dim cant As Integer 'Va a contener la cantidad de elementos
    Dim numero As Currency
    Command2.Enabled = True
    If Form12.Text1.Text = "" Or Not IsNumeric(Form12.Text1.Text) Or Text3.Text = "" Or Text4.Text = "" Or Not IsNumeric(Text4.Text) Or Label39.Caption = "" Then
        MsgBox "Algunos de los campos cantidad, concepto o precio en euros estan vacíos o tienen valores NO NUMERICOS .Por favor compruebelo antes de continuar y haga el calculo del IVA antes de continuar añadiendo elementos a factura", vbExclamation, "Error de datos"
    Else
    
        cant = 0 'Inicializamos cant
        'Añadimos el cliente a la tabla auxiliar de clientes
        If bandera = False Then 'Controlamos que solo lo introduzcamos una sola vez
            Form12.Data4.Recordset.AddNew
            Form12.Data4.Recordset.Fields("numfac") = Form12.Label2.Caption
            Form12.Data4.Recordset.Fields("fecha") = Form12.Text2.Text
            Form12.Data4.Recordset.Fields("dni") = Form12.Label7.Caption
            Form12.Data4.Recordset.Fields("nombre") = Form12.Label8.Caption
            Form12.Data4.Recordset.Fields("apellidos") = Form12.Label9.Caption
            Form12.Data4.Recordset.Fields("direccion") = Form12.Label10.Caption
            Form12.Data4.Recordset.Fields("cp") = Form12.Label12.Caption
            Form12.Data4.Recordset.Fields("localidad") = Form12.Label14.Caption
            Form12.Data4.Recordset.Fields("provincia") = Form12.Label16.Caption
            Form12.Data4.Recordset.Fields("vehiculo") = Form12.Label26.Caption
            Form12.Data4.Recordset.Fields("matricula") = Form12.Label28.Caption
            Form12.Data4.Recordset.Fields("tlf") = Form12.Label20.Caption
            Form12.Data4.Recordset.Fields("fax") = Form12.Label24.Caption
            Form12.Data4.Recordset.Fields("movil") = Form12.Label22.Caption
            Form12.Data4.Recordset.Fields("observaciones") = Form12.Label18.Caption
            
            Form12.Data4.Recordset.Update
            Form12.Data4.Refresh
            bandera = True
        End If
        
        'Vamos a añadir elementos a la tabla auxiliar de facturas
    
        Form12.Data3.Recordset.AddNew
        Form12.Data3.Recordset.Fields("facnum") = Form12.Label2.Caption
        Form12.Data3.Recordset.Fields("concepto") = Form12.Text3.Text
        Form12.Data3.Recordset.Fields("cantidad") = Form12.Text1.Text
   
    
        'Vamos a hacer las operaciones matematicas para los totales
        'Las mostraremos primero por pantalla en sus correspondientes
        'textboxs y despues los meteremos en la tabla
        
        cant = Val(Form12.Text1.Text) 'Le asignamos a cant la cantidad de elementos
    
        
        'Total en Euros
        a = Form12.Text4.Text
        numero = cant * a
        teuros = teuros + numero
        euros.Text = Format(teuros, "#########.##")
        
        
        
        'Total en euros + IVA
        b = Form12.Label39.Caption
        numero = cant * b
        teurosiva = teurosiva + numero
        eurosiva.Text = Format(teurosiva, "#########.##")
    
        'Introducimos las cantidades y los precios en la tabla auxiliar
        If Text4.Text = 0 Or Label39.Caption = 0 Then 'Comprobamos que sea 0 el precio
            Form12.Data3.Recordset.Fields("peuros") = 0
            Form12.Data3.Recordset.Fields("peurosiva") = 0
        Else
            Form12.Data3.Recordset.Fields("peuros") = Format(cant * Form12.Text4.Text, "#####.##")
            Form12.Data3.Recordset.Fields("peurosiva") = Format(cant * Form12.Label39.Caption, "#####.##")
        End If
            Form12.Data3.Recordset.Update
            Form12.Data3.Refresh
            Form12.Text1.Text = 1
        
    End If
    Text3.Text = ""
    Text4.Text = ""
    Label39.Caption = ""
    añadir.Enabled = False
End Sub

Private Sub Command1_Click()
    Dim comodin As Double
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Or Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
        MsgBox "Imposible hacer calculos ya que hay algun campo vacio o erroneo. Compruebe los campos precio en euros y cantidad antes de continuar", vbExclamation, "Error de datos"
    Else
        If Text4.Text = 0 Then
            Label39.Caption = 0
        Else
            
            comodin = Text4.Text '* Text1.Text
            Label39.Caption = (comodin * 16 / 100) + comodin
            Label39.Caption = Format(Label39.Caption, "#######.##")
        End If
        añadir.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
    Load Form6
    Form6.Show vbModal
End Sub

Private Sub imprimirinforme_Click()
    If Not Form12.Data3.Recordset.BOF Then
        Form12.añadir.Enabled = False 'Desactivamos el boton añadir a factura
        Form12.eliminar.Enabled = False 'Desactivamos el boton elimianr componente de factura
        Form12.Command2.Enabled = False 'Desactivamos el boton de observaciones en factura
        guardarfactura
        guardartotales
        guardar = True
        imprimir = True
        DataEnvironment2.Connection1.Open
        DataEnvironment2.Command1
        DataReport5.Show vbModal
        DataEnvironment2.Connection1.Close
    Else
        MsgBox "No puedo imprimir nada por que no se han introducido datos en la factura", vbInformation
    End If
End Sub

Private Sub eliminar_Click()
    If Form12.Data3.Recordset.BOF Then
        MsgBox "Imposible hacer cambios ya que no hay datos en la factura", vbInformation
    Else
        Load Form10
        Form10.Show vbModal
    End If
    
End Sub

Private Sub salir_Click()
    Dim opc As Integer
    If Not Form12.Data3.Recordset.BOF Then
        If guardar = False Then
            opc = MsgBox("¿Desea guardar la factura creada?", vbOKCancel, "Pulse ACEPTAR para guardar la factura,CANCELAR para no guardarla")
            If opc = vbOK Then
                guardar = True
                guardartotales
                guardarfactura
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub nuevafactura_Click()
    Dim opc As Integer
    If Not Form12.Data3.Recordset.BOF Then
        If guardar = False Then
            opc = MsgBox("¿Desea guardar la factura?", vbOKCancel, "Elija ACEPTAR para guardar la factura o CANCEL para NO guardarla")
            If opc = vbOK Then
                guardartotales
                guardarfactura
            End If
        End If
    End If
    Unload Me
    Load Form9
    Form9.Show vbModal
End Sub


Private Sub Form_Load()
    Dim consulta As String
    Command2.Enabled = False
    guardar = False
    imprimir = False
    bandera = False
    tpesetas = 0
    teuros = 0
    tptasiva = 0
    teurosiva = 0
    Form12.añadir.Enabled = True
    Form12.Label2.Caption = numfactura
    'Las siguientes instrucciones son para buscar el cliente
    'y que lo muestre en el formulario de diseño de factura
    consulta = "select * from clientes where dni=""" + identificacion + """"
    Form12.Data1.RecordSource = consulta
    Form12.Data1.Refresh
    Form12.Text2.Text = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Borramos registro de Auxiliar de clientes
    
    borrarauxclientes
    
    'Borramos registros de auxiliar de facturas
    borrarauxfacturas
    
End Sub
Private Sub guardarfactura() 'Guarda de facturas auxiliares a Facturas
     If imprimir = False Then
        
        If Form12.Data4.Recordset.BOF = True Then
            MsgBox "No hay factura para guardar ni para imprimir.Añada elementos a factura y guarde o imprima despues", vbInformation, "Error de datos"
        Else
            Form12.Data3.Recordset.MoveFirst
            Do While Not Form12.Data3.Recordset.EOF
                Form12.Data5.Recordset.AddNew
                Form12.Data5.Recordset.Fields("nfac") = Val(Form12.Label2.Caption)
                Form12.Data5.Recordset.Fields("fecha") = Form12.Text2.Text
                Form12.Data5.Recordset.Fields("dni") = Form12.Label7.Caption
                Form12.Data5.Recordset.Fields("nombre") = Form12.Label8.Caption
                Form12.Data5.Recordset.Fields("apellidos") = Form12.Label9.Caption
                Form12.Data5.Recordset.Fields("direccion") = Form12.Label10.Caption
                Form12.Data5.Recordset.Fields("vehiculo") = Form12.Label26.Caption
                Form12.Data5.Recordset.Fields("matricula") = Form12.Label28.Caption
                Form12.Data5.Recordset.Fields("cp") = Form12.Label12.Caption
                Form12.Data5.Recordset.Fields("localidad") = Form12.Label14.Caption
                Form12.Data5.Recordset.Fields("provincia") = Form12.Label16.Caption
                Form12.Data5.Recordset.Fields("tlf") = Form12.Label20.Caption
                Form12.Data5.Recordset.Fields("movil") = Form12.Label22.Caption
                Form12.Data5.Recordset.Fields("fax") = Form12.Label24.Caption
                Form12.Data5.Recordset.Fields("observaciones") = Form12.Label12.Caption
                Form12.Data5.Recordset.Fields("concepto") = Form12.Data3.Recordset.Fields("concepto")
                Form12.Data5.Recordset.Fields("cantidad") = Form12.Data3.Recordset.Fields("cantidad")
                Form12.Data5.Recordset.Fields("peuros") = Form12.Data3.Recordset.Fields("peuros")
                Form12.Data5.Recordset.Fields("peurosiva") = Form12.Data3.Recordset.Fields("peurosiva")
                Form12.Data5.Recordset.Fields("OBSENFAC") = Form12.Data4.Recordset.Fields("OBSENFAC")
                Form12.Data5.UpdateRecord
                Form12.Data3.Recordset.MoveNext
            Loop
            MsgBox "Factura guardada sin problemas", vbInformation
            Form12.Data5.Refresh
        End If
    End If
End Sub
Private Sub borrarauxclientes()
    If Not Form12.Data4.Recordset.BOF Then
        Form12.Data4.Recordset.MoveFirst
        Form12.Data4.Recordset.Delete
        Form12.Data4.UpdateRecord
        Form12.Data4.Refresh
    End If
End Sub
Private Sub borrarauxfacturas()
    If Not Form12.Data3.Recordset.BOF Then
        Form12.Data3.Recordset.MoveFirst
        Do While Not Form12.Data3.Recordset.EOF
            Form12.Data3.Recordset.Delete
            Form12.Data3.Recordset.MoveNext
        Loop
        Form12.Data3.Refresh
    End If
End Sub
Private Sub guardartotales()
    If imprimir = False Then
        Form12.Data6.Recordset.AddNew
        Form12.Data6.Recordset.Fields("nfac") = Form12.Label2.Caption
        If euros.Text = "," Or eurosiva.Text = "," Then
            Form12.Data6.Recordset.Fields("totaleuros") = 0
            Form12.Data6.Recordset.Fields("totaleurosiva") = 0
        Else
            Form12.Data6.Recordset.Fields("totaleuros") = Form12.euros.Text
            Form12.Data6.Recordset.Fields("totaleurosiva") = Form12.eurosiva.Text
        End If
        Form12.Data6.Recordset.Update
        Form12.Data6.Refresh
    End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Not IsDate(Form12.Text2.Text) Then
        MsgBox "No reconozco este formato de fecha. Utilice el formato dia/mes/año. Por ejemplo: 10/05/2002 sería 10 de mayo del año 2002", vbExclamation, "Error de tipos de datos"
        Cancel = True
    End If
End Sub
