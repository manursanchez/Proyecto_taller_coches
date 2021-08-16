VERSION 5.00
Begin VB.Form Form20 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Numero de presupuesto a recuperar"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "Recuperar presupuesto.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "DATA 1 PRESUPUESTOS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuestos"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Data Data2 
      Caption         =   "DATA 2 AUXILIAR DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Data Data3 
      Caption         =   "DATA 3 AUXILIAR DE FACTURAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 TOTALPRESUPUESTOS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "totalespre"
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce el numero de presupuesto a recuperar:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     'Buscamos el presupuesto para ver si existe o no
    Dim consulta As String
    Dim consultat As String
    consulta = "select * from presupuestos where npresupuesto=""" + Form20.Text1.Text + """"
    Form20.Data1.RecordSource = consulta
    Form20.Data1.Refresh
    If Form20.Data1.Recordset.RecordCount = 1 Then
        'Vamos a guardar el cliente en auxiliar de clientes
        
        Form20.Data2.Recordset.AddNew
        Form20.Data2.Recordset.Fields("numfac") = Form20.Data1.Recordset("npresupuesto")
        Form20.Data2.Recordset.Fields("fecha") = Form20.Data1.Recordset("fecha")
        Form20.Data2.Recordset.Fields("dni") = Form20.Data1.Recordset("dni")
        Form20.Data2.Recordset.Fields("nombre") = Form20.Data1.Recordset("nombre")
        Form20.Data2.Recordset.Fields("apellidos") = Form20.Data1.Recordset("apellidos")
        Form20.Data2.Recordset.Fields("direccion") = Form20.Data1.Recordset("direccion")
        Form20.Data2.Recordset.Fields("cp") = Form20.Data1.Recordset("cp")
        Form20.Data2.Recordset.Fields("localidad") = Form20.Data1.Recordset("localidad")
        Form20.Data2.Recordset.Fields("provincia") = Form20.Data1.Recordset("provincia")
        Form20.Data2.Recordset.Fields("vehiculo") = Form20.Data1.Recordset("vehiculo")
        Form20.Data2.Recordset.Fields("matricula") = Form20.Data1.Recordset("matricula")
        Form20.Data2.Recordset.Fields("tlf") = Form20.Data1.Recordset("tlf")
        Form20.Data2.Recordset.Fields("movil") = Form20.Data1.Recordset("movil")
        Form20.Data2.Recordset.Fields("fax") = Form20.Data1.Recordset("fax")
        Form20.Data2.Recordset.Fields("observaciones") = Form20.Data1.Recordset("observaciones")
        Form20.Data2.Recordset.Fields("OBSENFAC") = Form20.Data1.Recordset.Fields("obsenpres")
        Form20.Data2.UpdateRecord
        Form20.Data2.Refresh
        
        'Vamos a añadir todos los elementos del presupuesto a Auxiliar de Facturas
        
        Form20.Data1.Recordset.MoveFirst
        Form20.Data1.Refresh
        Do While Not Form20.Data1.Recordset.EOF
           If Form20.Data1.Recordset.Fields("npresupuesto") = Form20.Text1.Text Then
                Form20.Data3.Recordset.AddNew
                Form20.Data3.Recordset.Fields("facnum") = Form20.Data1.Recordset.Fields("npresupuesto")
                
                Form20.Data3.Recordset.Fields("concepto") = Form20.Data1.Recordset.Fields("concepto")
                Form20.Data3.Recordset.Fields("cantidad") = Form20.Data1.Recordset.Fields("cantidad")
                Form20.Data3.Recordset.Fields("peuros") = Form20.Data1.Recordset.Fields("peuros")
                
                Form20.Data3.Recordset.Fields("peurosiva") = Form20.Data1.Recordset.Fields("peurosiva")
                
                Form20.Data3.UpdateRecord
            End If
            Form20.Data1.Recordset.MoveNext
         Loop
         consultat = "select * from totalespre where npresupuesto=""" + Form20.Text1.Text + """"
         Form20.Data4.RecordSource = consultat
         Form20.Data4.Refresh
         If Data4.Recordset.RecordCount = 1 Then
         
           
            teuros = Form20.Data4.Recordset.Fields("teuros")
            
            teurosiva = Form20.Data4.Recordset.Fields("teurosiva")
            Unload Form20
            Load Form21
            Form21.Show vbModal
         Else
            MsgBox "No encuentro los totales del presupuesto", vbCritical, "Error en la búsqueda"
         End If
    Else
        MsgBox "El número de presupuesto introducido no existe.Introduzca un nuevo número de factura", vbInformation, "Factura inexistente"
    End If
        
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
