VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form InventarFrm 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descripción articulos."
   ClientHeight    =   6204
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6204
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   672
      Index           =   8
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5400
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   9720
      Top             =   4440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir."
      Height          =   3192
      Left            =   5640
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   3672
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   2040
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   435
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credencial todos"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   24
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credencial individual"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   23
         Top             =   900
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   22
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   300
      ScaleHeight     =   564
      ScaleWidth      =   8964
      TabIndex        =   15
      Tag             =   "Dinero.bmp"
      Top             =   4500
      Width           =   9015
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Regresa al primer registro."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Regresa al registro anterior."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   2
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Avanza al siguiente registro."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   3
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Avanza al ultimo registro."
         Top             =   0
         Width           =   615
      End
      Begin VB.Label LabelDatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro 0 de 0."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   192
         Left            =   1440
         TabIndex        =   20
         Top             =   180
         Width           =   4680
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox contenido 
      BorderStyle     =   0  'None
      Height          =   5172
      Left            =   -12
      ScaleHeight     =   5172
      ScaleWidth      =   10692
      TabIndex        =   27
      Top             =   0
      Width           =   10692
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   2
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "mas.bmp"
         ToolTipText     =   "Cancela los cambios echos al registro."
         Top             =   1080
         Width           =   315
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Foto"
         Height          =   2652
         Left            =   7980
         ScaleHeight     =   2652
         ScaleWidth      =   2412
         TabIndex        =   35
         Top             =   1560
         Width           =   2412
      End
      Begin VB.TextBox Text1 
         DataField       =   "Clave"
         DataMember      =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   300
         Width           =   1152
      End
      Begin VB.TextBox Text1 
         DataField       =   "Descripcion"
         DataMember      =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   7392
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "Existencia"
         DataMember      =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9060
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   300
         Width           =   1272
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "Costo"
         DataMember      =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   1272
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Clasif"
         DataMember      =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   0
         Left            =   1620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "comprascla"
         Top             =   1080
         Width           =   2952
      End
      Begin VB.TextBox Text1 
         DataField       =   "Obser"
         DataMember      =   "T"
         Height          =   1908
         Index           =   16
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1920
         Width           =   7692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         Height          =   252
         Index           =   0
         Left            =   180
         TabIndex        =   33
         Top             =   60
         Width           =   1152
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Height          =   252
         Index           =   1
         Left            =   1500
         TabIndex        =   32
         Top             =   60
         Width           =   7392
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Existencia"
         Height          =   252
         Index           =   2
         Left            =   9060
         TabIndex        =   31
         Top             =   60
         Width           =   1272
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo"
         Height          =   252
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   840
         Width           =   1272
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clasificación"
         Height          =   252
         Index           =   10
         Left            =   1620
         TabIndex        =   29
         Top             =   840
         Width           =   2952
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones:"
         Height          =   312
         Index           =   13
         Left            =   120
         TabIndex        =   28
         Top             =   1620
         Width           =   1452
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fotografía"
      Height          =   672
      Index           =   7
      Left            =   7980
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "Camara.jpg"
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   615
      Index           =   1
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1512
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   615
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1512
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar &Nombre"
      Height          =   672
      Index           =   5
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar &Clave"
      Height          =   672
      Index           =   4
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   672
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Borrar"
      Height          =   672
      Index           =   2
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   672
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   672
      Index           =   0
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1212
   End
End
Attribute VB_Name = "InventarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rAlu As ADODB.Recordset
Dim rTem As ADODB.Recordset
Dim rCar As ADODB.Recordset
Dim rGra As ADODB.Recordset
Dim rGru As ADODB.Recordset
Dim rCag As ADODB.Recordset

Dim Mayor As Long
Dim NR As Long
Dim NrA As Long
Sub Posicion(NDoc As Variant, TipoM)
  If NR = 0 Then
    LabelDatos = "Sin registros": Exit Sub
  End If
  Dim tIma As ADODB.Stream
  Set tIma = New ADODB.Stream
  tIma.Type = adTypeBinary
  If IsNull(NDoc) Then Exit Sub
  If NDoc = 0 Or NDoc > Mayor Then Exit Sub
  nclv = CPosicion(NDoc, TipoM, "inventario", "Clave", 0)
  LabelDatos = "Total inventario: " & NR
  Set rAlu = conn.Execute("SELECT * FROM inventario WHERE Empresa=" & rEmp!Clave & " And Clave=" & nclv)
  If rAlu.RecordCount = 0 Then
    MsgBox "No existe la clave " & nclv, 48, rEmp!Nombre: Exit Sub
  Else
    MostrarTexN Me, rAlu
'    If rAlu!Status = 0 Then Me.backcolor = QBColor(12) Else Me.backcolor = rEmp!Fondo
  End If
End Sub
Private Sub Combo1_Click(Index As Integer)
  Combo1(Index).ToolTipText = "1"
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub Command1_Click(Index As Integer)
  If NR = 0 Then
    If Index >= 1 And Index <= 7 Then Exit Sub
  End If
  If Index = 0 Then                     'Agregar
    Limpiar Me
    Activa Me, True
    sql = "SELECT Max(Clave) as NMa FROM inventario WHERE Empresa=" & rEmp!Clave
    Set rTem = conn.Execute(sql)
    Text1(0) = Valor(rTem!NMa) + 1
    Command2(0).Tag = "A"
    Picture2.Picture = LoadPicture()
    Check1 = 1
    Text1(0).Enabled = True: Text1(0).SetFocus
  ElseIf Index = 1 Then                 'Editar
    If Val(Text1(0)) = 0 Then
      MsgBox "No existen datos para editar.", 48, rEmp!Nombre
      Exit Sub
    End If
    Activa Me, True
    Command2(0).Tag = "E"
    Text1(0).Enabled = False
    Text1(1).SetFocus
  ElseIf Index = 2 Then                 'Borrar
    If Val(Text1(0)) = 0 Then
      MsgBox "No existen datos para borrar.", 48, rEmp!Nombre
      Exit Sub
    End If
    If MsgBox("Esta seguro de borrar la clave " & Text1(0), 36, rEmp!Nombre) <> 6 Then Exit Sub
    conn.Execute "DELETE FROM inventario WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
    NR = NR - 1
    If NR <= 0 Then
      Limpiar Me
    Else
      Posicion Val(Text1(0)) - 1, "<"
    End If
  ElseIf Index = 3 Then                 'Imprimir
    Frame1.Visible = True
  ElseIf Index = 4 Then                 'Busca Clv
    r = InputBox("Ingrese el numero de clave.", Trim(rEmp!Nombre))
    If Val(r) = 0 Then Exit Sub
    Set rTem = conn.Execute("SELECT Clave FROM inventario WHERE Empresa=" & rEmp!Clave & " And Clave=" & Val(r))
    If rTem.RecordCount = 0 Then
      MsgBox "No existe la clave " & r, 48, rEmp!Nombre
    Else
      Posicion Val(r), ""
    End If
  ElseIf Index = 5 Then                 'Busca Nombre
    Busca
  ElseIf Index = 7 Then
    On Error Resume Next
    With Dialog1
      .Filter = "Imagenes (*.jpg)|*.jpg"
      .DialogTitle = "Seleccione la imagen"
      .ShowOpen
      If (.Filename = "") Then Exit Sub
      SaveFoto Me, Picture2, "", .Filename, False, 0, 0
    End With
    Set rTem = New ADODB.Recordset
    rTem.Open "SELECT * FROM inventario WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0), conn, adOpenDynamic, adLockOptimistic
    rTem!Foto = ImaD
    rTem.Update
    If Err = -2147217864 Then Err = 0
  ElseIf Index = 8 Then
    Unload Me
  End If
End Sub
Sub Busca()
  If OpcBuscar <> "inventario" Then If rFormulario(BuscarFrm) Then Unload BuscarFrm
  OpcBuscar = "inventario"
  BuscarFrm.Show 1
  Cad = ClipText
  If InStr(Cad, ";") > 0 Then
    ArraT = Split(Cad, ";")
    Posicion ArraT(1), ""
  End If
End Sub
Private Sub Command2_Click(Index As Integer)
  If Index = 0 Then
    If Val(Text1(0)) = 0 Then
      MsgBox "El número de clave es obligatorio.", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
    End If
    Text1(1) = Trim(Text1(1)): Text1(2) = Trim(Text1(2)): Text1(3) = Trim(Text1(3))
    If Command2(0).Tag = "A" Then
      Set rTem = conn.Execute("SELECT Clave FROM inventario WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0))
      If rTem.RecordCount > 0 Then
        MsgBox "El número de clave ya existe.", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
      End If
      Set rTem = conn.Execute("SELECT clave FROM inventario WHERE Empresa=" & rEmp!Clave & " And Descripcion='" & Text1(1) & "'")
      If rTem.RecordCount > 0 Then
        MsgBox "El nombre del alumno ya existe.", 48, rEmp!Nombre: Text1(1).SetFocus: Exit Sub
      End If
      sql = SqlCad(Me, "inventario", "A", "")
      conn.Execute sql
      NR = NR + 1: If Val(Text1(0)) > Mayor Then Mayor = Val(Text1(0))
      Posicion Text1(0), ""
    Else
      sql = SqlCad(Me, "inventario", "E", "Empresa=" & rEmp!Clave & " And Clave=" & Text1(0))
      If Len(sql) > 0 Then conn.Execute sql
    End If
    Activa Me, False
  ElseIf Index = 1 Then
    If NR = 0 Then Limpiar Me Else Posicion rAlu!Clave, ""
    Activa Me, False
  ElseIf Index = 2 Then
    ClaveCat = 17
    AltasFrm.Show 1
    If InStr(ClipText, ";") > 0 Then
      cam = Split(ClipText, ";")
      Combo1(0).AddItem cam(1)
      Combo1(0) = cam(1)
    End If
  
  End If
End Sub
Private Sub Command3_Click(Index As Integer)
  If Index = 0 Then
    tEmp = "": tTit = ""
    If Option1(0) Then
      Arch = "Alumnos.Rpt": sql = "{alumnos.Empresa}=" & rEmp!Clave: tEmp = rEmp!Nombre: tTit = "Listado de alumnos"
    ElseIf Option1(1) Then
      Arch = "Credencial.Rpt": sql = "{alumnos.Empresa}=" & rEmp!Clave & " And {alumnos.Matricula}=" & Val(Text1(0))
    ElseIf Option1(2) Then
      Arch = "Credencial.Rpt": sql = "{alumnos.Empresa}=" & rEmp!Clave
    End If
    Menu.Report1Im Arch, tEmp, tTit, sql, "", ""
  End If
  Frame1.Visible = False
End Sub
Sub Report1Im(nArch, nEmpr, nTitu, nSele, nOrde, nFor3)
  MousePointer = 11
  For a = 0 To 10
    Report1.Formulas(a) = ""
    Report1.SortFields(a) = ""
  Next
  'If Option1(1) Then Report1.PrinterSelect
  ChDir App.Path
  Report1.SelectionFormula = ""
  Report1.ReportFileName = DirRep & nArch
  CRConn = "DSN=Escu; UID=vicsol; PWD=Admin2012"  'DSQ =" & Directorio & "empresa1.dsn"

  Report1.Connect = CRConn
  If Len(nEmpr) = 0 Then Report1.Formulas(0) = "" Else Report1.Formulas(0) = "Empresa='" & Trim(nEmpr) & "'"
  If Len(nTitu) = 0 Then
    Report1.Formulas(1) = "": Report1.Formulas(2) = ""
  Else
    Report1.Formulas(1) = "Titulo='" & nTitu & "'"
  End If
  Report1.Formulas(2) = "Archivo='" & nArch & "'"
  Report1.Formulas(3) = nFor3
  Report1.SelectionFormula = nSele
  Report1.SortFields(0) = nOrde
  Report1.Destination = 0
  Report1.Action = 0
  MousePointer = 0
End Sub

Private Sub Datos_Click(Index As Integer)
  Select Case Index
    Case 0: Posicion -2, ""
    Case 1: Posicion Val(Text1(0)) - 1, "<"
    Case 2: Posicion Val(Text1(0)) + 1, ">"
    Case 3: Posicion -1, ""
  End Select
End Sub

Private Sub Form_Load()
  Limpiar Me
  Colores Me
  Frame1.Height = 2955: Frame1.Left = 3240: Frame1.Top = 1020: Frame1.Width = 3855
  Me.Height = 6576: Me.Width = 10932
  BotonPic Me
  
  TextLon Me, "inventario"   'Este pone los colores
  CentrarFrm Me
  Me.Top = 0
  Activa Me, False
  Set rTem = conn.Execute("SELECT count(*) AS TRe FROM inventario WHERE Empresa=" & rEmp!Clave)
  NR = Valor(rTem!tRe): Mayor = 0
  If NR > 0 Then
    Set rTem = conn.Execute("SELECT Max(Clave) as Clave FROM inventario WHERE Empresa=" & rEmp!Clave)
    Mayor = Valor(rTem!Clave)
  Else
    Set rAlu = conn.Execute("SELECT * FROM inventario WHERE Empresa=" & rEmp!Clave)
  End If
  Set rTem = conn.Execute("SELECT count(*) AS TRe FROM inventario WHERE Empresa=" & rEmp!Clave)
  NrA = Valor(rTem!tRe)
  Posicion -1, ""
End Sub
Private Sub Text1_Change(Index As Integer)
  Text1(Index).ToolTipText = "1"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
  TextG Text1(Index)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
  TextL Text1(Index)
End Sub


