VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BuscarFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar clientes."
   ClientHeight    =   5628
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   13332
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5628
   ScaleWidth      =   13332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   8880
      Top             =   1560
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4935
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   13215
      _ExtentX        =   23305
      _ExtentY        =   8700
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1140
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   5835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "BuscarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rTem  As New ADODB.Recordset
Dim ContSe As Integer
Dim Entra As Boolean
Dim OpcClv As Integer
Sub Llenar_List()
  Dim x As Variant ' ListView
  If Len(Text1) < 2 Then Exit Sub
  If OpcBuscar = "DIPUTADOS" Then LlenarDip
  If OpcBuscar = "PROVEEDOR" Or OpcBuscar = "CLIENTE1" Or OpcBuscar = "CLIENTE2" Then LlenarPro
  If OpcBuscar = "PRODUCTOT" Then LlenarPT
  If OpcBuscar = "PERSONAL" Then LlenarPer
  If OpcBuscar = "VENDEDOR" Then LlenarVen
End Sub
Sub LlenarDip()
  sql = "SELECT Clave,Nombre,Telefono,mail FROM diputados WHERE Empresa=" & rEmp!Clave & " And Nombre Like '%" & Text1 & "%' ORDER BY Nombre"
  Set rTem = conn.Execute(sql)
  LV1.ListItems.Clear
  With rTem
    Do Until .EOF
      Set x = LV1.ListItems.Add()
      x.Text = !Nombre
      x.SubItems(1) = !Clave
      x.SubItems(2) = TNull(!Telefono, 1)
      x.SubItems(3) = TNull(!Mail, 1)
      .MoveNext
    Loop
    Label2 = .RecordCount
  End With

End Sub

Sub LlenarPer()
  sql = "SELECT Clave,Nombre,RFC,IMSS,Telefono,Status FROM personal WHERE Empresa=" & rEmp!Clave & " And Nombre Like '%" & Text1 & "%' ORDER BY Nombre"
  Set rTem = conn.Execute(sql)
  LV1.ListItems.Clear
  With rTem
    Do Until .EOF
      Set x = LV1.ListItems.Add()
      x.Text = !Nombre
      x.SubItems(1) = !Clave
      x.SubItems(2) = TNull(!rfc, 1)
      x.SubItems(3) = TNull(!IMSS, 1)
      x.SubItems(4) = TNull(!Telefono, 1)
      If !Status = 0 Then x.SubItems(5) = "BAJA" Else x.SubItems(5) = " "
      .MoveNext
    Loop
    Label2 = .RecordCount
  End With
End Sub

Sub LlenarPro()
  If OpcBuscar = "PROVEEDOR" Then
    noTab = "proveedor": sql2 = ""
  Else
    noTab = "cliente": sql2 = " And ClvCat=" & OpcClv
  End If
  If rEmp.EOF Then
    MsgBox "Existe un problema con los datos de la empresa.", 48, "Corporativo OCTAVA.": Exit Sub
  End If
  sql = "SELECT Clave," & nCampo & " as Nombre,Telefono1,RFC,Status FROM " & noTab & " WHERE Empresa=" & rEmp!Clave & sql2 _
  & " And " & nCampo & " Like '%" & Text1 & "%' ORDER BY " & nCampo
  Set rTem = conn.Execute(sql)
  LV1.ListItems.Clear
  With rTem
    Do Until .EOF
      Set x = LV1.ListItems.Add()
      x.Text = !Nombre
      x.SubItems(1) = !Clave
      x.SubItems(2) = TNull(!Telefono1, 1)
      x.SubItems(3) = TNull(!rfc, 1)
      If !Status = 0 Then x.SubItems(4) = "BAJA" Else x.SubItems(4) = " "
      .MoveNext
    Loop
    Label2 = .RecordCount
  End With
End Sub
Sub LlenarPT()
  If Val(Text1.Tag) > 0 Then
    sql = "SELECT a.Clave,Nombre,Codigo,Precio,Existencia,Status FROM productot a INNER JOIN listappar b ON a.Empresa=b.Empresa" _
    & " And a.Clave=b.Clave WHERE a.Empresa=" & rEmp!Clave & " And Lista=" & Text1.Tag & " And Nombre like '%" & Text1 & _
    "%' and Status ORDER BY Nombre"
    Set rTem = conn.Execute(sql)
    LV1.ListItems.Clear
    With rTem
      Do Until .EOF
        Set x = LV1.ListItems.Add()
        x.Text = !Nombre
        x.SubItems(1) = !Clave
        x.SubItems(2) = TNull(!Codigo, 1)
        x.SubItems(3) = Fmoneda(!Precio)
        x.SubItems(4) = !Existencia
        If !Status = 0 Then x.SubItems(5) = "BAJA" Else x.SubItems(5) = " "
        .MoveNext
      Loop
      Label2 = .RecordCount
    End With
  Else
    sql = "SELECT Clave,Nombre,Codigo,Costo,Existencia,Status FROM productot WHERE Empresa=" _
    & rEmp!Clave & " And Nombre Like '%" & Text1 & "%' ORDER BY Nombre"
    Set rTem = conn.Execute(sql)
    LV1.ListItems.Clear
    With rTem
      Do Until .EOF
        Set x = LV1.ListItems.Add()
        x.Text = !Nombre
        x.SubItems(1) = !Clave
        x.SubItems(2) = TNull(!Codigo, 1)
        x.SubItems(3) = Fmoneda(!Costo)
        x.SubItems(4) = !Existencia
        If !Status = 0 Then x.SubItems(5) = "BAJA" Else x.SubItems(5) = " "
        .MoveNext
      Loop
      Label2 = .RecordCount
    End With
  End If
End Sub
Sub LlenarVen()
  sql = "SELECT Clave,Nombre FROM vendedor WHERE Empresa=" & rEmp!Clave & " And Nombre Like '%" & Text1 & "%' And Status<>0 ORDER BY Nombre"
  Set rTem = conn.Execute(sql)
  LV1.ListItems.Clear
  With rTem
    Do Until .EOF
      Set x = LV1.ListItems.Add()
      x.Text = !Nombre
      x.SubItems(1) = !Clave
      .MoveNext
    Loop
    Label2 = .RecordCount
  End With
End Sub

Sub Seleccion(Lin As Variant)
  linea = Lin
  ClipText = ""
  If OpcBuscar = "VENDEDOR" Or OpcBuscar = "CODIGOSAT" Or OpcBuscar = "CUENTAS" Or OpcBuscar = "CUENTASD" Then
    Cad = LV1.ListItems(linea) & ";" & LV1.ListItems(linea).SubItems(1)
  Else
    Cad = LV1.ListItems(linea) & ";" & LV1.ListItems(linea).SubItems(1) & ";" & LV1.ListItems(linea).SubItems(2) & ";" _
    & LV1.ListItems(linea).SubItems(3)
  End If
  ClipText = Cad
  Me.Visible = False
End Sub

Private Sub Form_Load()
  ClipText = ""
'  If Menu.FacturacionMMnu.Visible Then OpcClv = 1 Else OpcClv = 2
  Formato
  Text1 = ""
End Sub
Sub Formato()
  OpcBuscar = UCase(OpcBuscar)
  Me.Caption = "Buscar " & OpcBuscar & "."
  With LV1
    .Height = Me.Height - 1000: .Left = 60: .Top = 600: .Width = Me.Width - 210
    .View = 3: .LabelEdit = 1: .FullRowSelect = True:  .ColumnHeaders.Clear:   a = 140
    If OpcBuscar = "DIPUTADOS" Then
      .ColumnHeaders.Add , , "Nombre", a * 38
      .ColumnHeaders.Add , , "Clave", a * 4, 1
      .ColumnHeaders.Add , , "Telefono", a * 15, 0
      .ColumnHeaders.Add , , "mail", a * 15, 0
    ElseIf OpcBuscar = "CODIGOSAT" Or OpcBuscar = "CUENTAS" Or OpcBuscar = "CUENTASD" Then
      .Height = Me.Height - 1000: .Left = 60: .Top = 600: .Width = Me.Width - 210
      .View = 3: .LabelEdit = 1: .FullRowSelect = True:  .ColumnHeaders.Clear:   a = 140
      .ColumnHeaders.Add , , "Codigo", a * 16
      .ColumnHeaders.Add , , "Cuenta", a * 60, 0
    ElseIf OpcBuscar = "PROVEEDOR" Or OpcBuscar = "CLIENTE1" Or OpcBuscar = "CLIENTE2" Then
      .ColumnHeaders.Add , , "Nombre", a * 40
      .ColumnHeaders.Add , , "Clave", a * 8, 1
      .ColumnHeaders.Add , , "Telefono", a * 22, 0
      .ColumnHeaders.Add , , "RFC", a * 15, 0
      .ColumnHeaders.Add , , "Stat", a * 6, 1
    ElseIf OpcBuscar = "PRODUCTOSF" Then
      .ColumnHeaders.Add , , "Nombre", a * 40
      .ColumnHeaders.Add , , "Clave", a * 8, 1
      .ColumnHeaders.Add , , "Codigo", a * 15, 0
      .ColumnHeaders.Add , , "Precio", a * 7, 0
      .ColumnHeaders.Add , , "Existencia", a * 7, 1
      .ColumnHeaders.Add , , "Status", a * 4, 1
        
    ElseIf OpcBuscar = "PRODUCTOS" Then
      .ColumnHeaders.Add , , "Nombre", a * 40
      .ColumnHeaders.Add , , "Clave", a * 8, 1
      .ColumnHeaders.Add , , "Codigo", a * 15, 0
      .ColumnHeaders.Add , , "UCom", a * 6, 0
      .ColumnHeaders.Add , , "Factor", a * 4, 1
      .ColumnHeaders.Add , , "Costo", a * 7, 1
      .ColumnHeaders.Add , , "Exist.", a * 7, 1
      .ColumnHeaders.Add , , "Status", a * 4, 1
    ElseIf OpcBuscar = "PRODUCTOT" Then
      .ColumnHeaders.Add , , "Nombre", a * 40
      .ColumnHeaders.Add , , "Clave", a * 8, 1
      .ColumnHeaders.Add , , "Codigo", a * 15, 0
      .ColumnHeaders.Add , , "Precio", a * 7, 1
      .ColumnHeaders.Add , , "Exist.", a * 7, 1
      .ColumnHeaders.Add , , "Status", a * 4, 1
    ElseIf OpcBuscar = "PERSONAL" Then
      .ColumnHeaders.Add , , "Nombre", a * 38
      .ColumnHeaders.Add , , "Clave", a * 8, 1
      .ColumnHeaders.Add , , "RFC", a * 13, 0
      .ColumnHeaders.Add , , "IMSS", a * 13, 0
      .ColumnHeaders.Add , , "Telefono", a * 17, 0
      .ColumnHeaders.Add , , "Status", a * 7, 1
    ElseIf OpcBuscar = "VENDEDOR" Then
      .ColumnHeaders.Add , , "Nombre", a * 42
      .ColumnHeaders.Add , , "Clave", a * 8, 1
    End If
  End With
End Sub

Private Sub LV1_DblClick()
  If LV1.ListItems.Count <= 0 Then Exit Sub
  linea = LV1.SelectedItem.Index
  Seleccion linea
End Sub

Private Sub LV1_GotFocus()
  ContSe = 1
  Timer1.Interval = 0
End Sub


Private Sub LV1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If LV1.ListItems.Count <= 0 Then Exit Sub
    KeyAscii = 0
    linea = LV1.SelectedItem.Index
    Seleccion linea
  ElseIf KeyAscii = 27 Then
    ClipText = ""
    Unload Me
  End If
End Sub

Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then Text1.SetFocus
  If KeyCode = 37 Or KeyCode = 39 Then Text1.SetFocus
End Sub



Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
  Text1.backcolor = QBColor(11)
  If OpcBuscar = "CLIENTES" Then
    If Text1.Tag = "1" Then
      Me.Caption = "Buscar cliente."
    ElseIf Text1.Tag = "2" Then
      Me.Caption = "Buscar cliente(Pz)"
    End If
  Else
    If Text1.Tag = "5" Then Me.Caption = "Buscar maquila"
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If LV1.ListItems.Count = 1 Then
      ClipText = ""
      Seleccion 1
      KeyAscii = 0
    ElseIf LV1.ListItems.Count > 1 Then
      SendKeys "{TAB}"
      KeyAscii = 0
    End If
  ElseIf KeyAscii = 27 Then
    ClipText = ""
    Unload Me
  End If
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then
    LV1.SetFocus
  Else
    If KeyCode <> 13 Then
      Timer1.Interval = 900: ContSe = 1
    End If
  End If
End Sub


Private Sub Text1_LostFocus()
  Text1.backcolor = QBColor(15)
End Sub
Private Sub Timer1_Timer()
  ContSe = ContSe - 1
  If ContSe <= 0 Then
    ContSe = 0
    Timer1.Interval = 0
    Llenar_List
  End If
End Sub


