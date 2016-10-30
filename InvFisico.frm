VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form InvFisicoFrm 
   Caption         =   "Inventario Fisico."
   ClientHeight    =   6324
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   16296
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6324
   ScaleWidth      =   16296
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir"
      Height          =   3192
      Left            =   5880
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   3432
      Begin VB.CommandButton Command3 
         Caption         =   " Cerrar"
         Height          =   612
         Index           =   1
         Left            =   1860
         TabIndex        =   15
         Top             =   2280
         Width           =   1392
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   612
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1392
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Costo inventario"
         Height          =   252
         Index           =   1
         Left            =   300
         TabIndex        =   13
         Top             =   1200
         Width           =   2172
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Etiquetas"
         Height          =   252
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   660
         Value           =   -1  'True
         Width           =   2172
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   13440
      Top             =   240
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "Foto"
      Height          =   1212
      Left            =   8040
      ScaleHeight     =   1164
      ScaleWidth      =   1224
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1272
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   492
      Index           =   1
      Left            =   12000
      TabIndex        =   9
      Top             =   60
      Width           =   1272
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   492
      Index           =   0
      Left            =   10500
      TabIndex        =   8
      Top             =   60
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Index           =   3
      Left            =   6840
      TabIndex        =   7
      Top             =   60
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   60
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   60
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   60
      Width           =   1272
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4848
      Left            =   2220
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   3492
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5052
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   16152
      _ExtentX        =   28490
      _ExtentY        =   8911
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1140
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   792
   End
End
Attribute VB_Name = "InvFisicoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rUbi As ADODB.Recordset
Dim rTem As ADODB.Recordset
Dim Primera As Boolean
Dim CA As Integer
Dim CN As Integer
Sub Activo(Opc)
  For a = 0 To 3
    Command1(a).Visible = Not Opc
  Next
  Text1.Enabled = Not Opc
  Grid1.Enabled = Opc
  Command2(0).Visible = Opc
  Command2(1).Visible = Opc
End Sub

Sub Consulta()
  sql = "SELECT c.Nombre AS Ubica,ClvArt,Descripcion,d.Nombre AS Clasif,Cantidad,Serie,Foto,Ubicacion,a.Costo FROM ((invfisico a INNER " _
  & "JOIN inventario b ON a.Empresa=b.Empresa AND ClvArt=b.Clave) INNER JOIN ubicacion c ON a.Empresa=c.Empresa AND a.Ubicacion=c.Clave)" _
  & " LEFT JOIN comprascla d ON a.Empresa=d.Empresa AND b.Clasif=d.Clave WHERE Fecha=" & Fecha6(Text1, 3)
  Set rTem = conn.Execute(sql)
  Set Picture1.DataSource = rTem
  Grid1.Rows = 1: Grid1.Redraw = False
  With rTem
  Do Until .EOF
    Cad = !Ubica & L & !ClvArt & L & !Descripcion & L & !Clasif & L & !Cantidad & L & !Serie & L & L & !Ubicacion & L & !Costo
    Grid1.AddItem Cad
    Grid1.Row = Grid1.Rows - 1: Grid1.Col = 6
    Grid1.RowHeight(Grid1.Row) = 2000
    Set Grid1.CellPicture = Picture1.Picture
    .MoveNext
  Loop
  End With
  Grid1.Redraw = True
End Sub

Sub Etiqueta()
  Dim rEti As ADODB.Recordset
  sql = "SELECT * FROM invfisico WHERE Fecha=" & Fecha6(Text1, 3)
  Set rTem = conn.Execute(sql)
  conn.Execute "DELETE FROM Etiqueta"
  Do Until rTem.EOF
    For a = 1 To rTem!Cantidad
      sql = "INSERT INTO etiqueta(Empresa,ClvArt,Fecha,Serie) VALUES(" & rEmp!Clave & "," & rTem!ClvArt & "," & Fecha6(Text1, 3) & ",'" _
      & rTem!Serie & "')"
      conn.Execute sql
    Next
    rTem.MoveNext
  Loop
End Sub

Sub MostrarA(Clv)
  If Val(Clv) = 0 Then Exit Sub
  sql = "SELECT a.Clave,Descripcion,Clasif,Nombre,Foto,Costo FROM inventario a LEFT JOIN comprascla b ON a.Empresa=b.Empresa and " _
  & "a.Clasif=b.Clave WHERE a.Empresa=" & rEmp!Clave & " And a.Clave=" & Clv
  Set rTem = conn.Execute(sql)
  nl = Grid1.Row
  If rTem.RecordCount = 0 Then
    Grid1.TextMatrix(nl, 2) = "NO EXISTE"
  Else
    Grid1.TextMatrix(nl, 1) = rTem!Clave: Grid1.TextMatrix(nl, 2) = rTem!Descripcion: Grid1.TextMatrix(nl, 3) = TNull(rTem!Nombre, 0)
    Grid1.TextMatrix(nl, 8) = rTem!Costo
    If Not IsNull(rTem!Foto) Then
      Set Picture1.DataSource = rTem
      Grid1.Col = 6
      Set Grid1.CellPicture = Picture1.Picture
    End If
    Grid1.Col = 4
  End If
End Sub

Private Sub Command1_Click(Index As Integer)
  If Index = 0 Then
    Activo True
    Grid1.Rows = 1: Grid1.AddItem "": Grid1.RowHeight(Grid1.Rows - 1) = 2000: Grid1.Col = 0
    Grid1.SetFocus
  ElseIf Index = 1 Then
    Consulta
    Activo True
  ElseIf Index = 2 Then
    Frame1.Visible = True
    Command3(0).SetFocus
  Else
    Unload Me
  End If
End Sub
Private Sub Command2_Click(Index As Integer)
  If Index = 0 Then
    If Not IsDate(Text1) Then
      MsgBox "El formato de la fecha es incorrecto.", 48
      Text1.SetFocus: Exit Sub
    End If
    conn.Execute "DELETE FROM invfisico WHERE Fecha=" & Fecha6(Text1, 3)
    For a = 1 To Grid1.Rows - 1
      Cant = Val(Grid1.TextMatrix(a, 4)): ClvA = Val(Grid1.TextMatrix(a, 1))
      If Cant > 0 And ClvA > 0 Then
        ClvU = Val(Grid1.TextMatrix(a, 7))
        sql = "INSERT INTO invfisico VALUES(" & rEmp!Clave & "," & ClvU & "," & ClvA & "," & Fecha6(Text1, 3) & "," & Cant & ",'" & Ns _
        & "'," & Val(Grid1.TextMatrix(a, 8)) & ")"
        conn.Execute sql
      End If
    Next
  End If
  Activo False
End Sub

Private Sub Command3_Click(Index As Integer)
  If Index = 0 Then
    If Option1(0) Then
      Etiqueta
      ChDir App.Path
      MsgBox "Pulse <ENTER> para ver las etiquetas."
      Report1.ReportFileName = "Etiqueta.rpt"
      Report1.Destination = 0
      Report1.Action = 0
    Else
      Report1.ReportFileName = "Inventario.rpt"
      Report1.Destination = 0
      Report1.Action = 0
    End If
  End If
  Frame1.Visible = False
End Sub

Private Sub Form_Load()
  Report1.WindowState = 2
  Report1.WindowTitle = "Congreso del estado"
  Me.Height = Menu.Height - 1200: Me.Width = 16392
  CentrarFrm Me
  Grid1.Height = Me.Height - Grid1.Top - 1000
  Text1 = date
  Grid Grid1, "Ubicación,Artículo,Descripcion,Clasif,Cant,N. Serie,Imagen,ClvU,Costo", "15,8,45,12,9,25,22,.001,.001", "1,1,1,1,4"
  Set rUbi = conn.Execute("SELECT * FROM ubicacion")
  Do Until rUbi.EOF
    If Not IsNull(rUbi!Nombre) Then List1.AddItem rUbi!Nombre
    rUbi.MoveNext
  Loop
  Activo False
End Sub


Private Sub Form_Resize()
  If Me.WindowState = 1 Then Exit Sub
  Grid1.Width = Me.Width - 350
  Grid1.Height = Me.Height - 1200
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Grid1.Col = 0 Then
      List1.Visible = True: Grid1.Enabled = False
      If List1.ListIndex < 0 Then
        List1.ListIndex = 0
      End If
      List1.SetFocus
    ElseIf Grid1.Col = 1 Then
      If Val(Grid1) = 0 Then
        'Mostra
      Else
        MostrarA Val(Grid1)
      End If
    ElseIf Grid1.Col = 4 Then
      Grid1.Col = 5
    ElseIf Grid1.Col = 5 Then
      If Grid1.Row = Grid1.Rows - 1 Then
        Grid1.AddItem "": Grid1.Row = Grid1.Rows - 1: Grid1.RowHeight(Grid1.Rows - 1) = 2000: Grid1.Col = 0
      Else
        Grid1.Row = Grid1.Row + 1: Grid1.Col = 0
      End If
    End If
  Else
    If Grid1.Col = 1 Or Grid1.Col = 4 Or Grid1.Col = 5 Then
      If KeyAscii <> 13 Then
        If KeyAscii = 8 Then
          If Len(Grid1) > 0 Then Grid1 = Mid(Grid1, 1, Len(Grid1) - 1)
        Else
          If Primera Then
            Grid1 = Chr(KeyAscii)
            Primera = False
          Else
            Grid1 = Grid1 & Chr(KeyAscii)
          End If
        End If
      End If
    End If
  End If
End Sub


Private Sub Grid1_LeaveCell()
  If Grid1.Row > 0 Then Grid1.CellBackColor = rEmp!GridL2
End Sub

Private Sub Grid1_RowColChange()
  If Grid1.Row = 0 Then Exit Sub
  Primera = True
  Grid1.CellBackColor = rEmp!GRIDL1
End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If List1 = "" Then Exit Sub
    If rUbi.RecordCount = 0 Then Exit Sub
    rUbi.MoveFirst
    rUbi.Find "Nombre='" & List1 & "'"
    Grid1 = List1
    Grid1.TextMatrix(Grid1.Row, 7) = rUbi!Clave
    List1.Visible = False
    Grid1.Col = 1: Grid1.Enabled = True
    Grid1.SetFocus
  ElseIf KeyAscii = 27 Then
    List1.Visible = False: Grid1.Enabled = True
    Grid1.SetFocus
  End If
End Sub

Private Sub Text1_GotFocus()
  TextG Text1
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub


Private Sub Text1_LostFocus()
  TextL Text1
End Sub


