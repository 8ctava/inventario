VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CatalogosFrm 
   Caption         =   "Catalogos."
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11016
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11016
   Begin VB.Frame Frame2 
      Caption         =   "Impresión saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1752
      Left            =   3000
      TabIndex        =   22
      Top             =   1140
      Visible         =   0   'False
      Width           =   3312
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "Cancel.bmp"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "Print.jpg"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   2700
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   420
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   2700
         TabIndex        =   25
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   420
         TabIndex        =   23
         Top             =   540
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   4380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   795
         Index           =   1
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "cancel.bmp"
         Top             =   1680
         Width           =   1635
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         Height          =   795
         Index           =   0
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "ok.bmp"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1020
         Width           =   5052
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S/Cobrar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   17
         Top             =   2220
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   16
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta<F2>:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   300
         TabIndex        =   15
         Top             =   1620
         Visible         =   0   'False
         Width           =   1152
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      Height          =   5712
      Left            =   1020
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   780
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   855
      Index           =   5
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B&uscar"
      Height          =   855
      Index           =   4
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3660
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   855
      Index           =   3
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Borrar"
      Height          =   855
      Index           =   2
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1860
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   855
      Index           =   1
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   855
      Index           =   0
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9015
      _ExtentX        =   15896
      _ExtentY        =   12933
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   9120
      TabIndex        =   21
      Top             =   5460
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "CatalogosFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rTem As ADODB.Recordset
Dim nTabla As String
Sub Formato()
  If ClaveCatC = 2 Then    'Concepto Cargo/Abono
    Cad = ",Clave,Nombre,Cuenta,Saldo,S/Cobrar": cad1 = ".5,5,26,11,11,11"
    For a = 2 To 4
      Label1(a).Visible = True: Text1(a).Visible = True
    Next
  ElseIf ClaveCatC = 5 Or ClaveCatC = 6 Or ClaveCatC = 15 Then  'Compras Clasificacion, 15=Vales
    Cad = ",Clave,Nombre,Cuenta": cad1 = "1,5,40,18"
  ElseIf ClaveCatC = 7 Then    'Monedas
    Cad = ",Clave,Nombre,T. Cambio": cad1 = "1,3,40,18": Label1(2) = "T.Cambio:"
  ElseIf ClaveCatC = 9 Then    'Deptos
    Cad = ",Clave,Nombre,M.O.": cad1 = "1,3,40,10":  Label1(2) = "M. Obra:"
  ElseIf ClaveCatC = 11 Then    'Vendedor
    Cad = ",Clave,Nombre,Stat": cad1 = ".5,5,34,4": Label1(2) = "Status"
  ElseIf ClaveCatC = 16 Then    'Metodo de pago
    Cad = ",Clave,Nombre,ClvCue": cad1 = ".5,5,40,10": Label1(2) = "ClvCue"
  ElseIf ClaveCatC = 17 Then    'Cargos Colegiaturas
    Cad = ",Clave,Nombre,Monto,Unico": cad1 = ".5,5,40,10,4"
    For a = 2 To 3
      Label1(a).Visible = True: Text1(a).Visible = True
    Next
    Label1(2) = "Monto": Label1(3) = "Unico"
  ElseIf ClaveCatC = 18 Then    'Carreras o Nivel
    Cad = ",Clave,Nombre,RVOE": cad1 = ".5,5,40,14": Label1(2) = "RVOE"
  Else
    Cad = ",Clave,Nombre": cad1 = "1,7,50"
  End If
  Select Case ClaveCatC
    Case 5, 6, 7, 9, 15, 11, 16, 18
      Label1(2).Visible = True: Text1(2).Visible = True
  End Select
  Grid Grid1, Cad, cad1, "0,4,0"
End Sub
Sub Mostrar()
  sql = "SELECT * FROM " & nTabla
'  If ClaveCatC <> 7 And ClaveCatC <> 16 Then SQL = SQL & " WHERE Empresa=" & rEmp!Clave
  If ClaveCatC = 18 Then sql = sql & " WHERE Empresa=" & rEmp!Clave
  Set rTem = conn.Execute(sql)
  If ClaveCatC = 2 Then
    If rTem.RecordCount = 0 Then
      conn.Execute "INSERT INTO conceptosca(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",1,'CAJA')"
      conn.Execute "INSERT INTO conceptosca(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",2,'NOTAS DE CREDITO')"
      rTem.Requery
    End If
  ElseIf ClaveCatC = 8 Then
    If rTem.RecordCount = 0 Then
      conn.Execute "INSERT INTO unidades(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",1,'PIEZA')"
      conn.Execute "INSERT INTO unidades(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",2,'PAR')"
      conn.Execute "INSERT INTO unidades(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",3,'KILO')"
      rTem.Requery
    End If
  ElseIf ClaveCatC = 12 Then
    If rTem.RecordCount = 0 Then
      conn.Execute "INSERT INTO cargos(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",1,'PRESTAMOS')"
      conn.Execute "INSERT INTO cargos(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",2,'COMIDA')"
      conn.Execute "INSERT INTO cargos(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",3,'EXTRAS')"
      conn.Execute "INSERT INTO cargos(Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & ",4,'FALTAS')"
      rTem.Requery
    End If
  ElseIf ClaveCatC = 16 Then
    If rTem.RecordCount = 0 Then
      ActualizaMP
      rTem.Requery
    End If
  ElseIf ClaveCatC = 19 Then
    If rTem.RecordCount = 0 Then
      conn.Execute "INSERT INTO cttipopoliza VALUES(1,'VENTAS')"
      conn.Execute "INSERT INTO cttipopoliza VALUES(2,'COMPRAS')"
      rTem.Requery
    End If
  End If
  Grid1.Rows = 1
  With rTem
  Do Until .EOF
    If ClaveCatC = 2 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !Cuenta & L & Fmoneda(!Saldo) & L & Fmoneda(!NoCobrado)
    ElseIf ClaveCatC = 5 Or ClaveCatC = 6 Or ClaveCatC = 15 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !Cuenta
    ElseIf ClaveCatC = 7 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !Cambio
    ElseIf ClaveCatC = 9 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !Directo
    ElseIf ClaveCatC = 11 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !Status
    ElseIf ClaveCatC = 16 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !ClvCue
    ElseIf ClaveCatC = 17 Then
      Cad = L & !Clave & L & !Nombre & L & Fmoneda(!Valor) & L & !Repetir
    ElseIf ClaveCatC = 18 Then
      Cad = "" & L & !Clave & L & !Nombre & L & !RVOE1
    Else
      Cad = "" & L & !Clave & L & !Nombre
    End If
    If Not IsNull(!Nombre) Then Combo1.AddItem !Nombre
    Grid1.AddItem Cad
    .MoveNext
  Loop
  CambiaColor Grid1
  End With
End Sub
Private Sub Combo1_DblClick()
  Encontro = 0
  For a = 1 To Grid1.Rows - 1
    If Combo1 = Grid1.TextMatrix(a, 2) Then
      Grid1.Row = a: Grid1.Col = 0: Encontro = True: Combo1.Visible = False
      SendKeys "{LEFT}": Grid1.SetFocus
      Exit For
    End If
  Next
  If Encontro = 0 Then MsgBox "No existe " & Combo1, 48, rEmp!Nombre
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Combo1.ListIndex < 0 Then
      MsgBox "Seleccione el nombre de " & nTabla, 48, rEmp!Nombre: Combo1.SetFocus
    Else
      Encontro = 0
      For a = 1 To Grid1.Rows - 1
        If Combo1 = Grid1.TextMatrix(a, 2) Then
          Grid1.Row = a: Grid1.Col = 0: Encontro = True: Combo1.Visible = False
          SendKeys "{LEFT}": Grid1.SetFocus
          Exit For
        End If
      Next
      If Encontro = 0 Then MsgBox "No existe " & Combo1, 48, rEmp!Nombre
    End If
  ElseIf KeyAscii = 27 Then
    Combo1.Visible = False
  End If
End Sub


Private Sub Command1_Click(Index As Integer)
  If Index = 0 Then             'Agregar
    Frame1.Caption = "Agregar": Frame1.Visible = True: Command2(0).Tag = "A"
    sql = "SELECT MAX(Clave) as Mayor FROM " & nTabla
'    If ClaveCatC <> 7 And ClaveCatC <> 16 Then SQL = SQL & " WHERE Empresa=" & rEmp!Clave
    Set rTem = conn.Execute(sql)
    Text1(0) = Valor(rTem!Mayor) + 1: Text1(1) = "": Text1(2) = "": Text1(3) = "": Text1(4) = "": Text1(0).Enabled = True
    Text1(1).SetFocus
  ElseIf Index = 1 Then         'Editar
    Clv = Grid1.TextMatrix(Grid1.Row, 1)
    If Len(Clv) = 0 Then Exit Sub
    Grid1.Enabled = False: Frame1.Caption = "Editar"
    Text1(0) = Clv: Text1(1) = Grid1.TextMatrix(Grid1.Row, 2): Command2(0).Tag = "E"
    Text1(0).Enabled = False: Frame1.Visible = True
    Select Case ClaveCatC
      Case 2
        Text1(2) = Grid1.TextMatrix(Grid1.Row, 3)
        Text1(3) = Grid1.TextMatrix(Grid1.Row, 4): Text1(3).Tag = Text1(3)
        Text1(4) = Grid1.TextMatrix(Grid1.Row, 5)
      Case 17
        Text1(2) = Grid1.TextMatrix(Grid1.Row, 3)
        Text1(3) = Grid1.TextMatrix(Grid1.Row, 4): Text1(3).Tag = Text1(3)
      Case 5, 6, 7, 9, 11, 16, 18
        Text1(2) = Grid1.TextMatrix(Grid1.Row, 3)
    End Select
    Text1(1).SetFocus
  ElseIf Index = 2 Then         'Borrar
    If Grid1.Row < 1 Then Exit Sub
    Clv = Grid1.TextMatrix(Grid1.Row, 1)
    If Len(Clv) = 0 Then Exit Sub
    r = MsgBox("Desea borrar el concepto " & Grid1.TextMatrix(Grid1.Row, 2), 36, rEmp!Nombre)
    If r = 6 Then
      If ClaveCatC = 24 Or ClaveCatC = 27 Then
        conn.Execute "DELETE FROM " & nTabla & " WHERE Clave='" & Clv & "'"
      Else
        conn.Execute "DELETE FROM " & nTabla & " WHERE Clave=" & Clv
      End If
      If Grid1.Rows = 2 Then Grid1.Rows = 1 Else Grid1.RemoveItem Grid1.Row
    End If
  ElseIf Index = 3 Then
    If ClaveCatC = 2 Then
      If Val(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then Exit Sub
      Frame2.Visible = True
      Text2(0).SetFocus
    Else
      Imprimir
    End If
  ElseIf Index = 4 Then
    Combo1.Visible = True: Combo1.SetFocus
  Else
    Unload Me
  End If
End Sub
Private Sub Command2_Click(Index As Integer)
  If Index = 0 Then
    If Len(Text1(0)) = 0 Then
      MsgBox "La clave no puede estar vacia", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
    End If
    If Len(Trim(Text1(1))) = 0 Then
      MsgBox "La descripción es obligatoria.", 48, rEmp!Nombre: Text1(1).SetFocus: Exit Sub
    End If
    If Command2(0).Tag = "A" Then
'      If ClaveCatC = 7 Or ClaveCatC = 16 Then SQL2 = " WHERE " Else SQL2 = " WHERE Empresa=" & rEmp!Clave & " And "
      SQL2 = " WHERE "
      sql = "SELECT Clave FROM " & nTabla & SQL2 & "Clave=" & Text1(0)
      Set rTem = conn.Execute(sql)
      If rTem.RecordCount > 0 Then
        MsgBox "La clave ya existe.", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
      End If
      Set rTem = conn.Execute("SELECT Clave FROM " & nTabla & SQL2 & "Nombre='" & Text1(1) & "'")
      If rTem.RecordCount > 0 Then
        MsgBox "El nombre ya existe.", 48, rEmp!Nombre: Text1(1).SetFocus: Exit Sub
      End If
      If ClaveCatC = 2 Then  'Conceptos CA
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre,Cuenta,Saldo,NoCobrado) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) _
        & "','" & Text1(2) & "'," & Valor(Text1(3)) & "," & Valor(Text1(4)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2) & L & Text1(3) & L & Text1(4)
      ElseIf ClaveCatC = 5 Or ClaveCatC = 6 Or ClaveCatC = 15 Then  'ComprasCla, concepto NC
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre,Cuenta) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) _
        & "'," & Val(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2) & L & Text1(3)
      ElseIf ClaveCatC = 7 Then
        sql = "INSERT INTO " & nTabla & " (Clave,Nombre,Cambio) VALUES(" & Text1(0) & ",'" & Text1(1) & "'," & Val(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2)
      ElseIf ClaveCatC = 9 Then  'Deptos
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre,Directo) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) _
        & "'," & Val(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2) & L & Text1(3)
      ElseIf ClaveCatC = 11 Then  'Vendeores
        Text1(2) = -1
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre,Status) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) _
        & "'," & Valor(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2)
      ElseIf ClaveCatC = 16 Then    'metodo de pago
        sql = "INSERT INTO " & nTabla & " (Clave,Nombre,ClvCue) VALUES(" & Text1(0) & ",'" & Text1(1) & "'," & Valor(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2)
      ElseIf ClaveCatC = 17 Then      'Cargos de colegiatura
        If Val(Text1(3)) = 0 Then Text1(3) = 0 Else Text1(3) = -1
        sql = "INSERT INTO " & nTabla & " (Clave,Nombre,Valor,Repetir) VALUES(" & Text1(0) & ",'" & Text1(1) & "'," & Valor(Text1(2)) & "," & Valor(Text1(3)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2) & L & Text1(3)
      ElseIf ClaveCatC = 18 Then      'Cargos de colegiatura
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre,RVOE1) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) & "'," & Valor(Text1(2)) & ")"
        Grid1.AddItem L & Text1(0) & L & Text1(1) & L & Text1(2)
      ElseIf ClaveCatC = 19 Then
        sql = "INSERT INTO " & nTabla & " (Clave,Nombre) VALUES(" & Text1(0) & ",'" & Text1(1) & "')"
        Grid1.AddItem L & Text1(0) & L & Text1(1)
      Else
        sql = "INSERT INTO " & nTabla & " (Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & "," & Text1(0) & ",'" & Text1(1) & "')"
        Grid1.AddItem L & Text1(0) & L & Text1(1)
      End If
      conn.Execute sql
      Grid1.SetFocus
      Grid1.Col = 0: Grid1.Row = Grid1.Rows - 1
      SendKeys "{LEFT}"
    Else
      If ClaveCatC = 2 Then    'Conceptos cargos agonos
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Cuenta='" & Text1(2) & "',Saldo=" & Valor(Text1(3)) _
        & ",NoCobrado=" & Val(Text1(4)) & " WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
        Sa1 = Valor(Text1(3).Tag): Sa2 = Valor(Text1(3))
        If Sa1 <> Sa2 Then
          Dife = Abs(Sa1 - Sa2)
          If Sa1 > Sa2 Then tSi = "-" Else tSi = "+"
'          CAMovi Text1(0), Dife, tSi, Date, "Ajuste de saldo", -1
        End If
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2): Grid1.TextMatrix(Grid1.Row, 4) = Text1(3): Grid1.TextMatrix(Grid1.Row, 5) = Text1(4)
      ElseIf ClaveCatC = 5 Or ClaveCatC = 6 Or ClaveCatC = 15 Then  'ComprasCla, ConceptoNC, Vales
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Cuenta='" & Text1(2) & "' WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 7 Then
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Cambio=" & Valor(Text1(2)) & " WHERE Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 9 Then
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Directo=" & Valor(Text1(2)) & " WHERE Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 11 Then    'Vendedores
        If Val(Text1(2)) = 0 Then Text1(2) = 0 Else Text1(2) = -1
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Status=" & Text1(2) & " WHERE Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 16 Then      'Forma de pago
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',ClvCue=" & Valor(Text1(2)) & " WHERE Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 17 Then      'Cargos colegiatura
        If Val(Text1(3)) = 0 Then Text1(3) = 0 Else Text1(3) = -1
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',Valor=" & Valor(Text1(2)) & ",Repetir=" & Valor(Text1(3)) & " WHERE Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2): Grid1.TextMatrix(Grid1.Row, 4) = Text1(3)
      ElseIf ClaveCatC = 18 Then      'Carreras universidad
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "',RVOE1='" & Text1(2) & "' WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
        Grid1.TextMatrix(Grid1.Row, 3) = Text1(2)
      ElseIf ClaveCatC = 19 Then      'Tipos de poliza
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "' WHERE Clave=" & Text1(0)
      Else
        sql = "UPDATE " & nTabla & " SET Nombre='" & Text1(1) & "' WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
      End If
      conn.Execute sql
      Grid1.TextMatrix(Grid1.Row, 2) = Text1(1)
    End If
  End If
  Frame1.Visible = False
  Grid1.Enabled = True
End Sub

Private Sub Command3_Click(Index As Integer)
  If Index = 0 Then
    If Not IsDate(Text2(0)) Then
      MsgBox "Formato de la fecha incorrecto.", 48, rEmp!Nombre: Text2(0).SetFocus: Exit Sub
    End If
    If Not IsDate(Text2(1)) Then
      MsgBox "Formato de la fecha incorrecto.", 48, rEmp!Nombre: Text2(1).SetFocus: Exit Sub
    End If
    sql = "{camovi.Fecha}>=Date(" & Format(Text2(0), "yyyy,mm,dd") & ") And {camovi.Fecha}<=Date(" & Format(Text2(1), _
    "yyyy,mm,dd)") & " And {camovi.Cuenta}=" & Grid1.TextMatrix(Grid1.Row, 1)
    Titu = "Movimientos bancarios del: " & Text2(0) & " al " & Text2(1)
    Menu.Report1Im "EdoCuenta.rpt", rEmp!Nombre, Titu, sql, "", ""
  End If
  Frame2.Visible = False
End Sub

Private Sub Form_Load()
  CentrarFrm Me
  Frame1.Height = 2895: Frame1.Left = 180: Frame1.Top = 1440: Frame1.Width = 8535
  Frame2.Height = 3195: Frame2.Left = 2040: Frame2.Top = 1020: Frame2.Width = 4695
  Combo1.Height = 5835: Combo1.Left = 1020: Combo1.Top = 780: Combo1.Width = 6975
  Text2(0) = date: Text2(1) = date
  Formato
  BotonPic Me
  e = Chr(13)
  If ClaveCatC = 1 Then
    nTabla = "proveedorcd": Me.Caption = "Descuentos a proveedores."
  ElseIf ClaveCatC = 2 Then
    nTabla = "conceptosca": Me.Caption = "Cuentas bancarias."
  ElseIf ClaveCatC = 3 Then
    nTabla = "familia": Me.Caption = "Catalogo de familias."
  ElseIf ClaveCatC = 4 Then
    nTabla = "familia2": Me.Caption = "Catalogo de sub familia."
  ElseIf ClaveCatC = 5 Then
    nTabla = "comprascla": Me.Caption = "Compras clasificación."
  ElseIf ClaveCatC = 6 Then
    nTabla = "conceptonc": Me.Caption = "Conceptos notas de credito."
  ElseIf ClaveCatC = 7 Then
    nTabla = "moneda": Me.Caption = "Tipos de moneda."
  ElseIf ClaveCatC = 8 Then
    nTabla = "unidades": Me.Caption = "Unidades de medida."
  ElseIf ClaveCatC = 9 Then
    nTabla = "depto": Me.Caption = "Departamentos."
    Label2 = "0 = M.O.D." & e & "1 = G. Fabricación" & e & "2 = G. Admon" & e & "3 = G. Venta"
    Label2.Visible = True
  ElseIf ClaveCatC = 10 Then
    nTabla = "puesto": Me.Caption = "Puestos de trabajo."
  ElseIf ClaveCatC = 11 Then
    nTabla = "vendedor": Me.Caption = "Agentes de venta."
  ElseIf ClaveCatC = 12 Then
    nTabla = "cargos": Me.Caption = "Cargos nómina."
  ElseIf ClaveCatC = 13 Then
    nTabla = "pais": Me.Caption = "País."
  ElseIf ClaveCatC = 14 Then
    nTabla = "conceptomi": Me.Caption = "Conceptos Movimientos al Inventario."
    Label2 = "1 a 50 Entradas" & e & ">50 Salidas": Label2.Visible = True
  ElseIf ClaveCatC = 15 Then
    nTabla = "conceptoval": Me.Caption = "Vales INGRESOS/EGRESOS."
  ElseIf ClaveCatC = 16 Then
    nTabla = "metodopago": Me.Caption = "Metodos de pago."
  ElseIf ClaveCatC = 17 Then
    nTabla = "cargos": Me.Caption = "Conceptos cargos."
  ElseIf ClaveCatC = 18 Then
    nTabla = "carreras": Me.Caption = "Nivel educativo."
  ElseIf ClaveCatC = 19 Then
    nTabla = "cttipopoliza": Me.Caption = "Tipos de póliza."
  ElseIf ClaveCatC = 20 Then
    nTabla = "ciudad": Me.Caption = "Ciudades."
  ElseIf ClaveCatC = 21 Then
    nTabla = "ubicacion": Me.Caption = "Ubicaciones."
  End If
  Mostrar
End Sub
Sub Imprimir()
  If Dir("c:\tempo", vbDirectory) = "" Then MkDir "c:\Tempo"
  NomPdf = "c:\Tempo\Imprimir.pdf"
  If Dir(NomPdf) <> "" Then Kill NomPdf
  Set pdf = New PdfDoc
  pdf.AddPage 1
'Cabecera
'  pdf.HEADER True
  pdf.SetY 5
  pdf.SetFont "Arial", "B", 14
  pdf.Cell 190, 8, Trim(rEmp!Nombre), 0, 1, 1, 0, ""
  pdf.SetFont "Arial", "B", 12
  Ti = Me.Caption
'  Ti = Ti & " (" & rUsu!Usuario & ")"
  pdf.Cell 190, 8, Ti, 0, 1, 1, 0, ""
  y = pdf.GetY:  pdf.SetXY 10, y - 10
  pdf.SetFont "Arial", "", 10
  pdf.Cell 20, 5, "Fecha:", 15, 0, 0, 0, ""
  pdf.Cell 20, 5, Format(date, "dd/mm/yy"), 15, 1, 0, 0, ""
  pdf.Cell 20, 5, "Hora:", 15, 0, 0, 0, ""
  pdf.Cell 20, 5, Format(Time, "hh:mm"), 15, 1, 0, 0, ""
'  pdf.HEADER False
'Pie de pagina
  pdf.AliasNbPages "{nb}"
  pdf.FOOTER True
  pdf.SetY -10
  pdf.SetFont "Arial", "B", 10
  pdf.Cell 180, 5, "Pag: {pg}/{nb}", 0, 0, 2, 0, ""
  pdf.FOOTER False
'Detalle
  pdf.SetFont "Arial", "B", 9
  pdf.SetY 25
  pdf.SetTextColor 255, 255, 255
  For b = 0 To Grid1.Cols - 1
    pdf.Cell CInt(Grid1.ColWidth(b) / 45), 5, Grid1.TextMatrix(0, b), 15, ln, 0, 1, ""
  Next
  pdf.ln 5
  pdf.SetTextColor 0, 0, 0
  For a = 1 To Grid1.Rows - 1
    For b = 0 To Grid1.Cols - 1
      If b < Grid1.Cols - 1 Then ln = 0 Else ln = 1
      If b < 3 Then al = 0 Else al = 2
      pdf.Cell CInt(Grid1.ColWidth(b) / 45), 5, Grid1.TextMatrix(a, b), 15, ln, al, 0, ""
    Next
  Next
  If Grid1.FixedRows = 2 Then
    For b = 0 To Grid1.Cols - 1
        If b < 3 Then al = 0 Else al = 2
      pdf.Cell CInt(Grid1.ColWidth(b) / 45), 5, Grid1.TextMatrix(1, b), 0, 0, al, 0, ""
    Next
  End If
  pdf.SaveAsFile NomPdf
  MousePointer = 0
  MsgBox "Pulse <ENTER> para ver el reporte.", 48, rEmp!Nombre
  Call ShellExecute(Me.hwnd, "open", NomPdf, "", "", 3)
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    Grid1.Height = Me.Height - 660
  End If
End Sub


Private Sub Grid1_DblClick()
  If ClaveCatC = 5 And Grid1.MouseCol = 0 Then
    r = InputBox("Ingrese el número de trabajador.", rEmp!Nombre)
    If Val(r) <= 0 Then Exit Sub
    Set rTem = conn.Execute("SELECT Clave,Nombre FROM personal WHERE Clave=" & Val(r))
    If rTem.RecordCount = 0 Then
      MsgBox "No existe la clave " & r, 48, rEmp!Nombre: Exit Sub
    End If
    MousePointer = 11
    Report1.ReportFileName = DirRep & "Depto.Rpt"
    Report1.Connect = CRConn
    Report1.Formulas(0) = "Clave=" & Val(r)
    Report1.Formulas(1) = "Nombre='" & rTem!Nombre & "'"
    Report1.SelectionFormula = "{depto.clave}=" & Grid1.TextMatrix(Grid1.Row, 1)
    Report1.Destination = 0
    Report1.Action = 0
    MousePointer = 0
  End If
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Grid1.MouseRow > 0 Then Exit Sub
  Grid1.SelectionMode = flexSelectionFree
  Grid1.Col = Grid1.MouseCol
  Grid1.Sort = 1
  Grid1.SelectionMode = flexSelectionByRow
  CambiaColor Grid1
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

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 2 Then
    If KeyCode = vbKeyF2 Then
      If ClaveCatC = 5 Then
        If OpcBuscar <> "CUENTAS" Then If rFormulario(BuscarFrm) Then Unload BuscarFrm
        OpcBuscar = "CUENTASD"
        BuscarFrm.Show 1
        If InStr(ClipText, ";") > 0 Then
          cam = Split(ClipText, ";")
          Text1(Index) = cam(0)
        End If
      End If
    End If
  End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  TextL Text1(Index)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
  Text2(Index).SelStart = 0
  Text2(Index).SelLength = Len(Text2(Index))
  Text2(Index).backcolor = QBColor(11)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
  Text2(Index).backcolor = QBColor(15)
  Text2(Index) = Fecha6(Text2(Index), 0)
End Sub
