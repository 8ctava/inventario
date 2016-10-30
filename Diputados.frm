VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form DiputadosFrm 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diputados"
   ClientHeight    =   8052
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   16032
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8052
   ScaleWidth      =   16032
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   672
      Index           =   8
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7140
      Width           =   1452
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   15000
      Top             =   7140
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir."
      Height          =   3192
      Left            =   11820
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   3672
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   2040
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   435
         Index           =   0
         Left            =   480
         TabIndex        =   33
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credencial todos"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   32
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credencial individual"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   31
         Top             =   900
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   600
      ScaleHeight     =   564
      ScaleWidth      =   8964
      TabIndex        =   23
      Tag             =   "Dinero.bmp"
      Top             =   6420
      Width           =   9015
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Regresa al primer registro."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Regresa al registro anterior."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   2
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Avanza al siguiente registro."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Datos 
         Height          =   555
         Index           =   3
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   24
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
         TabIndex        =   28
         Top             =   180
         Width           =   4680
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox contenido 
      BorderStyle     =   0  'None
      Height          =   7032
      Left            =   -12
      ScaleHeight     =   7032
      ScaleWidth      =   15672
      TabIndex        =   35
      Top             =   0
      Width           =   15672
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "Foto"
         Height          =   2652
         Left            =   12540
         ScaleHeight     =   2652
         ScaleWidth      =   2412
         TabIndex        =   55
         Top             =   120
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
         Top             =   900
         Width           =   1152
      End
      Begin VB.TextBox Text1 
         DataField       =   "APaterno"
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
         Top             =   900
         Width           =   1932
      End
      Begin VB.TextBox Text1 
         DataField       =   "AMaterno"
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
         Index           =   2
         Left            =   3600
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   900
         Width           =   2052
      End
      Begin VB.TextBox Text1 
         DataField       =   "ANombre"
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
         Index           =   3
         Left            =   5820
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   900
         Width           =   3732
      End
      Begin VB.TextBox Text1 
         DataField       =   "Direccion"
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
         Height          =   348
         Index           =   4
         Left            =   180
         MaxLength       =   60
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1740
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         DataField       =   "Colonia"
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
         Index           =   5
         Left            =   5760
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1740
         Width           =   3672
      End
      Begin VB.TextBox Text1 
         DataField       =   "Telefono"
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Index           =   6
         Left            =   9660
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1740
         Width           =   2712
      End
      Begin VB.TextBox Text1 
         DataField       =   "Mail"
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
         Index           =   8
         Left            =   3600
         MaxLength       =   60
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2580
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         DataField       =   "FechaNa"
         DataMember      =   "F"
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
         Index           =   9
         Left            =   10440
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2520
         Width           =   1152
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Partido"
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
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3420
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         DataField       =   "Obser"
         DataMember      =   "T"
         Height          =   1908
         Index           =   16
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   4260
         Width           =   10872
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Baja"
         DataField       =   "Status"
         DataMember      =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   36
         Top             =   3180
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         DataField       =   "RFC"
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
         Index           =   13
         Left            =   8580
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2580
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         DataField       =   "Celular"
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
         Index           =   7
         Left            =   180
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2580
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         DataField       =   "Cuenta"
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   11
         Left            =   180
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   3420
         Width           =   1932
      End
      Begin VB.TextBox Text1 
         DataField       =   "Posicion"
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
         Index           =   17
         Left            =   9720
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   900
         Width           =   2652
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         Height          =   252
         Index           =   0
         Left            =   180
         TabIndex        =   53
         Top             =   660
         Width           =   1152
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A. Paterno"
         Height          =   252
         Index           =   1
         Left            =   1500
         TabIndex        =   52
         Top             =   660
         Width           =   1932
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A. Materno"
         Height          =   252
         Index           =   2
         Left            =   3600
         TabIndex        =   51
         Top             =   660
         Width           =   2052
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre(s)"
         Height          =   252
         Index           =   3
         Left            =   5820
         TabIndex        =   50
         Top             =   660
         Width           =   3732
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion:"
         Height          =   252
         Index           =   4
         Left            =   180
         TabIndex        =   49
         Top             =   1500
         Width           =   5412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Colonia:"
         Height          =   252
         Index           =   5
         Left            =   5760
         TabIndex        =   48
         Top             =   1500
         Width           =   3672
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono"
         Height          =   252
         Index           =   6
         Left            =   9660
         TabIndex        =   47
         Top             =   1500
         Width           =   2712
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mail:"
         Height          =   252
         Index           =   7
         Left            =   3600
         TabIndex        =   46
         Top             =   2340
         Width           =   4812
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F. Alta:"
         Height          =   252
         Index           =   8
         Left            =   10440
         TabIndex        =   45
         Top             =   2280
         Width           =   1152
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partido"
         Height          =   252
         Index           =   10
         Left            =   2280
         TabIndex        =   44
         Top             =   3180
         Width           =   2712
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones:"
         Height          =   312
         Index           =   13
         Left            =   60
         TabIndex        =   43
         Top             =   3960
         Width           =   1452
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         DataField       =   "Nombre"
         DataMember      =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   432
         Index           =   0
         Left            =   180
         TabIndex        =   42
         Top             =   60
         Width           =   9252
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RFC"
         Height          =   252
         Index           =   15
         Left            =   8580
         TabIndex        =   41
         Top             =   2340
         Width           =   1632
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Celular"
         Height          =   252
         Index           =   16
         Left            =   180
         TabIndex        =   40
         Top             =   2340
         Width           =   3252
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         Height          =   252
         Index           =   19
         Left            =   180
         TabIndex        =   39
         Top             =   3180
         Width           =   1932
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ubicaciòn oficina"
         Height          =   252
         Index           =   23
         Left            =   9720
         TabIndex        =   38
         Top             =   660
         Width           =   2652
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fotografía"
      Height          =   672
      Index           =   7
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "Camara.jpg"
      Top             =   7140
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   615
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7140
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   615
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7140
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar &Nombre"
      Height          =   672
      Index           =   5
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7140
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar &Clave"
      Height          =   672
      Index           =   4
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7140
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   672
      Index           =   3
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7140
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Borrar"
      Height          =   672
      Index           =   2
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7140
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   672
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7140
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   672
      Index           =   0
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7140
      Width           =   1452
   End
End
Attribute VB_Name = "DiputadosFrm"
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
  nclv = CPosicion(NDoc, TipoM, "diputados", "Clave", 0)
  LabelDatos = "Total diputados: " & NR
  Set rAlu = conn.Execute("SELECT * FROM diputados WHERE Empresa=" & rEmp!Clave & " And Clave=" & nclv)
  If rAlu.RecordCount = 0 Then
    MsgBox "No existe la clave " & nclv, 48, rEmp!Nombre: Exit Sub
  Else
    MostrarTexN Me, rAlu
    If rAlu!Status = 0 Then Me.backcolor = QBColor(12) Else Me.backcolor = rEmp!Fondo
  End If
End Sub
Private Sub Check1_Click()
  If Check1 Then Check1.Caption = "Activo" Else Check1.Caption = "Baja"
  Check1.ToolTipText = "1"
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
    sql = "SELECT Max(Clave) as NMa FROM diputados WHERE Empresa=" & rEmp!Clave
    Set rTem = conn.Execute(sql)
    Text1(0) = Valor(rTem!NMa) + 1
    Command2(0).Tag = "A":  Check1.ToolTipText = "1"
    Picture2.Picture = LoadPicture()
    Check1 = 1
    Text1(9) = date
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
    conn.Execute "DELETE FROM diputados WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0)
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
    Set rTem = conn.Execute("SELECT Clave FROM diputados WHERE Empresa=" & rEmp!Clave & " And Clave=" & Val(r))
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
    rTem.Open "SELECT * FROM diputados WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0), conn, adOpenDynamic, adLockOptimistic
    rTem!Foto = ImaD
    rTem.Update
    If Err = -2147217864 Then Err = 0
  ElseIf Index = 8 Then
    Unload Me
  End If
End Sub
Sub Busca()
  If OpcBuscar <> "DIPUTADOS" Then If rFormulario(BuscarFrm) Then Unload BuscarFrm
  OpcBuscar = "DIPUTADOS"
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
    tNom = Trim(Text1(1))
    If Len(Text1(2)) > 0 Then tNom = tNom & " " & Trim(Text1(2))
    tNom = tNom & " " & Trim(Text1(3))
    Label2(0) = tNom: Label2(0).ToolTipText = "1"
    If Command2(0).Tag = "A" Then
      Set rTem = conn.Execute("SELECT Clave FROM diputados WHERE Empresa=" & rEmp!Clave & " And Clave=" & Text1(0))
      If rTem.RecordCount > 0 Then
        MsgBox "El número de clave ya existe.", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
      End If
      Set rTem = conn.Execute("SELECT clave FROM diputados WHERE Empresa=" & rEmp!Clave & " And Nombre='" & tNom & "'")
      If rTem.RecordCount > 0 Then
        MsgBox "El nombre del alumno ya existe.", 48, rEmp!Nombre: Text1(1).SetFocus: Exit Sub
      End If
      sql = SqlCad(Me, "diputados", "A", "")
      conn.Execute sql
      NR = NR + 1: If Val(Text1(0)) > Mayor Then Mayor = Val(Text1(0))
      Posicion Text1(0), ""
    Else
      sql = SqlCad(Me, "diputados", "E", "Empresa=" & rEmp!Clave & " And Clave=" & Text1(0))
      If Len(sql) > 0 Then conn.Execute sql
    End If
    Activa Me, False
  ElseIf Index = 1 Then
    If NR = 0 Then Limpiar Me Else Posicion rAlu!Clave, ""
    Activa Me, False
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
  Me.Height = 8424: Me.Width = 15516
  BotonPic Me
  Combo1(0).AddItem "PRI"
  Combo1(0).AddItem "PAN"
  Combo1(0).AddItem "PRD"
  Combo1(0).AddItem "PT"
  Combo1(0).AddItem "PV"
  
  TextLon Me, "diputados"   'Este pone los colores
  CentrarFrm Me
  Me.Top = 0
  Activa Me, False
  Set rTem = conn.Execute("SELECT count(*) AS TRe FROM diputados WHERE Empresa=" & rEmp!Clave)
  NR = Valor(rTem!tRe): Mayor = 0
  If NR > 0 Then
    Set rTem = conn.Execute("SELECT Max(Clave) as Clave FROM diputados WHERE Empresa=" & rEmp!Clave)
    Mayor = Valor(rTem!Clave)
  Else
    Set rAlu = conn.Execute("SELECT * FROM diputados WHERE Empresa=" & rEmp!Clave)
  End If
  Set rTem = conn.Execute("SELECT count(*) AS TRe FROM diputados WHERE Empresa=" & rEmp!Clave & " And Status<>0")
  NrA = Valor(rTem!tRe)
  Posicion -1, ""
End Sub
Private Sub Label2_Click(Index As Integer)
  Label2(Index).ToolTipText = "1"
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


