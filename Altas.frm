VERSION 5.00
Begin VB.Form AltasFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alta catalogos."
   ClientHeight    =   1944
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   6324
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1944
   ScaleWidth      =   6324
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "palomita.jpg"
      Top             =   900
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Clave"
      DataMember      =   "N"
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
      Left            =   60
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   6195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   6195
   End
End
Attribute VB_Name = "AltasFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClvC As Integer
Dim tTabla As String
Dim rTem As ADODB.Recordset
Private Sub Command1_Click()
  Text1(0) = Trim(Text1(0))
  If Len(Text1(0)) = 0 Then
    MsgBox "No puede ir vacio el nombre de " & tTabla & ".", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
  End If
  If ClaveCat = 15 Or ClaveCat = 16 Then SQL2 = "" Else SQL2 = " And Empresa=" & rEmp!Clave
  sql = "SELECT Nombre FROM " & tTabla & " WHERE Nombre='" & Text1(0) & "'" & SQL2
  Set rTem = conn.Execute(sql)
  If rTem.RecordCount > 0 Then
    MsgBox "El nombre de " & tTabla & " ya existe.", 48, rEmp!Nombre: Text1(0).SetFocus: Exit Sub
  End If
  If ClaveCat = 15 Or ClaveCat = 16 Then SQL2 = "" Else SQL2 = " WHERE Empresa=" & rEmp!Clave
  Set rTem = conn.Execute("SELECT Max(Clave) as Clave FROM " & tTabla & SQL2)
  Clv = Valor(rTem!Clave) + 1
  If ClaveCat = 9 Then        'Vendedor
    sql = "INSERT INTO " & tTabla & " VALUES(" & rEmp!Clave & "," & Clv & ",'" & Text1(0) & "',-1)"
  ElseIf ClaveCat = 15 Or ClaveCat = 16 Then
    sql = "INSERT INTO " & tTabla & " VALUES(" & Clv & ",'" & Text1(0) & "')"
  Else
    sql = "INSERT INTO " & tTabla & " (Empresa,Clave,Nombre) VALUES(" & rEmp!Clave & "," & Clv & ",'" & Text1(0) & "')"
  End If
  conn.Execute sql
  ClipText = Clv & ";" & Text1(0)
  Unload Me
End Sub
Private Sub Form_Load()
  CentrarFrm Me
  Me.Top = 3500
  x = &H8000000D
  BotonPic Me
  ClipText = "": Text1(0) = ""
  If ClaveCat = 1 Then tTabla = "ciudad"
  If ClaveCat = 2 Then tTabla = "estado"
  If ClaveCat = 3 Then tTabla = "pais"
  If ClaveCat = 4 Then tTabla = "familia"
  If ClaveCat = 5 Then tTabla = "apoyos"
  If ClaveCat = 6 Then tTabla = "nccompras"
  If ClaveCat = 7 Then tTabla = "familia2"
  If ClaveCat = 8 Then tTabla = "unidades"
  If ClaveCat = 9 Then tTabla = "vendedor"
  If ClaveCat = 10 Then tTabla = "conceptoval"
  If ClaveCat = 11 Then tTabla = "horario"
  If ClaveCat = 12 Then tTabla = "depto"
  If ClaveCat = 13 Then tTabla = "puesto"
  If ClaveCat = 14 Then tTabla = "carreras"
  If ClaveCat = 15 Then tTabla = "grado"
  If ClaveCat = 16 Then tTabla = "puesto"
  If ClaveCat = 17 Then tTabla = "comprascla"
  Label1(0) = "Agregar " & tTabla
End Sub

Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).SelStart = 0
  Text1(Index).SelLength = Len(Text1(Index))
  Text1(Index).BackColor = QBColor(11)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
  ElseIf KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  Text1(Index).BackColor = QBColor(15)
End Sub
