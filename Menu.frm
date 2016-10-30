VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm Menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu"
   ClientHeight    =   5184
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   12060
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Report1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu CatalogosMnu 
      Caption         =   "Catalogos"
      Begin VB.Menu DiputadosMnu 
         Caption         =   "Diputados"
         Visible         =   0   'False
      End
      Begin VB.Menu ArticulosMnu 
         Caption         =   "Articulos"
      End
      Begin VB.Menu ClasificacionMnu 
         Caption         =   "Clasificación"
         Visible         =   0   'False
      End
      Begin VB.Menu UbicacionMnu 
         Caption         =   "Ubicacion"
      End
      Begin VB.Menu Inventario 
         Caption         =   "Inventario"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Report1Im(nArch, nEmpr, nTitu, nSele, nOrde, nFor3)
  If Dir(DirRep & nArch) = "" Then
    r = MsgBox("No existe el archivo " & DirRep & nArch, 48, Empresa.Empresa)
    Exit Sub
  End If
  MousePointer = 11
  For a = 0 To 10
    Report1.Formulas(a) = ""
    Report1.SortFields(a) = ""
  Next
  Report1.SelectionFormula = ""
  Report1.ReportFileName = DirRep & nArch
  Report1.Connect = CRConn
  If Len(nEmpr) = 0 Then Report1.Formulas(0) = "" Else Report1.Formulas(0) = "Empresa='" & Trim(nEmpr) & "'"
  If Len(nTitu) = 0 Then
    Report1.Formulas(1) = "": Report1.Formulas(2) = ""
  Else
    Report1.Formulas(1) = "Titulo='" & nTitu & "'"
    Report1.Formulas(2) = "Archivo='" & nArch & "'"
  End If
  Report1.Formulas(3) = nFor3
  Report1.SelectionFormula = nSele
  Report1.SortFields(0) = nOrde
  Report1.Destination = 0
  Report1.Action = 0
  MousePointer = 0
End Sub

Private Sub ArticulosMnu_Click()
  InventarFrm.Show
End Sub

Private Sub ClasificacionMnu_Click()
  ClaveCatC = 17
  CatalogosFrm.Show 1
End Sub

Private Sub DiputadosMnu_Click()
  DiputadosFrm.Show
End Sub


Private Sub Inventario_Click()
  InvFisicoFrm.Show
End Sub

Private Sub MDIForm_Load()
  Conexion
  L = Chr(9)
  Me.Picture = LoadPicture("logo.wmf")
  Set rEmp = conn.Execute("SELECT * FROM empresa")
End Sub

Private Sub UbicacionMnu_Click()
  ClaveCatC = 21
  CatalogosFrm.Show 1
End Sub


