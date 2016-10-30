Attribute VB_Name = "Module1"
Enum TQRCodeEncoding
    ceALPHA
    ceBYTE
    ceNUMERIC
    ceKANJI
    ceAUTO
End Enum

Enum TQRCodeECLevel
    LEVEL_L
    LEVEL_M
    LEVEL_Q
    LEVEL_H
End Enum
Public Declare Sub FullQRCode Lib "CBBQR.dll" (ByVal autoConfigurate As Boolean, ByVal AutoFit As Boolean, ByVal backcolor As Long, ByVal barcolor As Long, ByVal Texto As String, ByVal correctionLevel As TQRCodeECLevel, ByVal encoding As TQRCodeEncoding, ByVal marginpixels As Integer, ByVal moduleWidth As Integer, ByVal Height As Integer, ByVal Width As Integer, ByVal Filename As String)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Para abrir archivos PDF lin anterior

Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4
Public Const CF_BITMAP = 2
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long


Global Directorio As String
Global CRConn As String
Global DirIcon As String
Global DirRep As String
Global OpcBuscar As String

Global IPMySQL As String
Global ClaveCat As Integer
Global ClipText As String
Global ClaveCatC As Integer
Global ClaveCatF As Integer
Global L
Global wB64 As String

Global conn As ADODB.Connection
Global rEmp As ADODB.Recordset
Global ctEmp As ADODB.Recordset

Global pdf As New PdfDoc

Global ImaD As Variant    'Esta variable almacena los datos binarios de la foto

'Variables para zip
Public Type ZIPUSERFUNCTIONS
DLLPrnt As Long
DLLPassword As Long
DLLComment As Long
DLLService As Long
End Type
Public Type ZPOPT
fSuffix As Long
fEncrypt As Long
fSystem As Long
fVolume As Long
fExtra As Long
fNoDirEntries As Long
fExcludeDate As Long
fIncludeDate As Long
fVerbose As Long
fQuiet As Long
fCRLF_LF As Long
fLF_CRLF As Long
fJunkDir As Long
fRecurse As Long
fGrow As Long
fForce As Long
fMove As Long
fDeleteEntries As Long
fUpdate As Long
fFreshen As Long
fJunkSFX As Long
fLatestTime As Long
fComment As Long
fOffsets As Long
fPrivilege As Long
fEncryption As Long
fRepair As Long
flevel As Byte
date As String
szRootDir As String
End Type
Public Type ZIPnames
    s(0 To 99) As String
End Type
Public Type CBChar
    ch(4096) As Byte
End Type
Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long
'Fin variables zip
Function Buscar(CadB, tTim)
  If InStr(tTim, CadB) = 0 Then
    Buscar = ""
  Else
    If Mid(CadB, 1, 1) = "<" Then
      CadF = "</" & Mid(CadB, 2)
      Cad = Mid(tTim, InStr(tTim, CadB) + Len(CadB))
      If InStr(Cad, CadF) = 0 Then
        Cad = Mid(Cad, 1, InStr(Cad, "/>") - 1)
      Else
        Cad = Mid(Cad, 1, InStr(Cad, CadF) - 1)
      End If
      If InStr(CadB, "<cfdi:Conceptos") > 0 Then
        If InStr(Cad, "/>") > 0 Then Concep = Split(Cad, "/>") Else Concep = Split(Cad, "</")
        resul = "": Sepa = ""
        For a = 0 To UBound(Concep)
          tCan = Buscar("cantidad", Concep(a))
          tDes = Buscar("descripcion", Concep(a))
          tPre = Buscar("valorUnitario", Concep(a))
          tImp = Buscar("importe", Concep(a))
          tUni = Buscar("unidad", Concep(a))
          If Val(tCan) > 0 Then resul = resul & Sepa & tCan & "|" & tDes & "|" & tPre & "|" & tImp & "|" & tUni
          Sepa = "||"
        Next
        Buscar = resul
      ElseIf InStr(CadB, "<cfdi:Receptor") > 0 Then
        tRFC = Buscar("rfc", Cad)
        tNom = Buscar("nombre", Cad)
        tCal = Buscar("calle", Cad)
        tNuE = Buscar("noExterior", Cad)
        tNuI = Buscar("noInterior", Cad)
        tCol = Buscar("colonia", Cad)
        tLoc = Buscar("localidad", Cad)
        tMun = Buscar("municipio", Cad)
        tEst = Buscar("estado", Cad)
        tPai = Buscar("pais", Cad)
        tCP = Buscar("codigoPostal", Cad)
        resul = tRFC & "|" & tNom & "|" & tCal & "|" & tNuE & "|" & tNuI & "|" & tCol & "|" & tLoc & "|" & tMun & "|" & tEst & "|" & tPai & "|" & tCP
        Buscar = resul
      ElseIf InStr(CadB, "<implocal:ImpuestosLocales") > 0 Then
        Concep = Split(Cad, "/>")
        resul = "": Sepa = ""
        For a = 0 To UBound(Concep)
          tImpue = Buscar("ImpLocTrasladado", Concep(a))
          tImpor = Buscar("Importe", Concep(a))
          tTasa = Buscar("TasadeTraslado", Concep(a))
          If Valor(tImpor) > 0 Then
            resul = resul & Sepa & tImpue & "|" & tImpor & "|" & tTasa
            Sepa = "||"
          End If
        Next
        Buscar = resul
'IVA Retenciones
      ElseIf InStr(CadB, "<cfdi:Retenciones") > 0 Then
        If InStr(Cad, "/>") = 0 Then
          Concep = Split(Cad, "</cfdi:Retencion")
        Else
          Concep = Split(Cad, "/>")
        End If
        resul = "": Sepa = ""
        For a = 0 To UBound(Concep)
          tImpue = Buscar("impuesto", Concep(a))
          tImpor = Buscar("importe", Concep(a))
          tTasa = Buscar("tasa", Concep(a))
          If tTasa = "" Then tTasa = " "
          If Valor(tImpor) > 0 Then
            resul = resul & Sepa & tImpue & "|" & tImpor & "|" & tTasa
            Sepa = "||"
          End If
        Next
        Buscar = resul
      ElseIf InStr(CadB, "<cfdi:Traslados") > 0 Then
        If InStr(Cad, "/>") = 0 Then
          Concep = Split(Cad, "</cfdi:Traslado")
        Else
          Concep = Split(Cad, "/>")
        End If
        resul = "": Sepa = ""
        For a = 0 To UBound(Concep)
          tImpue = Buscar("impuesto", Concep(a))
          tImpor = Buscar("importe", Concep(a))
          tTasa = Buscar("tasa", Concep(a))
          If Valor(tImpor) > 0 Then
            resul = resul & Sepa & tImpue & "|" & tImpor & "|" & tTasa
            Sepa = "||"
          End If
        Next
        Buscar = resul
      ElseIf InStr(CadB, "<cfdi:Emisor") > 0 Then
        tRFC = Buscar("rfc", Cad)
        tNom = Buscar("nombre", Cad)
        tCal = Buscar("calle", Cad)
        tNuE = Buscar("noExterior", Cad)
        tNuI = Buscar("noInterior", Cad)
        tCol = Buscar("colonia", Cad)
        tLoc = Buscar("localidad", Cad)
        tMun = Buscar("municipio", Cad)
        tEst = Buscar("estado", Cad)
        tPai = Buscar("pais", Cad)
        tCP = Buscar("codigoPostal", Cad)
        resul = tRFC & "|" & tNom & "|" & tCal & "|" & tNuE & "|" & tNuI & "|" & tCol & "|" & tLoc & "|" & tMun & "|" & tEst & "|" & tPai & "|" & tCP
        Buscar = resul
      End If
    Else
      Cade = Mid(tTim, InStr(tTim, CadB) + Len(CadB) + 2)
      cade2 = Mid(Cade, 1, InStr(Cade, Chr(34)) - 1)
      cade2 = BuscarValidaTex(cade2)
      Buscar = cade2
    End If
  End If
End Function
Function BuscarValidaTex(Cad)
  lO = Chr(195) & Chr(129)  'À
  Cad = Replace(Cad, lO, "A")
  lO = Chr(195) & Chr(137)  'É
  Cad = Replace(Cad, lO, "E")
  
  lO = Chr(195) & Chr(169)  'è
  Cad = Replace(Cad, lO, "e")
  
  lO = Chr(195) & Chr(172)  'í
  Cad = Replace(Cad, lO, "i")
  
  lO = Chr(195) & Chr(173)  'í
  Cad = Replace(Cad, lO, "i")
  
  lO = Chr(195) & Chr(147)  'Ó
  Cad = Replace(Cad, lO, "O")
  lO = Chr(195) & Chr(179)  'Ó
  Cad = Replace(Cad, lO, "o")
  lO = Chr(195) & Chr(178)  'Ó acento al reves
  Cad = Replace(Cad, lO, "o")
  
  lO = Chr(195) & Chr(154)  'Ú
  Cad = Replace(Cad, lO, "U")
  lO = Chr(195) & Chr(186)  'Ú
  Cad = Replace(Cad, lO, "u")
  
  lO = Chr(195) & Chr(177)  'ñ
  Cad = Replace(Cad, lO, "n")
  
  lO = Chr(195) & Chr(145)  'Ñ
  Cad = Replace(Cad, lO, "N")
  
  Cad = Replace(Cad, "&amp;", "&")
  BuscarValidaTex = Cad
'Busca caracteres especiales
  For a = 1 To Len(Cad)
    Le = Mid(Cad, a, 1)
    If Le <> " " Then
      If InStr(wB64, Le) = 0 And Le <> "&" Then
        MsgBox Asc(Le) & "  -  " & Le & "  -  " & Cad & "  " & a
      End If
    End If
  Next

End Function


Sub cPartidaPol(ClvC, tFec, Tipo, Cta, TipoM, Monto, Concep, Acumula, ClvDoc, Folio, UUID, Opc2)
'Opc2 1=Operacion normal    2=Cancelacion
  Dim rCue As ADODB.Recordset
  If InStr(Cta, "-") > 0 Then
    Set rCue = conn.Execute("SELECT ID FROM ctcuentas WHERE Cuenta2='" & Cta & "'")
    If rCue.EOF Then CtaI = 1 Else CtaI = rCue!Id
  Else
    CtaI = Cta
  End If
  Ejer = Year(tFec): Peri = Month(tFec): nPol = Day(tFec)
  SQL2 = " WHERE ClvCat=" & ClvC & " And Ejercicio=" & Ejer & " And Periodo=" & Peri & " And Numero=" & nPol & " And Tipo=" & Tipo
  
  If Opc2 = 1 Then tSim = "+" Else tSim = "-"
  If Acumula Then
    sql = "UPDATE ctpartidapol SET Monto=Monto" & tSim & Monto & SQL2 & " And ID=" & CtaI
    conn.Execute sql, NuReg
    If NuReg = 0 And Opc2 = 1 Then
      sql = "INSERT INTO ctpartidapol(ClvCat,Ejercicio,Periodo,Numero,Tipo,Fecha,ID,Concepto,TipoM,Monto,ClvDoc,Folio,UUID) VALUES(" _
      & ClvC & "," & Ejer & "," & Peri & "," & nPol & "," & Tipo & "," & Fecha6(tFec, 3) & "," & CtaI & ",'" & Concep _
      & "','" & TipoM & "'," & Monto & "," & ClvDoc & "," & Folio & ",'" & UUID & "')"
      conn.Execute sql
    ElseIf Opc2 = 2 Then
      Set rTem = conn.Execute("SELECT Monto FROM ctpartidapol" & SQL2 & " AND ID=" & CtaI)
      If rTem!Monto = 0 Then conn.Execute "DELETE FROM ctpartidapol" & SQL2 & " AND ID=" & CtaI
    End If
  Else
    If Opc2 = 1 Then
      sql = "INSERT INTO ctpartidapol(ClvCat,Ejercicio,Periodo,Numero,Tipo,Fecha,ID,Concepto,TipoM,Monto,ClvDoc,Folio,UUID) VALUES(" & ClvC _
      & "," & Ejer & "," & Peri & "," & nPol & "," & Tipo & "," & Fecha6(tFec, 3) & "," & CtaI & ",'" & Concep & "','" & TipoM & "'," _
      & Monto & "," & ClvDoc & "," & Folio & ",'" & UUID & "')"
      conn.Execute sql
    Else
      SQL3 = SQL2 & " And ID=" & CtaI & " And ClvDoc=" & ClvDoc & " And Folio=" & Folio
      sql = "UPDATE ctpartidapol SET Monto=Monto" & tSim & Monto & SQL3
      conn.Execute sql
      Set rTem = conn.Execute("SELECT Monto FROM ctpartidapol" & SQL3)
      If rTem!Monto = 0 Then conn.Execute "DELETE FROM ctpartidapol" & SQL3
    End If
  End If
End Sub
Function RecuperaXML(Archi, Opc)
  Dim rTem As ADODB.Recordset
'0=Recupera solo los datos del xml,   1= Importacion masiva
  Cade = ""
  Open Archi For Input As #7
  Do Until EOF(7)
    Line Input #7, Cad
    Cade = Cade & Cad
  Loop
  Close #7
  tUUID = Buscar("UUID", Cade)
  tTipo = Buscar("tipoDeComprobante", Cade)
  tEmi = Buscar("<cfdi:Emisor", Cade)
  tPr = Split(tEmi, "|")
  RFCEmi = tPr(0)
  Cad = Buscar("<cfdi:Receptor", Cade)
  cam = Split(Cad, "|")
  RFCRec = cam(0)
  tFec = Buscar("fecha", Cade)
  tFec = Fecha6(Mid(tFec, 9, 2) & Mid(tFec, 6, 2) & Mid(tFec, 3, 2), 0)
  tFol = Buscar("folio", Cade)
  tDes = Buscar("descuento", Cade)
  tSub = Buscar("subTotal", Cade)
  tTot = Buscar("total", Cade)
  tCamb = Buscar("TipoCambio", Cade)
  tCamb = Val(tCamb)
  If tCamb = 0 Then tCamb = 1
  tMone = Buscar("Moneda", Cade)
' 1     2    3   4    5    6    7    8    9    0     1   2   3     4     5    6     7   8    9      0
'UUID,rfcE,NomE,DirE,NumE,NumI,ColE,LocE,MunE,EstE,PaiE,CP,RFCRe,Fecha,Folio,Desc,SubT,Tot,Clasif,tCambio
  Cad = tUUID & "|" & tEmi & "|" & RFCRec & "|" & tFec & "|" & tFol & " |" & tDes & "|" & tSub & "|" & tTot & "|"
  If Len(tUUID) < 36 Then
    RecuperaXML = "El XML no esta timbrado": Exit Function
  End If
  If rEmp!rfc <> RFCRec Then
    RecuperaXML = "01 El comprobante no pertenece a " & rEmp!rfc & "|" & Cad: Exit Function
  End If
  Set rTem = conn.Execute("SELECT Clave,ComprasCla,Plazo FROM proveedor WHERE RFC='" & RFCEmi & "'")
  If rTem.RecordCount = 0 Then
    RecuperaXML = "02 No existe el proveedor|" & Cad: Exit Function
  ElseIf rTem!ComprasCla = 0 Then
    RecuperaXML = "03 El proveedor no tiene clasificación|" & Cad: Exit Function
  Else
    clvP = rTem!Clave: ClvC = rTem!ComprasCla: FecV = CDbl(DateValue(tFec)) + rTem!Plazo: FecV = Format(FecV, "dd/mm/yy"): Plazo = rTem!Plazo
    Set rTem = conn.Execute("SELECT * FROM comprascla WHERE Empresa=" & rEmp!Clave & " And Clave=" & rTem!ComprasCla)
    If rTem.RecordCount = 0 Then
      RecuperaXML = "04 No existe la clasificación|" & Cad: Exit Function
    Else
      If Len(TNull(rTem!Cuenta, 0)) = 0 Then
        RecuperaXML = "05 La clasificación no tiene cuenta contable|" & Cad & rTem!Nombre: Exit Function
      Else
        Cad = Cad & rTem!Nombre
      End If
    End If
  End If
  Cad = Cad & "|" & tCamb
  Set rTem = conn.Execute("SELECT UUID FROM compras WHERE UUID='" & tUUID & "' And Status<>3")
  If rTem.RecordCount > 0 Then
    RecuperaXML = "06 Ya se capturo el comprobante|" & Cad: Exit Function
  End If
  If tTipo <> "ingreso" Then
    RecuperaXML = "07 No es un ingreso (" & tTipo & ")|" & Cad: Exit Function
  End If
  
  tConc = Buscar("<cfdi:Conceptos", Cade)  'Cant,Descrip,Pre,Importe
  tConc = Split(tConc, "||"): suma = 0
  For a = 0 To UBound(tConc)
    cam = Split(tConc(a), "|")
    suma = suma + Valor(cam(3))
  Next
'<implocal:ImpuestosLocales
  tConc = Buscar("<implocal:ImpuestosLocales", Cade)  'Devuelve Impuesto|Importe|Tasa
'ISH se agrega a las partidas del documento
  If Len(tConc) > 3 Then
    tConc = Split(tConc, "||")
    For a = 0 To UBound(tConc)
      cam = Split(tConc(a), "|")
      If cam(0) = "ISH" Then suma = suma + Valor(cam(1)) Else MsgBox "Impuesto no considerado " & cam(0)
    Next
  End If
'Impuestos RETENIDOS
  tImpuR = Buscar("<cfdi:Retenciones", Cade)  'Devuelve Impuesto|Importe|Tasa
  tConc = Split(tImpuR, "||"): tImpR = 0
  For a = 0 To UBound(tConc)
    cam = Split(tConc(a), "|")
    If cam(0) = "IVA" Then tImpR = tImpR + Valor(cam(1))
    suma = suma - Valor(cam(1))
  Next
'Impuestos
  tTras = Buscar("<cfdi:Traslados", Cade)  'Devuelve Impuesto|Importe|Tasa
  tConc = Split(tTras, "||")
  tImp = 0: tIEPS = 0
  For a = 0 To UBound(tConc)
    cam = Split(tConc(a), "|")
    If cam(0) = "IVA" Then tImp = tImp + Valor(cam(1))
    If cam(0) = "IEPS" Then tIEPS = tIEPS + Valor(cam(1))
    suma = suma + Valor(cam(1))
  Next
  suma = suma - Valor(tDes)
  suma = Format(suma, "0.0000"): tTot = Format(tTot, "0.0000")
  dif = suma - tTot
  If dif <> 0 Then
    If Not (dif > -0.1 And dif < 0.1) Then
      RecuperaXML = "08 Total diferente " & suma & " <> " & tTot & "|" & Cad: Exit Function
    End If
  End If
'Renglones del documento
  tConc = Buscar("<cfdi:Conceptos", Cade)  'Cant,Descrip,Pre,Importe
'<implocal:ImpuestosLocales ISH lo graba como una partida del documento
  tImpu = Buscar("<implocal:ImpuestosLocales", Cade)  'Devuelve Impuesto|Importe|Tasa
  If Opc = 0 Then
'DatosGenerales && Concepto && ImpuestroLocal && Traslado && Retenido
    RecuperaXML = "00|" & Cad & "&&" & tConc & "&&" & tImpu & "&&" & tTras & "&&" & tImpuR
    Exit Function
  ElseIf Opc = 1 Then
    Set rTem = conn.Execute("SELECT Max(Numero) AS Mayor FROM compras WHERE Empresa=" & rEmp!Clave & " And ClvCat=2")
    If tImp = 0 Then tIVA = 0 Else tIVA = 16
    nuevo = Valor(rTem!Mayor) + 1: FecE = Fecha6(tFec, 3):
    If Len(tFol) > 10 Then tFol = Right(tFol, 10) Else tFol = Space(10 - Len(tFol)) & tFol
    If tMone = "USD" Then tMone = 2 Else tMone = 1
    sql = "INSERT INTO compras(Empresa,ClvCat,Numero,Provee,FechaEla,FechaRec,FechaVen,FechaCap,Moneda,Plazo,IVA,SubTotal,Impuesto," _
    & "SubTotal2,Total,Descuento,Saldo,Status,tCambio,Factura,IEPS,UUID,RetIVA) VALUES(" & rEmp!Clave & ",2," & nuevo & "," & clvP & "," _
    & FecE & "," & FecE & "," & Fecha6(FecV, 3) & "," & Fecha6(date, 3) & "," & tMone & "," & Plazo & "," & tIVA & "," _
    & Valor(tSub) + Valor(tDes) & "," & tImp & "," & Valor(tSub) & "," & Valor(tTot) & "," & Valor(tDes) & "," & Valor(tTot) & ",1," _
    & tCamb & ",'" & tFol & "'," & tIEPS & ",'" & tUUID & "'," & tImpR & ")"
    conn.Execute sql
    tConc = Split(tConc, "||"): Cont = 1
    For a = 0 To UBound(tConc)
      cam = Split(tConc(a), "|")
      tDes = Replace(cam(1), "'", "")
      sql = "INSERT INTO partidacog(Empresa,ClvCat,Numero,Renglon,Descrip,Cantidad,Precio,Clasif,Importe) VALUES(" & rEmp!Clave & ",2," _
      & nuevo & "," & Cont & ",'" & tDes & "'," & cam(0) & "," & Valor(cam(2)) & "," & ClvC & "," & Valor(cam(3)) & ")"
      conn.Execute sql
      Cont = Cont + 1
      suma = suma + Valor(cam(3))
    Next
    tConc = Split(tImpu, "||")
    For a = 0 To UBound(tConc)
      cam = Split(tConc(a), "|")
      Descri = "Impuesto local " & cam(0) & " tasa " & cam(2) & "%"
      sql = "INSERT INTO partidacog(Empresa,ClvCat,Numero,Renglon,Descrip,Cantidad,Precio,Clasif,Importe) VALUES(" & rEmp!Clave & ",2," _
      & nuevo & "," & Cont & ",'" & Descri & "',1," & Valor(cam(1)) & "," & ClvC & "," & Valor(cam(1)) & ")"
      conn.Execute sql
      Cont = Cont + 1
      suma = suma + Valor(cam(1))
    Next
  End If
'Organizacion de archivos
  tDir = Directorio & rEmp!Directorio & Year(tFec) & "\" & RFCEmi: tDir = Replace(tDir, "/", "\")
  If Dir(tDir, vbDirectory) = "" Then CreaDirectorio tDir
  Destino = tDir & "\" & Format(tFec, "yyyy-mm-dd") & "_" & tUUID & ".xml"
  FileCopy Archi, Destino
  Kill Archi
  Contabilizar 1, 2, nuevo, 1
  RecuperaXML = "00|" & Cad
  DoEvents
End Function
Sub Colores(xForm As Form)
'  Label.wordwrap= true HACE QUE NO CAMBIE EL COLOR
'Tambien llena los combo con recorderset
  xForm.backcolor = rEmp!Fondo
  For Each x In xForm.Controls
    If TypeName(x) = "OptionButton" Then
      If x.Tag <> "NO" Then
        x.ForeColor = rEmp!LetraLa
        x.backcolor = rEmp!FondoLa
      End If
    ElseIf TypeName(x) = "CheckBox" Then
      If Len(x.DataField) > 1 Then x.backcolor = rEmp!Fondo
    ElseIf TypeName(x) = "Label" Then
      If Not x.WordWrap Then
        x.ForeColor = rEmp!LetraLa
        x.backcolor = rEmp!FondoLa
      End If
    ElseIf TypeName(x) = "MSFlexGrid" Then
      x.BackColorBkg = rEmp!GridFo
    ElseIf TypeName(x) = "ComboBox" Then
      If Len(x.Tag) > 2 Then
        Set rTem = conn.Execute("SELECT Nombre FROM " & x.Tag)
        If x.Tag = "moneda" Then
          If rTem.RecordCount = 0 Then
            conn.Execute "INSERT INTO moneda VALUES(1,'M.N.',1)"
            rTem.Requery
          End If
        End If
        If x.Tag = "pais" Then
          If rTem.RecordCount = 0 Then
            conn.Execute "INSERT INTO pais VALUES(" & rEmp!Clave & ",1,'MEXICO')"
            rTem.Requery
          End If
        End If
        Do Until rTem.EOF
          If Not IsNull(rTem!Nombre) Then x.AddItem rTem!Nombre
          rTem.MoveNext
        Loop
      End If
    End If
  Next
End Sub
Function Contabilizar(Opc, ClvC, Numero, Opc2)
'Opc=1  Gastos
  If Opc2 <> 1 And Opc2 <> 2 Then
    MsgBox "La opc2 es diferente de 1 y 2.", 48, rEmp!Nombre: Exit Function
  End If
'Opc2=1 Contabilizar, =2 Resta la contabilidad
  Dim rTem As ADODB.Recordset
  Dim rDoc As ADODB.Recordset
  Dim rCue As ADODB.Recordset
  Contabilizar = False
  Set rDoc = conn.Execute("SELECT a.*,b.Nombre,b.RFC FROM compras a INNER JOIN proveedor b ON a.Empresa=b.Empresa And a.Provee=b.Clave " _
  & "WHERE a.Empresa=" & rEmp!Clave & " And ClvCat=" & ClvC & " And Numero=" & Numero)
  If rDoc!moneda = 1 Then
    tCamb = 1
  Else
    Set rTem = conn.Execute("SELECT * FROM tcambio WHERE Fecha=" & Fecha6(rDoc!FechaEla, 3))
    If rTem.RecordCount = 0 Then
      MsgBox "No existe tipo de cambio fecha: " & rDoc!FechaEla, 48, rEmp!Nombre
      tCamb = 1
    Else
      tCamb = rTem!TCambio
    End If
  End If
  If rDoc.RecordCount = 0 Then
    MsgBox "Error no existe el documento " & Numero, 48, rEmp!Nombre: Exit Function
  End If
  If ClvC = 1 Or ClvC = 2 Then
    If Len(TNull(rDoc!UUID, 0)) < 30 Then
      sql = "SELECT * FROM ctpartidapol WHERE ClvDoc=" & ClvC & " And Folio=" & Numero
    Else
      sql = "SELECT * FROM ctpartidapol WHERE UUID='" & rDoc!UUID & "'"
    End If
    Set rTem = conn.Execute(sql)
  Else
    Set rTem = conn.Execute("SELECT * FROM ctpartidapol WHERE ClvDoc=" & ClvC & " And Folio=" & Numero)
  End If
  If Opc2 = 1 Then
    If rTem.RecordCount > 0 Then
      MsgBox "Ya se capturo el folio " & Numero, 48, rEmp!Nombre: Exit Function
    End If
  Else
    If rTem.RecordCount = 0 Then
      MsgBox "No existe el documento digital " & Numero, 48, rEmp!Nombre: Exit Function
    End If
  End If
  If rDoc!ClvCat = 2 Then tClv = 1 Else tClv = 2
  tPol = 3: tEje = Year(rDoc!FechaEla): tMes = Month(rDoc!FechaEla): nPol = Day(rDoc!FechaEla): tFec = Fecha6(rDoc!FechaEla, 3)
  sql = "SELECT a.*,b.Cuenta,b.Nombre FROM partidacog a INNER JOIN comprascla b ON a.Empresa=b.Empresa And a.Clasif=b.Clave WHERE " _
  & "a.Empresa=" & rDoc!Empresa & " And ClvCat=" & rDoc!ClvCat & " And Numero=" & rDoc!Numero
  Set rPar = conn.Execute(sql)
  Cad = ""
'Conceptos de la compra
  Do Until rPar.EOF
    If rPar!Importe <> 0 Then
      tMon = rPar!Importe * tCamb
'Contabiliza el tipo de gasto partida
      tConc = rPar!Nombre & " " & rDoc!FechaEla
      ctaGasto = rPar!Cuenta
      cPartidaPol tClv, rDoc!FechaEla, tPol, rPar!Cuenta, 1, tMon, tConc, True, ClvC, 0, "", Opc2
    End If
    rPar.MoveNext
  Loop
'Proveedor
  Set rcuep = conn.Execute("SELECT * FROM Cuentas")
  tLen = Right(ctEmp!Niveles, 1)
  Pro = String(tLen - Len(rDoc!Provee), "0") & rDoc!Provee
  Cta = rcuep!Proveedores & "-" & Pro
  Set rCue = conn.Execute("SELECT ID,Nombre FROM ctcuentas WHERE Cuenta2='" & Cta & "'")
'Alta del proveedor
  If rCue.RecordCount = 0 Then
    Cue = ctCompleta(Cta, 2)
    nNiv = ctCompleta(Cta, 3)
    NombreCu = Space((nNiv - 1) * 5) & rDoc!Nombre
    sql = "INSERT INTO ctcuentas(Cuenta,Cuenta2,Nombre,Tipo,RFC,FechaA,Status,Nivel,Codigo,Detalle) VALUES('" & Cue & "','" & Cta & "','" _
    & NombreCu & "','D','" & rDoc!rfc & "'," & Fecha6(date, 3) & ",-1," & nNiv & ",'201-01',-1)"
    conn.Execute sql
    rCue.Requery
  End If
  tConcep = Trim(rCue!Nombre) & " F_" & Trim(rDoc!Factura)
  cPartidaPol tClv, rDoc!FechaEla, tPol, rCue!Id, 2, rDoc!Total * tCamb, tConcep, False, rDoc!ClvCat, rDoc!Numero, rDoc!UUID, Opc2
'Retencion de IVA
  If rDoc!retIVA > 0 Then
    tConcep = "IVA retenido F-" & Trim(rDoc!Factura)
    cPartidaPol tClv, rDoc!FechaEla, tPol, rcuep!RetencionIVA, 2, rDoc!retIVA * tCamb, tConcep, False, rDoc!ClvCat, rDoc!Numero, "", Opc2
  End If

'IVA POR Acreditar
  If rDoc!Impuesto > 0 Then
    tConcep = "IVA por acreditar del " & rDoc!FechaEla
    cPartidaPol tClv, rDoc!FechaEla, tPol, rcuep!IVAAcreNP, 1, rDoc!Impuesto * tCamb, tConcep, True, 0, 0, "", Opc2
  End If
'IEPSAcreNP
  If rDoc!IEPS > 0 Then
    If Len(TNull(rcuep!IEPSAcreNP, 0)) = 0 Then Cta = ctaGasto Else Cta = rcuep!IEPSAcreNP
    tConcep = "IEPS por acreditar del " & rDoc!FechaEla
    cPartidaPol tClv, rDoc!FechaEla, tPol, Cta, 1, rDoc!IEPS * tCamb, tConcep, True, 0, 0, "", Opc2
  End If
'Descuento
  If rDoc!Descuento > 0 Then
    tConcep = "Descuento F_" & Trim(rDoc!Descuento)
    cPartidaPol tClv, rDoc!FechaEla, tPol, rcuep!ComprasDesc, 2, rDoc!Descuento * tCamb, tConcep, False, 0, 0, "", Opc2
  End If

'Agregar Poliza
  SQL2 = " WHERE ClvCat=" & tClv & " And Ejercicio=" & tEje & " And Periodo=" & tMes & " And Numero=" & nPol & " And Tipo=" & tPol _
  & " And Numero=" & nPol
  If Opc2 = 1 Then
    Set rTem = conn.Execute("SELECT Total FROM ctpoliza" & SQL2)
    If rTem.RecordCount = 0 Then
      sql = "INSERT INTO ctpoliza(ClvCat,Ejercicio,Periodo,Numero,Tipo,Fecha,Concepto,Total) VALUES(" & tClv & "," & tEje & "," & tMes _
      & "," & nPol & "," & tPol & "," & tFec & ",'GASTOS DEL " & Format(rDoc!FechaEla, "dd/mm/yy") & "'," & rDoc!Total * tCamb & ")"
      conn.Execute sql
    Else
      conn.Execute "UPDATE ctpoliza SET Total=Total+" & rDoc!Total * tCamb & SQL2
    End If
  Else
    Set rTem = conn.Execute("SELECT Sum(Monto) as Total FROM ctpartidapol" & SQL2 & " And TipoM='1'")
    tTot = Valor(rTem!Total)
    If tTot < 0.2 Then
      conn.Execute "DELETE FROM ctpoliza" & SQL2
      conn.Execute "DELETE FROM ctpartidapol" & SQL2
    Else
      Set rTem = conn.Execute("UPDATE ctpoliza SET Total=" & tTot & SQL2)
      conn.Execute "DELETE FROM ctpartidapol WHERE ClvDoc=" & rDoc!ClvCat & " And Numero=" & rDoc!Numero
    End If
  End If
End Function
Function Traduce(Numero As Variant, Mone As Variant) As String
  Dim unidad(0 To 9) As String
  Dim decena(0 To 9) As String
  Dim centena(0 To 10) As String
  Dim deci(0 To 9) As String
  Dim otros(0 To 15) As String
  If IsNull(Numero) Then
    Traduce = "": Exit Function
  ElseIf Len(Numero) = 0 Then
    Traduce = "": Exit Function
  End If
  Nume = Format(Numero, "0.00")
  strNum = Mid(Nume, 1, Len(Nume) - 3)
  nDeci = Mid(Nume, Len(Nume) - 1)
  If Mone = 1 Then
    unidad(1) = "UN": unidad(2) = "DOS": unidad(3) = "TRES": unidad(4) = "CUATRO": unidad(5) = "CINCO"
    unidad(6) = "SEIS": unidad(7) = "SIETE": unidad(8) = "OCHO": unidad(9) = "NUEVE"
    decena(1) = "DIEZ": decena(2) = "VEINTE": decena(3) = "TREINTA": decena(4) = "CUARENTA": decena(5) = "CINCUENTA"
    decena(6) = "SESENTA": decena(7) = "SETENTA": decena(8) = "OCHENTA": decena(9) = "NOVENTA"
    centena(1) = "CIENTO": centena(2) = "DOSCIENTOS": centena(3) = "TRESCIENTOS": centena(4) = "CUATROCIENTOS": centena(5) = "QUINIENTOS"
    centena(6) = "SEISCIENTOS": centena(7) = "SETECIENTOS": centena(8) = "OCHOCIENTOS": centena(9) = "NOVECIENTOS"
    deci(1) = "DIECI": deci(2) = "VEINTI": deci(3) = "TREINTA Y ": deci(4) = "CUARENTA Y ": deci(5) = "CINCUENTA Y "
    deci(6) = "SESENTA Y ": deci(7) = "SETENTA Y ": deci(8) = "OCHENTA Y ": deci(9) = "NOVENTA Y "
    otros(11) = "ONCE": otros(12) = "DOCE": otros(13) = "TRECE": otros(14) = "CATORCE": otros(15) = "QUINCE"
  ElseIf Mone = 2 Then
    unidad(1) = "ONE": unidad(2) = "TWO": unidad(3) = "THREE": unidad(4) = "FOUR": unidad(5) = "FIVE"
    unidad(6) = "SIX": unidad(7) = "SEVEN": unidad(8) = "EIGHT": unidad(9) = "NINE"
    decena(1) = "TEN": decena(2) = "TWENTY": decena(3) = "THIRTY": decena(4) = "FORTY": decena(5) = "FIFTY"
    decena(6) = "SIXTY": decena(7) = "SEVENTY": decena(8) = "EIGHTY": decena(9) = "NINETY"
    centena(1) = "ONE HUNDRED": centena(2) = "TWO HUNDRED": centena(3) = "THREE HUNDRED": centena(4) = "FOUR HUNDRED": centena(5) = "FIVE HUNDRED"
    centena(6) = "SIX HUNDRED": centena(7) = "SEVEN HUNDRED": centena(8) = "EIGHT HUNDRED": centena(9) = "NINE HUNDRED"
    deci(1) = "TEN": deci(2) = "TWENTY": deci(3) = "THIRTY ": deci(4) = "FORTY ": deci(5) = "FIFTY "
    deci(6) = "SIXTY ": deci(7) = "SEVENTY ": deci(8) = "EIGHTY ": deci(9) = "NINETY "
    otros(11) = "ELEVEN": otros(12) = "TWELVE": otros(13) = "THIRTEEN": otros(14) = "FOURTEEN": otros(15) = "FIFTEEN"
  End If
  ReDim strN(1 To 4)
  Cad = String(12 - Len(strNum), "0") & strNum
  pos = 1
  For a = 1 To 12 Step 3
    strN(pos) = Mid(Cad, pos * 3 - 2, 3): pos = pos + 1
  Next
  Cont = 4: Re = ""
  For a = 1 To 4
    Uni = "": dec = "": Cen = ""
    de = Val(Right(strN(a), 2)): CE = Val(Left(strN(a), 1))
    If strN(a) <> "000" Then
      If strN(a) = "100" Then
        If Mone = 1 Then Cen = "CIEN" Else Cen = "ONE HUNDRED"
      Else
        If CE > 0 Then Cen = centena(CE)
      End If
      If Right(strN(a), 1) = "0" Then
        de = de \ 10: dec = decena(de)
      ElseIf de > 10 And de < 16 Then
        dec = otros(de)
      Else
        Uni = unidad(Val(Right(strN(a), 1))): k = Val(Mid(strN(a), 2, 1)): dec = deci(k)
      End If
      If Len(Cen) > 0 Then Re = Re & Cen & " "
      If Len(dec) > 0 Then Re = Re & dec
      If Len(Uni) > 0 Then Re = Re & Uni & " "
      If a = 1 Or a = 3 Then
        If Mone = 1 Then Re = Re & " MIL " Else Re = Re & " THOUSAND "
      End If
      If a = 2 Then
        If moneda = 1 Then
          If Val(strN(2)) = 1 Then Re = Re & " MILLON " Else Re = Re & " MILLONES "
        Else
          If Val(strN(2)) = 1 Then Re = Re & " MILLION " Else Re = Re & " MILLIONS "
        End If
      End If
    End If
  Next
  Re = Replace(Re, "  ", " ")
  Re = Trim(Re)
  If Mone = 1 Then Re = Re & " PESOS " & nDeci & "/100 M.N." Else Re = Re & " DLLS " & nDeci & "/100 US CY"
  Traduce = Re
End Function
Function Mail(Para, Asunto, Archivos, Mensaje)
  tCor = TNull(Para, 0)
  If tCor = "" Then
    Mail = "El cliente no tiene correo."
    Exit Function
  End If
  If InStr(tCor, "@") = 0 Or InStr(tCor, ".") = 0 Then
    Mail = "Existe un posible error en el nombre de la cuenta de correo: " & tCor
    Exit Function
  End If
  If IsNull(rEmp!servidor) Then
    MsgBox "No se han definido los parametros del correo.", 48, rEmp!Nombre: Exit Function
  End If
  Set oMail = New clsCDOmail
  With oMail
    .servidor = rEmp!servidor
    .puerto = rEmp!CPuerto
    .UseAuntentificacion = rEmp!Autentificacion
    .ssl = rEmp!Pssl
    .Usuario = rEmp!Usuario
    .PassWord = rEmp!CPass
    .Asunto = Asunto
    .Adjunto = Archivos
    .de = rEmp!Usuario
    .Para = Para
    .Mensaje = Mensaje
    varx = .Enviar_Backup
    If varx = "T" Then Mail = "OK" Else Mail = varx
  End With
  Set oMail = Nothing
End Function

Public Function DevolverDireccionMemoria(Direccion As Long) As Long
    DevolverDireccionMemoria = Direccion  'zip
End Function

Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function
Function FuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function
Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function
Function FuncionParaProcesarServicios(ByRef fname As CBChar, ByVal x As Long) As Long
  FuncionParaProcesarServicios = 0
End Function
Function ctCompleta(Obj, Opc)
'1= Agrega los guiones para la captura
'2= Quitar guiones y completar con 0 20 espacios
'3= Devuelve el numero de niveleS
'4= Pone los guiones a las cuentas
  If Len(Obj) = 0 Then
    ctCompleta = ""
  ElseIf Opc = 1 Then         'AGREGA LOS GUIENES EN LA CAPTURA
    cam = Split(ctEmp!Niveles, "-"): Cad = Replace(Obj, "-", ""): nLe = Len(Cad): TMAS = "": Guion = "": suma = 0
    For a = 0 To UBound(cam)
      TMAS = TMAS & Guion & String(Val(cam(a)), "0"): suma = suma + Val(cam(a)): Guion = "-"
    Next
    If nLe >= suma Then
      ctCompleta = Mid(Obj, 1, Len(TMAS)): Exit Function
    End If
    cadN = "": Cont = 1
    For a = 1 To Len(TMAS)
      If Cont <= nLe Then
        If Mid(TMAS, a, 1) = "0" Then
          cadN = cadN & Mid(Cad, Cont, 1): Cont = Cont + 1
        Else
          cadN = cadN & "-"
        End If
      Else
        Exit For
      End If
    Next
    ctCompleta = cadN
  ElseIf Opc = 2 Then
    Cad = Replace(Obj, "-", "")
    ctCompleta = Cad & String(20 - Len(Cad), "0")
  ElseIf Opc = 3 Then
    cam = Split(Obj, "-"): Cont = 0
    For a = 0 To UBound(cam)
      If Val(cam(a)) > 0 Then Cont = Cont + 1
    Next
    ctCompleta = Cont
  ElseIf Opc = 4 Then
    Obj = Replace(Obj, "-", "")
    If Len(Obj) <> 20 Then Obj = Obj & String(20 - Len(Obj), "0")
    cam = Split(ctEmp!Niveles, "-")
    Cad = "": suma = 1
    For a = 0 To UBound(cam)
      Cad = Cad & Mid(Obj, suma, Val(cam(a)))
      suma = suma + Val(cam(a))
      If a < UBound(cam) Then Cad = Cad & "-"
    Next
    ctCompleta = Cad
  End If
End Function

Function rFormulario(FormToCheck As Form) As Integer
  Dim y As Integer
  For y = 0 To Forms.Count - 1
    If Forms(y) Is FormToCheck Then
      rFormulario = True
      Exit Function
    End If
  Next
  rFormulario = False
End Function
Sub SaveFoto(mio As Form, Obj As PictureBox, sql, Arch, Resiz As Boolean, h, w)
  Dim hNew2 As Long
  Dim tFot As StdPicture
  If Dir(Arch) <> "" Then
    Set tFot = LoadPicture(Arch)
    Alto = Round(mio.ScaleY(tFot.Height, vbHimetric, vbPixels))
    Ancho = Round(mio.ScaleX(tFot.Width, vbHimetric, vbPixels))
    If h > 0 Then
      Porc = h / Alto
    ElseIf w > 0 Then
      Porc = w / Ancho
    Else
      If Ancho > Alto Then Porc = 230 / Ancho Else Porc = 230 / Alto
    End If
    hNew2 = CopyImage(tFot, IMAGE_BITMAP, CInt(Ancho * Porc), CInt(Alto * Porc), LR_COPYRETURNORG)
    tw = 11.9
    'If Resiz Then
      Obj.Height = CInt(Alto * Porc * tw)
      Obj.Width = CInt(Ancho * Porc * tw)
    'End If
    OpenClipboard mio.hwnd
    EmptyClipboard
    SetClipboardData CF_BITMAP, hNew2
    CloseClipboard
    Set Obj.Picture = Clipboard.GetData(2)
'Fin Reducir tamaño
    SaveJPEG Arch, Obj, True, 50
  End If
End Sub
Sub CreaDirectorio(TDire As Variant)
  Dim i As Integer
  Dim Array_Dir As Variant
  Dim Sub_Dir As String
  Dim El_Path As String
  Dim Uni As String
  If InStr(TDire, ":") > 0 Then
    Uni = Mid(TDire, 1, 2)
    If Not ValidaUnidad(Uni) Then Exit Sub
  End If
  El_Path = TDire
  If El_Path = vbNullString Then
      Exit Sub
  End If
  Array_Dir = Split(El_Path, "\")
  El_Path = vbNullString
  For i = LBound(Array_Dir) To UBound(Array_Dir)
    Sub_Dir = Array_Dir(i)
    If Sub_Dir <> vbNullString Then
      El_Path = El_Path & Sub_Dir & "\"
      If Right$(Sub_Dir, 1) <> ":" Then
        If Dir(El_Path, vbDirectory) = vbNullString Then Call MkDir(El_Path)
      End If
    End If
  Next
End Sub
Function ValidaUnidad(Uni As Variant)
  Dim objWMI
  Dim UActual
  Dim objDisco
  Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
  Set UActual = objWMI.ExecQuery("Select * from Win32_LogicalDisk")
  ValidaUnidad = False
  For Each objDisco In UActual
    If UCase(Uni) = UCase(objDisco.DeviceID) Then
      ValidaUnidad = True
      Exit For
    End If
  Next
End Function
Private Function SaveJPEG(ByVal Filename As String, Pic As PictureBox, Optional ByVal Overwrite As Boolean = True, Optional ByVal Quality As Byte = 90) As Boolean
    Dim JPEGclass As cJpeg
    Dim m_Picture As IPictureDisp
    Dim m_DC As Long
    Dim m_Millimeter As Single
'    m_Millimeter = ScaleX(100, vbPixels, vbMillimeters)
    m_Millimeter = 21.16669
   ' m_Millimeter = 12.1666
    Set m_Picture = Pic
    m_DC = Pic.hDC
    'this is not my code....from PSC
    'initialize class
    Set JPEGclass = New cJpeg
    'check there is image to save and the filename string is not empty
    If m_DC <> 0 And LenB(Filename) > 0 Then
        'check for valid quality
        If Quality < 1 Then Quality = 1
        If Quality > 100 Then Quality = 100
        'set quality
        JPEGclass.Quality = Quality
        'save in full color
        JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
        'copy image from hDC
        px = 0 'Exedente en pixeles
        If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter - px), CLng(m_Picture.Height / m_Millimeter - px)) = 0 Then
            'if overwrite is set and file exists, delete the file
'            If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
            'save file and return True if success
            SaveJPEG = JPEGclass.SaveFile(Filename) = 0
        End If
    End If
    'clear memory
    Set JPEGclass = Nothing
End Function


Public Function Leer_foto(Imagen As Variant) As Variant
  Dim i As Integer
  Dim x As Long
  Dim tb() As Byte
  i = FreeFile
  Dim ll_lon As Long
  If Dir(Imagen) = "" Then
    Leer_foto = Null
    Exit Function
  End If
  Open Imagen For Binary As i
  ll_lon = LOF(i)
  ReDim Preserve tb(ll_lon + 1)
      Get #i, , tb()
  Close #i
  Leer_foto = tb
End Function

Sub ActualizaMP()
  conn.Execute "INSERT INTO metodopago VALUES(1,'Efectivo')"
  conn.Execute "INSERT INTO metodopago VALUES(2,'Cheque nominativo')"
  conn.Execute "INSERT INTO metodopago VALUES(3,'Transferencia electronica de fondos')"
  conn.Execute "INSERT INTO metodopago VALUES(4,'Tarjeta de Credito')"
  conn.Execute "INSERT INTO metodopago VALUES(5,'Monedero Electronico')"
  conn.Execute "INSERT INTO metodopago VALUES(6,'Dinero electronico')"
  conn.Execute "INSERT INTO metodopago VALUES(8,'Vales de despensa')"
  conn.Execute "INSERT INTO metodopago VALUES(28,'Tarjeta de Debito')"
  conn.Execute "INSERT INTO metodopago VALUES(29,'Tarjeta de Servicio')"
  conn.Execute "INSERT INTO metodopago VALUES(99,'Otros')"
End Sub
Sub CambiaColor(tGri)
  tGri.Redraw = False
  If Len(TNull(rEmp!GRIDL1, 0)) = 0 Then
    tCol1 = &HFEFEF3: tcol2 = &HFFFF80
  Else
    tCol1 = rEmp!GRIDL1: tcol2 = rEmp!GridL2
  End If
  For a = tGri.FixedRows To tGri.Rows - 1
    tGri.Row = a
    For b = tGri.FixedCols To tGri.Cols - 1
      If a Mod 2 = 0 Then Color3 = tCol1 Else Color3 = tcol2
      tGri.Col = b: tGri.CellBackColor = Color3
    Next
  Next
  tGri.Redraw = True
End Sub

Sub Grid(Obj, tTit, tAnc, tAli)
  cad1 = Split(tTit, ",")
  cad2 = Split(tAnc, ",")
  Obj.Cols = UBound(cad1) + 1
  Obj.Row = 0
  For a = 0 To Obj.Cols - 1
    Obj.Col = a
    Obj.CellFontSize = 10
    Obj.TextMatrix(0, a) = cad1(a)
    Obj.ColWidth(a) = Val(cad2(a)) * 120
  Next
  If Len(tAli) > 0 Then
    Cad3 = Split(tAli, ",")
    For a = 0 To UBound(Cad3)
      Obj.ColAlignment(a) = Val(Cad3(a))
    Next
  End If
End Sub

Sub TextL(Obj)
  If Obj.DataMember = "F" Then Obj.Text = Fecha6(Obj.Text, 0)
  Obj.backcolor = QBColor(15)
End Sub
Sub TextLon(xForm As Form, rs)
  Dim rTab As ADODB.Recordset
  Set rTab = conn.Execute("SELECT * FROM " & rs & " LIMIT 1")
  For Each x In xForm.Controls
    If TypeName(x) = "TextBox" Then
      If Len(x.DataField) > 0 Then
        If x.DataMember = "T" Then
          If MostrarTextV(rTab, x.DataField) Then
            If rTab.Fields(x.DataField).Type = 201 Then
              x.MaxLength = 0
            Else
              x.MaxLength = rTab.Fields(x.DataField).DefinedSize
            End If
          End If
        End If
      End If
    End If
  Next
End Sub
Function SqlCad(xForm, tTab, Opc, tSql)
  Dim rCat As ADODB.Recordset
  If Opc = "A" Then
    If tSql = "" Or tTab = "compras" Then
      sql = "INSERT INTO " & tTab & " (Empresa"
      SQL2 = rEmp!Clave: cm = ","
    Else
      sql = "INSERT INTO " & tTab & " ("
      SQL2 = "": cm = ""
    End If
    If tSql = "ClvCat" Then
      sql = sql & cm & "ClvCat": SQL2 = SQL2 & cm & ClaveCatF: cm = ","
    ElseIf tSql = "Empresa" Then
      sql = sql & "Empresa": SQL2 = rEmp!Clave: cm = ","
    End If
    For Each x In xForm.Controls
      If TypeName(x) <> "CommonDialog" Then
      If x.ToolTipText = "1" Then
        If TypeName(x) = "TextBox" And Len(x.DataField) > 0 Then
          x.Text = Trim(x.Text)
          If x.DataField = "Cliente" Then
            sql = sql & ",ClvCli": SQL2 = SQL2 & "," & ClaveCatF
          End If
          sql = sql & cm & x.DataField
          If x.DataMember = "N" Then SQL2 = SQL2 & cm & Valor(x.Text)
          If x.DataMember = "T" Then SQL2 = SQL2 & cm & "'" & x.Text & "'"
          If x.DataMember = "F" Then SQL2 = SQL2 & cm & Fecha6(x.Text, 3)
          If x.DataMember = "FH" Then SQL2 = SQL2 & cm & Fecha6(x.Text, 4)
          cm = ","
        ElseIf TypeName(x) = "ComboBox" And Len(x.DataField) > 0 Then
          sql = sql & "," & x.DataField
          If x.Tag = "" Then
            SQL2 = SQL2 & "," & x.ListIndex
          Else
            Set rCat = conn.Execute("SELECT Clave FROM " & x.Tag & " WHERE Nombre='" & x & "'")
            If rCat.RecordCount = 0 Then
              If x.DataMember = "N" Then SQL2 = SQL2 & ",0"
              If x.DataMember = "T" Then SQL2 = SQL2 & ",''"
            Else
              If x.DataMember = "N" Then SQL2 = SQL2 & "," & rCat!Clave
              If x.DataMember = "T" Then SQL2 = SQL2 & ",'" & rCat!Clave & "'"
            End If
          End If
        ElseIf TypeName(x) = "Label" Then
          sql = sql & "," & x.DataField
          If x.DataMember = "N" Then SQL2 = SQL2 & "," & Valor(x.Caption)
          If x.DataMember = "T" Then SQL2 = SQL2 & ",'" & x.Caption & "'"
          If x.DataMember = "F" Then SQL2 = SQL2 & "," & Fecha6(x.Caption, 3)
        ElseIf TypeName(x) = "CheckBox" Then
          sql = sql & "," & x.DataField
          If x.Value = 0 Then tVal = 0 Else tVal = -1
          SQL2 = SQL2 & "," & tVal
        End If
      End If
      End If
    Next
    sql = sql & ") VALUES(" & SQL2 & ")"
    SqlCad = sql
  Else
    sql = "UPDATE " & tTab & " SET "
    SQL2 = "": coma = ""
    For Each x In xForm.Controls
      If TypeName(x) <> "CommonDialog" Then
      If x.ToolTipText = "1" Then
        If TypeName(x) = "TextBox" And Len(x.DataField) > 0 Then
          x.Text = Trim(x.Text)
          If x.DataMember = "N" Then SQL2 = SQL2 & coma & x.DataField & "=" & Valor(x.Text)
          If x.DataMember = "T" Then SQL2 = SQL2 & coma & x.DataField & "='" & x.Text & "'"
          If x.DataMember = "F" Then SQL2 = SQL2 & coma & x.DataField & "=" & Fecha6(x.Text, 3)
          If x.DataMember = "FH" Then SQL2 = SQL2 & coma & x.DataField & "=" & Fecha6(x.Text, 4)
          coma = ","
        ElseIf TypeName(x) = "ComboBox" Then
          If x.Tag = "" Then
            SQL2 = SQL2 & coma & x.DataField & "=" & x.ListIndex
          Else
            Set rCat = conn.Execute("SELECT Clave FROM " & x.Tag & " WHERE Nombre='" & x & "'")
            If rCat.RecordCount = 0 Then
              If x.DataMember = "N" Then SQL2 = SQL2 & coma & x.DataField & "=0"
              If x.DataMember = "T" Then SQL2 = SQL2 & coma & x.DataField & "=''"
            Else
              If x.DataMember = "N" Then SQL2 = SQL2 & coma & x.DataField & "=" & rCat!Clave
              If x.DataMember = "T" Then SQL2 = SQL2 & coma & x.DataField & "='" & rCat!Clave & "'"
            End If
          End If
          coma = ","
        ElseIf TypeName(x) = "Label" Then
          If x.DataMember = "N" Then SQL2 = SQL2 & coma & x.DataField & "=" & Valor(x.Caption)
          If x.DataMember = "T" Then SQL2 = SQL2 & coma & x.DataField & "='" & x.Caption & "'"
          If x.DataMember = "F" Then SQL2 = SQL2 & coma & x.DataField & "=" & Fecha6(x.Caption, 3)
          If x.DataMember = "FH" Then SQL2 = SQL2 & coma & x.DataField & "=" & Fecha6(x.Caption, 4)
          coma = ","
        ElseIf TypeName(x) = "CheckBox" Then
          If x.Value = 0 Then tVal = 0 Else tVal = -1
          SQL2 = SQL2 & coma & x.DataField & "=" & tVal
          coma = ","
        End If
      End If
      End If
    Next
    If SQL2 = "" Then
      SqlCad = ""
    Else
      SqlCad = sql & SQL2
      If Len(tSql) > 0 Then SqlCad = SqlCad & " WHERE " & tSql
    End If
  End If
End Function
Function Fecha6(Cad As Variant, Opc As Integer)
  '0 valida el año
  '1 continua aunque el año sea diferente
  '3 formato sql yyyy/mm/dd
  '4 formato sql yyyy/mm/dd + hora
  If Opc = 3 Then
    If IsDate(Cad) Then
      If CRConn <> "" Then Fecha6 = "#" & Format(Cad, "yyyy/mm/dd") & "#" Else Fecha6 = "'" & Format(Cad, "yyyy/mm/dd") & "'"
    Else
      If CRConn = "" Then Fecha6 = "'0000/00/00'" Else Fecha6 = "Null"
    End If
  ElseIf Opc = 4 Then
    If IsDate(Cad) Then
      If CRConn <> "" Then Fecha6 = "#" & Format(Cad, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "#" Else Fecha6 = "'" & Format(Cad, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & "'"
    Else
      Fecha6 = ""
    End If
  ElseIf Len(Trim(Cad)) = 0 Then
    Fecha6 = ""
  ElseIf IsNull(Cad) Then
    Fecha6 = ""
  ElseIf IsDate(Cad) Then
    FAño = Year(Cad)
    FAñoAc = Year(date)
    If FAño <> FAñoAc Then
      If Opc = 0 Then r = MsgBox("El año es diferente al actual desea continuar.", 36, rEmp!Nombre) Else r = 6
      If r = 6 Then
        If FAño > FAñoAc Then Fecha6 = Format(Cad, "dd/mm/yyyy") Else Fecha6 = Format(Cad, "dd/mm/yy")
      Else
        Fecha6 = "0"
      End If
    Else
      Fecha6 = Format(Cad, "dd/mm/yy")
    End If
  ElseIf Len(Cad) = 6 Then
    FDia = Mid(Cad, 1, 2)
    FMes = Mid(Cad, 3, 2)
    FAño = Mid(Cad, 5, 2)
    FAñoAc = Year(date)
    FAñoAc = Mid(FAñoAc, 3)
    If Not (Val(FDia) >= 1 And Val(FDia) <= 31) Then
      MsgBox "Formato de fecha incorrecto DIA.", 48, rEmp!Nombre
      Fecha6 = "0"
    ElseIf Not (Val(FMes) >= 1 And Val(FMes) <= 12) Then
      MsgBox "Formato de fecha incorrecto MES.", 48, rEmp!Nombre
      Fecha6 = "0"
    ElseIf FAñoAc <> FAño Then
      If Opc = 0 Then r = MsgBox("El año es diferente al actual desea continuar.", 36, rEmp!Nombre) Else r = 6
      If r = 6 Then
        CadFe = FDia & "/" & FMes & "/" & FAño
        Fecha6 = Format(CadFe, "dd/mm/yy")
      Else
        Fecha6 = "0"
      End If
    Else
      CadFe = FDia & "/" & FMes & "/" & FAño
      Fecha6 = Format(CadFe, "dd/mm/yy")
    End If
  Else
    MsgBox "Formato de fecha incorrecto.", 48, rEmp!Nombre
    Fecha6 = "0"
  End If
End Function

Function Valor(Numero As Variant) As Double
  If IsNull(Numero) Then
    Valor = 0
  Else
    For a = 1 To Len(Numero)
      Carac = Mid(Numero, a, 1)
      If Carac >= "0" And Carac <= "9" Or Carac = "." Or Carac = "-" Then Cade = Cade & Carac
    Next
    Valor = Val(Cade)
  End If
End Function
Sub Activa(mio As Form, Opc)
  Encontro = False
  For Each x In mio.Controls
    If TypeName(x) = "PictureBox" Then
      If x.Name = "contenido" Then
        Encontro = True: x.Enabled = Opc: x.backcolor = rEmp!Fondo
        Exit For
      End If
    End If
  Next
  mio.backcolor = rEmp!Fondo
  For Each x In mio.Controls
    If TypeName(x) = "TextBox" Or TypeName(x) = "CheckBox" Or TypeName(x) = "Label" Or TypeName(x) = "ComboBox" Then
      If Len(x.DataField) > 0 Then x.ToolTipText = ""
      If Not Encontro And Len(x.DataField) > 0 Then
        If TypeName(x) = "Label" Then
          If x.DataField = "N" Then x.Visible = Not Opc
        Else
          x.Enabled = Opc
        End If
        If Not Opc And x.Name = "TextE" Then x.Visible = False
      End If
    End If
    If x.Name = "Command1" Then x.Visible = Not Opc
    If x.Name = "Picture1" Then x.Enabled = Not Opc
    If x.Name = "Command2" Then x.Visible = Opc
    If x.Name = "Grid1" Then
      If Opc Then x.SelectionMode = 0 Else x.SelectionMode = 1
    End If
  Next
  MousePointer = 0
End Sub
Sub CAMovi(Clv, Monto, CA, Fecha, Obser, Cobrado)
  Dim rCA As ADODB.Recordset
  If Len(Clv) = 0 Then Exit Sub
  If Monto = 0 Then Exit Sub
  If Val(Clv) > 0 Then
    Set rTem = conn.Execute("SELECT * FROM conceptosca WHERE Empresa=" & rEmp!Clave & " And Clave=" & Clv)
  Else
    Set rTem = conn.Execute("SELECT * FROM conceptosca WHERE Empresa=" & rEmp!Clave & " And Nombre='" & Clv & "'")
  End If
  If rTem.EOF Then
    MsgBox "No existe el concepto CAJA CHICA. (Clv=" & Clv & ")", 48, rEmp!Nombre
  Else
    If Cobrado = 0 Then
      If CA = "+" Then tSal = Valor(rTem!NoCobrado) + Monto Else tSal = Valor(rTem!NoCobrado) - Monto
      conn.Execute "UPDATE conceptosca SET NoCobrado=" & tSal & " WHERE Empresa=" & rEmp!Clave & " And Clave=" & rTem!Clave
    Else
      If CA = "+" Then tSal = Valor(rTem!Saldo) + Monto Else tSal = Valor(rTem!Saldo) - Monto
      conn.Execute "UPDATE conceptosca SET Saldo=" & tSal & " WHERE Empresa=" & rEmp!Clave & " And Clave=" & rTem!Clave
    End If
'Detalle de movimientos
    Set rCA = conn.Execute("SELECT Max(Folio) as Folio FROM camovi WHERE Empresa=" & rEmp!Clave & " And Cuenta=" & rTem!Clave)
    Mayor = Valor(rCA!Folio) + 1
    tFec = Fecha6(Fecha, 3)
    If Len(Obser) > 100 Then Obser = Mid(Obser, 1, 100)
    sql = "INSERT INTO camovi(Empresa,Cuenta,Folio,Fecha,Saldo,Cargo,Monto,Obser,Cobrado) VALUES(" & rEmp!Clave & "," _
    & rTem!Clave & "," & Mayor & "," & tFec & "," & rTem!Saldo & ",'" & CA & "'," & Monto & ",'" & Obser & "'," & Cobrado & ")"
    conn.Execute sql
  End If
End Sub

Function Fmoneda(Cad As Variant)
  If IsNull(Cad) Then Fmoneda = 0 Else Fmoneda = Format(Cad, "#,0.00")
End Function
Function CPosicion(NDoc, TipoM, Tabla, Campo, ClvC)
  Dim rTem As ADODB.Recordset
  If Val(ClvC) > 0 Then
'    If Tabla = "alumnos" Then
'      SQL2 = " And Plantel=" & ClvC
'    Else
      SQL2 = " And ClvCat=" & ClvC
      If Tabla = "facturas" Then
        If ClvC = 1 Then tSer = rEmp!SerieF Else tSer = rEmp!SerieN
        SQL2 = SQL2 & " And Serie='" & tSer & "'"
      End If
'    End If
  Else
    SQL2 = ""
  End If
  If NDoc = -2 Then
    sql = "SELECT Min(" & Campo & ") as " & Campo & " FROM " & Tabla & " WHERE Empresa=" & rEmp!Clave & SQL2
    Set rTem = conn.Execute(sql)
    If rTem.RecordCount = 0 Then CPosicion = 0 Else CPosicion = rTem.Fields(Campo)
  ElseIf NDoc = -1 Then
    sql = "SELECT Max(" & Campo & ") as " & Campo & " FROM " & Tabla & " WHERE Empresa=" & rEmp!Clave & SQL2
    Set rTem = conn.Execute(sql)
    If rTem.RecordCount = 0 Then CPosicion = 0 Else CPosicion = rTem.Fields(Campo)
  Else
    sql = "SELECT " & Campo & " FROM " & Tabla & " WHERE Empresa=" & rEmp!Clave & " And " & Campo & "=" & NDoc & SQL2
    Set rTem = conn.Execute(sql)
    If rTem.RecordCount = 0 Then
      If TipoM = ">" Then
        sql = "SELECT Min(" & Campo & ") as " & Campo & " FROM " & Tabla & " WHERE Empresa=" & rEmp!Clave & " And " & Campo & ">" & tCo & NDoc & tCo & SQL2
      ElseIf TipoM = "<" Then
        sql = "SELECT Max(" & Campo & ") as " & Campo & " FROM " & Tabla & " WHERE Empresa=" & rEmp!Clave & " And " & Campo & "<" & tCo & NDoc & tCo & SQL2
      End If
      Set rTem = conn.Execute(sql): If IsNull(rTem.Fields(Campo)) Then CPosicion = 0 Else CPosicion = Valor(rTem.Fields(Campo))
    Else
      CPosicion = rTem.Fields(Campo)
    End If
  End If
End Function
Function MostrarTextV(rs As ADODB.Recordset, DataF)
  On Error Resume Next
  nx = rs.Fields(DataF)
  If Err > 0 Then
    MostrarTextV = False
    Err = 0
  Else
    MostrarTextV = True
  End If
End Function
Sub MostrarFrm(x As Form)
  x.Left = 0: x.Top = 0
  If x.MDIChild Then
    x.Width = Menu.Width - 300
    x.Height = Menu.Height - 1760
  Else
    x.Width = Screen.Width - 380
    x.Height = Screen.Height - 1500
  End If
End Sub
Function feRC(Cad As Variant, Opc As Variant)
  If IsNull(Cad) Then
    feRC = "": Exit Function
  End If
  If Opc = 2 Then         '1 Remplaza caracateres por codigo valido 2=Cadena Normal solo quita espacios dobles
    Cad = Trim(Cad)
    Cad = Replace(Cad, "¨", ".")
    Cad = Replace(Cad, Chr(30), ".")
    Cde = ""
    Cont = 0
    For a = 1 To Len(Cad)
      Letr = Mid(Cad, a, 1)
      If Letr = " " Then
        Cont = Cont + 1
        If Cont = 1 Then Cde = Cde & " "
      Else
        Cde = Cde & Letr
        Cont = 0
      End If
    Next
    feRC = Cde
  Else
    If IsNull(Cad) Then
      feRC = ""
      Exit Function
    End If
    Cad = Trim(Cad)
    Cde = ""
    Cont = 0
    For a = 1 To Len(Cad)
      Letr = Mid(Cad, a, 1)
      Select Case Letr
        Case Is = " "
          Cont = Cont + 1
          If Cont = 1 Then Cde = Cde & " "
        Case Is = "Ñ": Cde = Cde & "&#209;": Cont = 0
        Case Is = "ñ": Cde = Cde & "&#241;": Cont = 0
        Case Is = "½": Cde = Cde & "&#189;": Cont = 0
        Case Is = "&": Cde = Cde & "&#38;": Cont = 0
        Case Is = "<": Cde = Cde & "&#60;": Cont = 0
        Case Is = ">": Cde = Cde & "&#62;": Cont = 0
        Case Is = "'": Cde = Cde & "&#39;": Cont = 0
        Case Is = "Á": Cde = Cde & "&#193;": Cont = 0
        Case Is = "É": Cde = Cde & "&#201;": Cont = 0
        Case Is = "Í": Cde = Cde & "&#205;": Cont = 0
        Case Is = "Ó": Cde = Cde & "&#211;": Cont = 0
        Case Is = "Ú": Cde = Cde & "&#218;": Cont = 0
        Case Is = "á": Cde = Cde & "&#225;": Cont = 0
        Case Is = "é": Cde = Cde & "&#233;": Cont = 0
        Case Is = "í": Cde = Cde & "&#237;": Cont = 0
        Case Is = "ó": Cde = Cde & "&#243;": Cont = 0
        Case Is = "ú": Cde = Cde & "&#250;": Cont = 0
        Case Is = "°": Cde = Cde & "&#176;": Cont = 0
        Case Is = "`": Cde = Cde & "&#96;": Cont = 0
        Case Is = "Ü": Cde = Cde & "&#220;": Cont = 0
        Case Is = "ü": Cde = Cde & "&#252;": Cont = 0
        Case Is = "´"
          Cde = Cde & "&#180;": Cont = 0
        Case Is = "¨": Cde = Cde & ".": Cont = 0
        Case Is = Chr(30): Cde = Cde & ".": Cont = 0
        Case Else
          If Asc(Letr) = 34 Then
            Cde = Cde & "&#34;"
          Else
            If InStr(wB64, Letr) > 0 Then
              Cde = Cde & Letr
            Else
              MsgBox "Caracter invalido " & Letr & "   Ascc " & Asc(Letr)
              Cde = Cde & "."
            End If
          End If
          Cont = 0
      End Select
    Next
    feRC = Cde
  End If
End Function

Sub MostrarTexN(xForm As Form, rs As ADODB.Recordset)
  Dim rCat As ADODB.Recordset
  For Each x In xForm.Controls
    If TypeName(x) = "TextBox" Then
      If Len(x.DataField) > 0 Then
        If MostrarTextV(rs, x.DataField) Then
          Info = MostrarTextV(rs, x.DataField)
          If IsNull(rs.Fields(x.DataField)) Then
            x.Text = ""
          ElseIf x.DataMember = "T" Then
            x.Text = rs.Fields(x.DataField)
          ElseIf x.DataMember = "N" Then
            If x.Tag = "" Then x.Text = rs.Fields(x.DataField) Else x.Text = Format(rs.Fields(x.DataField), x.Tag)
          ElseIf x.DataMember = "F" Then
            If x.Tag = "" Then x.Text = Format(rs.Fields(x.DataField), "dd/mm/yy") Else x.Text = Format(rs.Fields(x.DataField), x.Tag)
          ElseIf x.DataMember = "FH" Then
            If x.Tag = "" Then x.Text = Format(rs.Fields(x.DataField), "dd/mm/yy hh:mm:ss") Else x.Text = Format(rs.Fields(x.DataField), x.Tag)
          End If
        Else
          x.Text = x.DataField & " no existe"
        End If
      End If
    ElseIf TypeName(x) = "Label" Then
      If Len(x.DataField) > 2 Then
        If MostrarTextV(rs, x.DataField) Then
          If IsNull(rs.Fields(x.DataField)) Then
            x.Caption = ""
          ElseIf x.DataMember = "T" Then
            x.Caption = rs.Fields(x.DataField)
          ElseIf x.DataMember = "N" Then
            If x.Tag = "" Then x.Caption = rs.Fields(x.DataField) Else x.Caption = Format(rs.Fields(x.DataField), x.Tag)
          ElseIf x.DataMember = "F" Then
            If x.Tag = "" Then x.Caption = Format(rs.Fields(x.DataField), "dd/mm/yy") Else x.Caption = Format(rs.Fields(x.DataField), x.Tag)
          End If
        Else
          x.Caption = x.DataField & " no existe"
        End If
      End If
    ElseIf TypeName(x) = "ComboBox" Then
      If Len(x.DataField) > 0 Then
        x.WhatsThisHelpID = 0
        If MostrarTextV(rs, x.DataField) Then
          If IsNull(rs.Fields(x.DataField)) Then
            x.ListIndex = -1
          ElseIf x.Tag = "" Then
            x.ListIndex = rs.Fields(x.DataField)
          Else
            If x.DataMember = "T" Then tCam = "'" & rs.Fields(x.DataField) & "'"
            If x.DataMember = "N" Then tCam = Val(rs.Fields(x.DataField))
            sql = "SELECT Clave,Nombre FROM " & x.Tag & " WHERE Clave=" & tCam
            Set rCat = conn.Execute(sql)
            If rCat.RecordCount = 0 Then
              x.ListIndex = -1
            Else
              For a = 0 To x.ListCount - 1
                If x.List(a) = rCat!Nombre Then
                  x.ListIndex = a: x.WhatsThisHelpID = rCat!Clave: Exit For
                End If
              Next
            End If
          End If
        Else
          MsgBox "no existe el campo (" & x.DataField & ")"
        End If
      End If
    ElseIf TypeName(x) = "CheckBox" Then
      If Len(x.DataField) > 0 Then x.Value = Abs(Valor(rs.Fields(x.DataField)))
    ElseIf TypeName(x) = "Image" Or TypeName(x) = "PictureBox" Then
      If Len(x.DataField) > 0 Then
        On Error Resume Next
        Set x.DataSource = rs
        If Err <> 0 Then
          MsgBox Err & "  -  " & Error, 48, rEmp!Nombre: Err = 0
        End If
      End If
    ElseIf TypeName(x) = "OptionButton" Then
      If Len(x.Tag) > 2 Then
        If rs.Fields(x.Tag) = 0 Then
          If x.Index = 0 Then x.Value = 1
        Else
          If x.Index = 1 Then x.Value = 1
        End If
      End If
    End If
  Next
End Sub
Sub Limpiar(xForm)
  For Each x In xForm.Controls
    If TypeName(x) = "TextBox" Then
      If Len(x.DataField) > 0 Then
        If x.DataMember = "N" Then x.Text = 0 Else x.Text = ""
      End If
    ElseIf TypeName(x) = "Label" Then
      If Len(x.DataField) > 1 Then x.Caption = "" Else If x.DataField = "X" Then x.Caption = ""
    ElseIf TypeName(x) = "ComboBox" Then
      x.ListIndex = -1
    ElseIf TypeName(x) = "MSFlexGrid" Then
      x.Rows = 1
    End If
  Next
End Sub
Sub Conexion()
  ChDir App.Path
  ChDrive App.Path
  Directorio = App.Path & "\"
  DirIcon = App.Path & "\Iconos\"
  
  Open Directorio & "mysql.txt" For Input As #7
  Line Input #7, TCONN
  Close #7
  If InStr(TCONN, "SERVER=") > 0 Then
    IPMySQL = Mid(TCONN, InStr(TCONN, "SERVER=") + 7)
    IPMySQL = Mid(IPMySQL, 1, InStr(IPMySQL, ";") - 1)
  End If
  strconectar = TCONN
  
  DirRep = "Reportes\"
  Set conn = New ADODB.Connection
  conn.CursorLocation = adUseClient
  conn.ConnectionString = strconectar
  conn.Open
End Sub

Sub CentrarFrm(mio As Form)
  mio.Left = Screen.Width / 2 - mio.Width / 2
  mio.Top = Menu.Height / 2 - mio.Height / 2
  If mio.MDIChild Then mio.Top = 0 Else mio.Top = 1500
  If mio.Left < 0 Then mio.Left = 0
End Sub
Sub GE(xFo, Opc)
  If Opc = 0 Then
    For a = xFo.Grid1.TopRow To xFo.Grid1.Rows - 1
      If Not xFo.Grid1.RowIsVisible(a) Then Exit For
    Next
    If xFo.Grid1.Row >= xFo.Grid1.TopRow And xFo.Grid1.Row <= a - 2 Then
      xFo.TextE.Left = xFo.Grid1.CellLeft + xFo.Grid1.Left
      xFo.TextE.Top = xFo.Grid1.CellTop + xFo.Grid1.Top
    Else
      xFo.TextE.Visible = False
    End If
  Else
    If xFo.Grid1.Col >= xFo.Grid1.FixedCols And xFo.Grid1.Col < xFo.Grid1.Cols Then
      xFo.TextE.FontName = xFo.Grid1.FontName
      xFo.TextE.FontSize = xFo.Grid1.FontSize
      xFo.TextE.Left = xFo.Grid1.CellLeft + xFo.Grid1.Left
      xFo.TextE.Top = xFo.Grid1.CellTop + xFo.Grid1.Top
      xFo.TextE.Width = xFo.Grid1.CellWidth
      xFo.TextE.Height = xFo.Grid1.CellHeight
      xFo.TextE.backcolor = QBColor(11)
      xFo.TextE = xFo.Grid1
      xFo.TextE.Visible = True
      xFo.TextE.SelStart = 0
      xFo.TextE.SelLength = Len(xFo.TextE)
      If xFo.TextE.Visible Then xFo.TextE.SetFocus
    End If
  End If
End Sub

Sub TextG(Obj)
  Obj.SelStart = 0: Obj.SelLength = Len(Obj): Obj.backcolor = QBColor(11)
End Sub
Function TNull(Cade As Variant, Opc)
  If IsNull(Cade) Then
    If Opc <> 0 Then TNull = " " Else TNull = ""
  Else
    If Len(Cade) = 0 Then
      If Opc <> 0 Then TNull = " "
    Else
      TNull = Trim(Cade)
    End If
  End If
End Function

Sub BotonPic(x As Form)
  Dim Ct As Control
  For Each Ct In x
    Archi = ""
    If Len(Ct.Tag) > 0 Then
      If InStr(Ct.Tag, "jpg") > 0 Or InStr(Ct.Tag, "bmp") > 0 Then
        Archi = DirIcon & Ct.Tag
        If Dir(Archi) = "" Then Ct.Picture = LoadPicture() Else Ct.Picture = LoadPicture(Archi)
      End If
    ElseIf Ct.Name = "Command1" Then
      If Ct.Caption = "&Agregar" Or Ct.Caption = "&1 Agregar" Then Archi = "New.bmp"
      If Ct.Caption = "&Borrar" Or Ct.Caption = "&3 Borrar" Then Archi = "Borrar.bmp"
      If Ct.Caption = "&Editar" Or Ct.Caption = "&2 Editar" Then Archi = "Editar.bmp"
      If Ct.Caption = "&Cancelar" Then Archi = "Cancelar.bmp"
      If Ct.Caption = "&Imprimir" Or Ct.Caption = "&4 Imprimir" Then Archi = "Print.bmp"
      If Ct.Caption = "&Salir" Or Ct.Caption = "&7 Salir" Then Archi = "Salir.bmp"
      If Ct.Caption = "&Filtrar" Then Archi = "Filtrar.bmp"
      If Ct.Caption = "" Then Archi = "Billetes.bmp"
      If InStr(Ct.Caption, "Buscar") > 0 Then Archi = "Buscar.bmp"
      If Len(Archi) > 0 Then
        Archi = DirIcon & Archi
        If Dir(Archi) = "" Then Ct.Picture = LoadPicture() Else Ct.Picture = LoadPicture(Archi)
      End If
    ElseIf Ct.Name = "Command2" Then
      If Ct.Caption = "&Aceptar" Then Archi = "ok.bmp"
      If InStr(Ct.Caption, "&Borrar") > 0 Then Archi = "Borrar.bmp"
      If Ct.Caption = "&Cancelar" Then Archi = "Cancel.bmp"
      If Len(Archi) > 0 Then
        Archi = DirIcon & Archi
        If Dir(Archi) = "" Then Ct.Picture = LoadPicture() Else Ct.Picture = LoadPicture(Archi)
      End If
    ElseIf Ct.Name = "Datos" Then
      If Ct.Index = 0 Then Archi = "Pri.bmp"
      If Ct.Index = 1 Then Archi = "Izq.bmp"
      If Ct.Index = 2 Then Archi = "Der.bmp"
      If Ct.Index = 3 Then Archi = "Ult.bmp"
      Archi = DirIcon & Archi
      If Dir(Archi) = "" Then Ct.Picture = LoadPicture() Else Ct.Picture = LoadPicture(Archi)
    End If
  Next
End Sub
