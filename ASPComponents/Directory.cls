VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

Dim gScriptContext As ScriptingContext

Dim gsFileName As String, _
    gsFullPath As String, _
    gbPathType As Boolean
    
Dim gsComments As String, _
    gsTitle As String, _
    gsVersion As String, _
    gsCompanyName As String, _
    gsProductName As String, _
    gsLegalCopyright As String

Dim gRequest As Request, _
    gResponse As Response, _
    gServer As Server, _
    gSession As Session, _
    gApplication As Application
    
    
Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Set gScriptContext = ScriptContext
  
  Set gRequest = gScriptContext.Request
  Set gResponse = gScriptContext.Response
  
End Sub

Private Sub Class_Initialize()

    gsComments = App.Comments
    gsTitle = App.Title
    gsVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    gsCompanyName = App.CompanyName
    gsProductName = App.ProductName

End Sub

Public Property Get Comments() As String
    
    Comments = gsComments
    
End Property

Public Property Get Title() As String
    
    Title = gsTitle
    
End Property

Public Property Get Version() As String
    
    Version = gsVersion
    
End Property
Public Property Get CompanyName() As String
    
    CompanyName = gsCompanyName
    
End Property

Public Property Get ProductName() As String
    
    ProductName = gsProductName
    
End Property

Public Property Get LegalCopyright() As String
    
    LegalCopyright = gsLegalCopyright
    
End Property

Public Function SetAttribute(ByVal sFileName As String, nAttributes As Integer)
Rem vbNormal    0   Normal
Rem vbReadOnly  1   Read-only
Rem vbHidden    2   Hidden
Rem vbSystem    4   System
Rem vbDirectory 16  Directory or folder
Rem vbArchive   32  File has changed since last backup
    
    SetAttr sFileName, nAttributes

End Function

Public Function GetAttribute(ByVal sFileName As String)
Rem vbNormal    0   Normal
Rem vbReadOnly  1   Read-only
Rem vbHidden    2   Hidden
Rem vbSystem    4   System
Rem vbDirectory 16  Directory or folder
Rem vbArchive   32  File has changed since last backup
    
    GetAttribute = GetAttr(sFileName)

End Function

Public Property Let FileName(sFileName As String)
    
    gsFileName = sFileName
    
End Property

Public Property Get FileName() As String
    
    FileName = gsFileName
    
End Property

Public Property Let FullPath(sFullPath As String)
    
    gsFullPath = sFullPath
    
End Property

Public Property Get FullPath() As String
    
    FullPath = gsFullPath
    
End Property

Public Property Let PathType(bPathType As Boolean)
    
    gbPathType = bPathType
    
End Property

Public Property Get PathType() As Boolean
   
    PathType = gbPathType
    
End Property

Public Function FileDT(ByVal sFileName As String) As Date

    FileDT = FileDateTime(sFileName)
    
End Function

Public Function LastDate(ByVal sFileName As String) As String

Dim nDay As Integer, _
    nYear As Integer, _
    nHour As Integer, _
    nMinute As Integer, _
    sLastDate As String
    
Dim dLastDate As Date
    
Dim sMonth As String, _
    sWeekDay As String
  
  dLastDate = FileDateTime(sFileName)
  
  nDay = Day(dLastDate)
  nYear = Year(dLastDate)
  
  sMonth = ZTIMonth(dLastDate)
  sWeekDay = ZTIWeekDay(dLastDate)
  
  nHour = Hour(dLastDate)
  nMinute = Minute(dLastDate)
        
  sLastDate = sWeekDay & ", " & nDay & " de "
  sLastDate = sLastDate & sMonth & " de " & nYear & " às "
  sLastDate = sLastDate & Format(nHour, "00")
  sLastDate = sLastDate & ":" & Format(nMinute, "00") & "h"
  
  LastDate = sLastDate

End Function

Public Function ZTIWeekDay(ByVal dDate As Date) As String

Rem vbSunday    1
Rem vbMonday    2
Rem vbTuesday   3
Rem vbWednesday 4
Rem vbThursday  5
Rem vbFriday    6
Rem vbSaturday  7

Dim aWeekDay As Variant, _
    nWeekDay As Integer, _
    sWeekDay As String
    
    
  aWeekDay = Array("Domingo", _
                   "Segunda", _
                   "Terça", _
                   "Quarta", _
                   "Quinta", _
                   "Sexta", _
                   "Sábado")
                   
  nWeekDay = WeekDay(dDate)
  
  If nWeekDay > 1 And nWeekDay < 7 Then
    sWeekDay = sWeekDay & "-Feira"
    
  End If
  
  ZTIWeekDay = aWeekDay(nWeekDay)
  
End Function

Public Function ZTIMonth(ByVal dDate As Date) As String

Dim aMonth As Variant

  aMonth = Array("Janeiro", _
           "Fevereiro", _
           "Março", _
           "Abril", _
           "Maio", _
           "Junho", _
           "Julho", _
           "Agosto", _
           "Setembro", _
           "Outubro", _
           "Novembro", _
           "Dezembro")

  ZTIMonth = aMonth(Month(dDate) - 1)
  
End Function

Public Function ConvertChar(nAction, sLine)
Rem ----------------------------------------------------------------------------
Rem  Converte Letras do padrao 850 para o utilizado no HTML
Rem ----------------------------------------------------------------------------
Rem  ISO-8859 - ISO International Standard --Information Processing
Rem  8-bit Single-Byte Coded Graphic
Rem ----------------------------------------------------------------------------
Rem  ISO-8859-1:1987 Coded Character Set Part 1: Latin Alphabet No. 1
Rem ----------------------------------------------------------------------------
Rem Private Type CharTypes
Rem   Normal As String * 1
Rem   Windows As String * 1
Rem   CP850 As String * 1
Rem   HTML As String * 10
Rem End Type

Const ZTIWindowsHTML As Integer = 1
Const ZTIWindowsNormal As Integer = 2
Const ZTIWindowsLatin As Integer = 3
Const ZTINormalHTML As Integer = 4

Dim aNormal As Variant, _
    aWindows As Variant, _
    aLatin As Variant, _
    aHTML As Variant, _
    aFrom As Variant, _
    aTo As Variant
    
Dim nI As Integer, _
    nPosition As Integer, _
    nLastPosition As Integer, _
    sResult As String

  aNormal = Array( _
             "A", "A", "A", "A", "A", _
             "a", "a", "a", "a", "a", _
             "C", "C", _
             "E", "E", "E", "E", _
             "e", "e", "e", "e", _
             "I", "I", "I", "I", _
             "i", "i", "i", "i", _
             "N", "n", _
             "O", "O", "O", "O", "O", _
             "o", "o", "o", "o", "o", _
             "U", "U", "U", "U", _
             "u", "u", "u", "u", _
             "<", ">")

  aHTML = Array( _
             "Agrave", "Aacute", "Acirc", "Auml", "Atilde", _
             "agrave", "aacute", "acirc", "auml", "atilde", _
             "Ccirc", "ccirc", _
             "Egrave", "Eacute", "Ecirc", "Euml", _
             "egrave", "eacute", "ecirc", "euml", _
             "Igrave", "Iacute", "Icirc", "Iuml", _
             "igrave", "iacute", "icirc", "iuml", _
             "Ntilde", "ntilde", _
             "Ograve", "Oacute", "Ocirc", "Ouml", "Otilde", _
             "ograve", "oacute", "ocirc", "ouml", "otilde", _
             "Ugrave", "Uacute", "Ucirc", "Uuml", _
             "ugrave", "uacute", "ucirc", "uuml", _
             "Yacute", _
             "yacute", "yuml", _
             "gt", "lt")

  aWindows = Array( _
             "À", "Á", "Â", "Ä", "Ã", _
             "à", "á", "â", "ä", "ã", _
             "Ç", "ç", _
             "È", "É", "Ê", "Ë", _
             "è", "é", "ê", "ë", _
             "Ì", "Í", "Î", "Ï", _
             "ì", "í", "î", "ï", _
             "Ñ", "ñ", _
             "Ò", "Ó", "Ô", "Ö", "Õ", _
             "ò", "ó", "ô", "ö", "õ", _
             "Ù", "Ú", "Û", "Ü", _
             "ù", "ú", "û", "ü", _
             "Ý", _
             "ý", "ÿ", _
             "<", ">")
             
  aLatin = Array( _
           "192", "193", "194", "195", "196", _
           "225", "226", "227", "228", "229", _
           "199", "231", _
           "200", "201", "202", "203", _
           "232", "233", "234", "235", _
           "204", "205", "206", "207", _
           "236", "237", "238", "239", _
           "209", "241", _
           "210", "211", "212", "213", "214", _
           "242", "243", "244", "245", "246", _
           "217", "219", "220", _
           "249", "250", "251", _
           "221", _
           "253", "254", _
           "", "")

  Select Case nAction
  Case ZTIWindowsHTML
    aFrom = aWindows
    aTo = aHTML
  
  Case ZTIWindowsNormal
    aFrom = aWindows
    aTo = aNormal
  
  Case ZTIWindowsLatin
    aFrom = aWindows
    aTo = aLatin
  
  Case ZTINormalHTML
    aFrom = aNormal
    aTo = aHTML
    
  Case Else
  End Select
  
  For nI = 1 To 66
    nPosition = InStr(aFrom(nI), sLine)
    sResult = ""
        
    nLastPosition = nPosition
    
    While nPosition > 0 And Not nLastPosition = nPosition
    
      nLastPosition = nPosition

      sResult = Left(sLine, nPosition - 1) & _
                 aTo(nI) & _
                 Mid(sLine, nPosition + 1, Len(sLine))
                 
      nPosition = InStr(nPosition, aFrom(nI), sLine)
      
    Wend
  Next

  ConvertChar = sResult

End Function

Public Function GetImage(ByVal strSource As String) As String
    Dim Buffer As String
    Dim Result As String
    
    Open strSource For Binary As #1
    Buffer = String(1024, " ")
    Result = ""
    
    Do
        Get #1, , Buffer
        Result = Result & Buffer
        
    Loop While Not EOF(1)
    
    Close 1, 2
    
    GetImage = Result
    
End Function

