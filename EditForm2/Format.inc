<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Format.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 11:00h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Rotina que altera o tamanho do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function BreakString(ByVal strValue, ByVal intSize)
  If IntSize > "" Then
    If Len(strValue) > intSize Then
      BreakString = Left(strValue, intSize) & "..."

    Else
      BreakString = strValue

    End If

  Else
    BreakString = strValue

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListSize

REM -------------------------------------------------------------------------
REM Retorna a letra correspondente ao parametro, sem acentuacoes
REM -------------------------------------------------------------------------
Private Function ChangeChar(ByVal Char)
  If InStr("áâãàa", LCase(Char)) > 0 Then
    ChangeChar = "A"

  ElseIf InStr("éêe", LCase(Char)) > 0 Then
    ChangeChar = "E"

  ElseIf InStr("íi", LCase(Char)) > 0 Then
    ChangeChar = "I"

  ElseIf InStr("óõôo", LCase(Char)) > 0 Then
    ChangeChar = "O"

  ElseIf InStr("úu", LCase(Char)) > 0 Then
    ChangeChar = "U"

  ElseIf InStr("çc", LCase(Char)) > 0 Then
    ChangeChar = "C"

  Else
    ChangeChar = UCase(Char)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ChangeChar

REM -------------------------------------------------------------------------
REM Retorna um texto de caption tratado
REM -------------------------------------------------------------------------
Private Function EditCaption(ByVal blnForm)
Dim i
  i = InStr(Session(EditCurrentField & "Caption"), "&")
  If i > 0 Then
    If blnForm Then
      EditCaption = Left(Session(EditCurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(EditCurrentField & "Caption"), i + 1 , 1) & "</U>" & Mid(Session(EditCurrentField & "Caption"), i + 2)
      FormSetAccessKey ChangeChar(Mid(Session(EditCurrentField & "Caption"), i + 1 , 1))

    Else
      EditCaption = Left(Session(EditCurrentField & "Caption"), i - 1) & Mid(Session(EditCurrentField & "Caption"), i + 1)

    End If

  Else
    EditCaption = Session(EditCurrentField & "Caption")

  End If


End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCaption

REM -------------------------------------------------------------------------
REM Retorna o proximo caracter numerico do texto a ser mascarado
REM -------------------------------------------------------------------------
Private Function MaskNextNumber(ByVal strMaskCompleter)
Dim blnOK, strChar

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If (Left(sstrText3, 1) >= "0") And (Left(sstrText3, 1) <= "9") Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextNumber = strMaskCompleter

  Else
    MaskNextNumber = strChar

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function MaskNextNumber

REM -------------------------------------------------------------------------
REM Retorna o proximo caracter alfa do texto a ser mascarado
REM -------------------------------------------------------------------------
Private Function MaskNextAlfa(ByVal strMaskCompleter)
Dim blnOK

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If ((Left(sstrText3, 1) >= "a") And (Left(sstrText3, 1) <= "z")) Or ((Left(sstrText3, 1) >= "A") And (Left(sstrText3, 1) <= "Z")) Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextAlfa = strMaskCompleter

  Else
    MaskNextAlfa = strChar

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function MaskNextAlfa

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo do form
REM -------------------------------------------------------------------------
Public Function PutMask(ByVal strText, ByVal strMask, ByVal strMaskCompleter)
Dim strAux, i

  If strMask > "" Then
    strAux = ""
    sstrText3 = CStr(ZTIReverse(strText))
    For i = Len(strMask) to 1 Step -1
REM          sobjShow.MessageError Mid(strMask, i , 1)
      Select Case Mid(strMask, i , 1)
        Case "0"
          strAux = strAux & MaskNextNumber(strMaskCompleter)

        Case "A"
          strAux = strAux & MaskNextAlfa(strMaskCompleter)

        Case Else
          strAux = strAux & Mid(strMask, i , 1)

      End Select
    Next
    PutMask = ZTIReverse(strAux)

  Else
    PutMask = strText

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function PutMask

REM -------------------------------------------------------------------------
REM Formata uma variavel de acordo com seu tipo
REM -------------------------------------------------------------------------
Public Function FormatVariable(ByVal Value, ByVal strType)
REM  sobjShow.JSAlert Value & "&" & strType
  If Value > "" Then
    If LCase(Typename(Value)) <> "string" Then
      Select Case strType
        Case conDataInt, conDataSmallInt, conDataTinyInt
          FormatVariable = ZTIInt(Value)
  
        Case conDataMoney             
          FormatVariable = ZTIMoney(Value)
  
        Case conDataFloat, conDataReal
          FormatVariable = ZTINumber(Value)
  
        Case conDataDateTime
          If LCase(TypeName(Value)) = "date" Then
            FormatVariable = ZTILongDate(Value)
            
          Else
            If IsValidDate(Value) Then
  REM        sobjShow.JSAlert Value
              FormatVariable = FormatDate(Value)
  REM          sobjShow.JSAlert Value
  
            Else
              FormatVariable = Value
          
            End If
          End If
  
        Case Else
          FormatVariable = Value

      End Select
   
    Else
      FormatVariable = Trim(Value)
   
   End If
  Else
    FormatVariable = Null

  End If
End Function
REM -------------------------------------------------------------------------
REM Final da Function FormatVariable

REM -------------------------------------------------------------------------
REM Formata uma variavel de acordo com seu tipo para ser usada numa chamada SQL
REM -------------------------------------------------------------------------
Public Function FormatForSQL(ByVal Value, ByVal strType)
REM  sobjShow.JSAlert Value & "&" & strType
  If Value > "" Then
    Select Case LCase(strType)
      Case conDataInt, conDataSmallInt, conDataTinyInt
        FormatForSQL = NormalizeString(Value)

      Case conDataVarChar, conDataChar, conDataText
        FormatForSQL = SQLString(Value)

      Case conDataDateTime
        FormatForSQL = SQLDateTime(Value)

      Case Else
        FormatForSQL = Value

    End Select

  Else
    FormatForSQL = Null

  End if
End Function
REM -------------------------------------------------------------------------
REM Final da Function FormatForSQL

REM -------------------------------------------------------------------------
REM Prepara Money para serem utilizadas dentro de uma Query SQL
REM -------------------------------------------------------------------------
Public Function ImportDouble(ByVal strValue)
dim target
dim i
dim cOneChar
dim intPos, monAux

  If Trim(strValue) > "" Then
  	target = ""
  	strValue = UCase(Trim(strValue))

  	For i = 1 to Len(strValue)
  		cOneChar = Mid(strValue, i, 1)
  		If cOneChar = "," Or cOneChar = "-" Or (cOneChar >= "0" and cOneChar <= "9") then
        target = target & cOneChar

  		End if

  	Next

    strValue = target

    intPos = InStr(strValue, ",")

    monAux = 0

    If intPos > 0 Then
      monAux = CLng(Left(strValue, intPos - 1))
      monAux = monAux + (CLng(Mid(strValue, intPos + 1)) / Exp((Len(strValue) - intPos)*Log(10)))

    Else
      monAux = CLng(strValue)

    End If

    ImportDouble = monAux

 Else
    ImportDouble = Null

 End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function ImportDouble

REM -------------------------------------------------------------------------
REM Testa uma Data
REM -------------------------------------------------------------------------
Function FormatDate(ByVal strDate)
Dim Dia, Mes, Ano, Date
  
  strDate = NormalizeString(CStr(strDate))
  Mes = Month(Now)
  Ano = Year(Now)
  If Len(strDate) > 1  Then
    Dia = CInt(Left(strDate, 2))
    If Len(strDate) > 2 Then
      Mes = CInt(Mid(strDate, 3, 2))
  
    End If
  
    If Len(strDate) > 4 Then
      Ano = CInt(Mid(strDate, 5, 4))
  
    End If
  
    If Ano < 100 Then
      If Ano + 1900 > Year(Now) - 70 Then
        Ano = Ano + 1900
  
      Else
        Ano = Ano + 2000
  
      End If        
    End If
  End If
  FormatDate = LeadingZeroes(Dia, 2) & "/" & LeadingZeroes(Mes, 2) & "/" & LeadingZeroes(Ano, 4)

End Function
REM ---------------------------------------------------------------------
REM Fim do FormatDate

REM -------------------------------------------------------------------------
REM Formata uma variavel
REM -------------------------------------------------------------------------
Public Function EditFormatField(ByVal Value, ByVal strType, ByVal intSize, ByVal intCharCase)
  If Value > "" Then
    Select Case LCase(strType)
      Case conDataInt, conDataSmallInt, conDataTinyInt
REM        sobjShow.JSAlert Value & TotalNormalizeString(Value) & CLng(TotalNormalizeString(Value))
        If LCase(TypeName(Value)) = "string" Then
          EditFormatField = CLng(TotalNormalizeString(Value))
        
        Else
          EditFormatField = Value
        
        End If

      Case conDataDateTime
        If TypeName(Value) = "date" Then
          EditFormatField = ZTILongDate(Value)
          
        Else
          If IsValidDate(Value) Then
REM        sobjShow.JSAlert Value
            Value = FormatDate(Value)
REM          sobjShow.JSAlert Value
        
          End If
        End If
        EditFormatField = Value

      Case conDataVarChar, conDataChar, conDataText
        Select Case intCharCase
          Case conLCase
            EditFormatField = LCase(Left(Trim(CStr(Value)), intSize))
            
          Case conUCase
            EditFormatField = UCase(Left(Trim(CStr(Value)), intSize))
            
          Case Else
            EditFormatField = Left(Trim(CStr(Value)), intSize)
            
        End Select

      Case conDataMoney, conDataFloat, conDataReal
        Select Case LCase(TypeName(Value))
          Case "currency", "double", "single", "decimal"
            EditFormatField = Value
        
          Case Else
REM            sobjShow.JSAlert TypeName(Value) & Value
            EditFormatField = ImportDouble(Value)
REM            sobjShow.JSAlert TypeName(ImportDouble(Value)) & ImportDouble(Value)
        
        End Select
  
      Case Else
        EditFormatField = Value

      End Select

  Else
    EditFormatField = Null

  End If
End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormatField

REM -------------------------------------------------------------------------
REM Formata uma variavel
REM -------------------------------------------------------------------------
Public Function EditFormatFieldtoSave(ByVal Value, ByVal strType, ByVal intSize, ByVal intCharCase)
  If Value > "" Then
    Select Case LCase(strType)
      Case conDataInt, conDataSmallInt, conDataTinyInt
REM        sobjShow.JSAlert Value 
REM        sobjShow.JSAlert TotalNormalizeString(Value)
REM        sobjShow.JSAlert CLng(TotalNormalizeString(Value))
        If LCase(TypeName(Value)) = "string" Then
          EditFormatFieldtoSave = ToInteger(TotalNormalizeString(Value))
        
        Else
          EditFormatFieldtoSave = Value
        
        End If

      Case conDataDateTime   
        If TypeName(Value) = "date" Then
          EditFormatFieldtoSave = Value
            
        Else
          EditFormatFieldtoSave = ImportDate(Value, conDateTypeDDMMAAAA)
          
        End If

      Case conDataVarChar, conDataChar, conDataText
        Select Case intCharCase
          Case conLCase
            EditFormatFieldtoSave = LCase(Left(CStr(Value), intSize))
            
          Case conUCase
            EditFormatFieldtoSave = UCase(Left(CStr(Value), intSize))
            
          Case Else
            EditFormatFieldtoSave = Left(CStr(Value), intSize)
            
        End Select

      Case conDataMoney, conDataFloat, conDataReal
        If LCase(TypeName(Value)) = "string" Then
          EditFormatFieldtoSave = ImportDouble(Value)
        
        Else
          EditFormatFieldtoSave = Value
        
        End If

      Case Else
        EditFormatFieldtoSave = Value

      End Select

  Else
    EditFormatFieldtoSave = Null

  End If
End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormatField

REM -------------------------------------------------------------------------
REM Retorna o valor 1 se o parametro passado for verdadeiro
REM -------------------------------------------------------------------------
Private Function ToString(ByVal strText)

  If Not IsNull(strText) Then
    ToString = CStr(strText)

  Else
    ToString = ""

  End if

End Function
REM -------------------------------------------------------------------------
REM Final da Function BitBln

REM -------------------------------------------------------------------------
REM Retorna o valor 1 se o parametro passado for verdadeiro
REM -------------------------------------------------------------------------
Private Function ToInteger(ByVal strText)
  If Not IsNull(strText) And strText > "" Then
    ToInteger = CLng(strText)

  Else
    ToInteger = 0

  End if

End Function
REM -------------------------------------------------------------------------
REM Final da Function BitBln

%>