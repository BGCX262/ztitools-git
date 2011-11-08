<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Edit.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Edicao
REM  Cria‡ao     : 11:23h 7/2/1998
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Constantes do EDIT
REM -------------------------------------------------------------------------
Const conFormInputText     = "text"
Const conFormInputHidden   = "hidden"
Const conFormInputCheckBox = "checkbox"
Const conFormInputRadio    = "radio"
Const conFormInputTextArea = "textarea"
Const conFormInputSubmit   = "submit"
Const conFormInputReset    = "reset"
Const conFormInputPassword = "password"
Const conFormInputFile     = "file"

REM -------------------------------------------------------------------------
REM Globais do EDIT
REM -------------------------------------------------------------------------
  REM -----------------------------------------------------------------------
  REM Globais das Funcoes dos formularios
  REM -----------------------------------------------------------------------
  Dim sstrFormSetEncType

  Dim sFormInputonClick  
  Dim sFormInputonChange
  Dim sFormAccessKey  
  Dim sFormAlt
  Dim sFormTitle
  Dim sFormReadOnly
  Dim sFormDisabled
  Dim sFormClass
  Dim sFormMaxLength
  REM -----------------------------------------------------------------------
  REM Globais do formulario
  REM -----------------------------------------------------------------------
  Dim sblnEditFormState

  Dim sintEditFormAction
  Dim sintEditFormMethod
  Dim sstrEditFormName
  Dim sstrEditFormTarget
  Dim sstrEditFormOnSubmit
  Dim sFormId

  REM -----------------------------------------------------------------------
  REM Globais dos Inputs
  REM -----------------------------------------------------------------------
  Dim sintEditCounter

  Dim sstrEditType(50)
  Dim sstrEditText(50)
  Dim sstrEditName(50)
  Dim sintEditSize(50)
  Dim sstrEditValue(50)
  Dim sintEditMaxLength(50)
  Dim sstrEditAlign(50)
  Dim sblnEditState(50)
  Dim sstrEditSRC(50)

  Dim sblnEditTableState

  Dim sblnEditTableBegin(50)
  Dim sblnEditTableEnd(50)

  Dim sblnEditTableRowBegin(50)
  Dim sblnEditTableRowEnd(50)

  Dim sblnFormSetNoCrLf
REM -------------------------------------------------------------------------
REM Retorna o contador do form
REM -------------------------------------------------------------------------
Function EditCounter
  EditCounter = sintEditCounter

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCounter

REM -------------------------------------------------------------------------
REM Soma um ao contador
REM -------------------------------------------------------------------------
Sub EditCounterPlus
  sintEditCounter = sintEditCounter + 1

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCounterPlus

REM -------------------------------------------------------------------------
REM Inicia a automacao de uma entrada de dados via formulario
REM -------------------------------------------------------------------------
Sub EditBegin(strAction, strMethod, strName, strTarget, strOnSubmit)

  sblnEditFormState    = False
 
  sintEditFormAction   = strAction
  sintEditFormMethod   = strMethod
  sstrEditFormName     = strName
  sstrEditFormTarget   = strTarget
  sstrEditFormOnSubmit = strOnSubmit

  sintEditCounter = 1

  sblnEditTableBegin(EditCounter)    = False
  sblnEditTableEnd(EditCounter)      = False

  sblnEditTableRowBegin(EditCounter) = False
  sblnEditTableRowEnd(EditCounter)   = False

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditBegin

REM -------------------------------------------------------------------------
REM Finaliza a automacao de uma entrada de dados
REM -------------------------------------------------------------------------
Sub EditEnd
  sblnEditFormState = True

  Session("sblnEditFormState")     = sblnEditFormState
  
  Session("sintEditFormAction")    = sintEditFormAction
  Session("sintEditFormMethod")    = sintEditFormMethod
  Session("sstrEditFormTarget")    = sstrEditFormTarget
  Session("sstrEditFormOnSubmit")  = sstrEditFormOnSubmit

  Session("sintEditCounter")       = sintEditCounter

  Session("sstrEditType")          = sstrEditType
  Session("sstrEditText")          = sstrEditText
  Session("sstrEditName")          = sstrEditName
  Session("sintEditSize")          = sintEditSize
  Session("sstrEditValue")         = sstrEditValue
  Session("sintEditMaxLength")     = sintEditMaxLength
  Session("sstrEditAlign")         = sstrEditAlign
  Session("sblnEditState")         = sblnEditState
  Session("sstrEditSRC")           = sstrEditSRC

  Session("sblnEditTableState")    = sblnEditTableState

  Session("sblnEditTableBegin")    = sblnEditTableBegin
  Session("sblnEditTableEnd")      = sblnEditTableEnd

  Session("sblnEditTableRowBegin") = sblnEditTableRowBegin
  Session("sblnEditTableRowEnd")   = sblnEditTableRowEnd
  
  Session("sstrEditFormName")      = LCase(sstrEditFormName)

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditEnd

REM -------------------------------------------------------------------------
REM Finaliza a automacao de uma entrada de dados
REM -------------------------------------------------------------------------
Function EditGet(strForm)
Dim blnResult

  blnResult = False

  If Session("sstrEditFormName") = LCase(strForm) Then
    sblnEditFormState     = Session("sblnEditFormState")    
                                                            
    sintEditFormAction    = Session("sintEditFormAction")   
    sintEditFormMethod    = Session("sintEditFormMethod")   
    sstrEditFormName      = Session("sstrEditFormName")     
    sstrEditFormTarget    = Session("sstrEditFormTarget")   
    sstrEditFormOnSubmit  = Session("sstrEditFormOnSubmit") 
                                                            
    sintEditCounter       = Session("sintEditCounter")      
                                                            
    sstrEditType          = Session("sstrEditType")         
    sstrEditText          = Session("sstrEditText")           
    sstrEditName          = Session("sstrEditName")         
    sintEditSize          = Session("sintEditSize")         
    sstrEditValue         = Session("sstrEditValue")        
    sintEditMaxLength     = Session("sintEditMaxLength")    
    sstrEditAlign         = Session("sstrEditAlign")        
    sblnEditState         = Session("sblnEditState")        
    sstrEditSRC           = Session("sstrEditSRC")          
                                                            
    sblnEditTableState    = Session("sblnEditTableState")   
                                                            
    sblnEditTableBegin    = Session("sblnEditTableBegin")   
    sblnEditTableEnd      = Session("sblnEditTableEnd")     
                                                            
    sblnEditTableRowBegin = Session("sblnEditTableRowBegin")
    sblnEditTableRowEnd   = Session("sblnEditTableRowEnd")  

    blnResult = True

  End If

  EditGet = blnResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditGet

REM -------------------------------------------------------------------------
REM Inicio da tabela dentro do EDIT
REM -------------------------------------------------------------------------
Sub EditTableBegin

  sblnEditTableBegin(EditCounter) = True

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTableBegin

REM -------------------------------------------------------------------------
REM Fim da tabela dentro do EDIT
REM -------------------------------------------------------------------------
Sub EditTableEnd

  sblnEditTableEnd(EditCounter - 1) = True
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTableEnd

REM -------------------------------------------------------------------------
REM Inicio de uma linha dentro da tabela dentro do EDIT
REM -------------------------------------------------------------------------
Sub EditTableRowBegin

  sblnEditTableRowBegin(EditCounter) = True

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTableRowBegin

REM -------------------------------------------------------------------------
REM Fim de uma linha dentro da tabela dentro do EDIT
REM -------------------------------------------------------------------------
Sub EditTableRowEnd

  sblnEditTableRowEnd(EditCounter - 1) = True

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTableRowEnd

REM -------------------------------------------------------------------------
REM Monta a estrutura de formularios do EDIT
REM -------------------------------------------------------------------------
Sub EditMountForm
  Dim i
  Dim blnTableState
  
  blnTableState = False

  FormBegin sintEditFormAction, sintEditFormMethod, sstrEditFormName, sstrEditFormTarget, sstrEditFormOnSubmit
  
  If sblnEditFormState Then
    For i = 1 To EditCounter - 1
      If sblnEditTableBegin(i) Then
        ShowHTMLCR "<TABLE BORDER=0>"
       
      End If

      If sblnEditTableRowBegin(i) Then
        Response.Write "<TR>"
        blnTableState = True

      End If

      If blnTableState Then
        Response.Write "<TD>"

      End If

      If sstrEditText(i) > "" Then
        ShowHTMLCR "<b>" & sstrEditText(i) & "</b> "
        
      End If

      FormInput sstrEditType(i), sstrEditName(i), sintEditSize(i), sstrEditValue(i), sintEditMaxLength(i), sstrEditAlign(i), sblnEditState(i), sstrEditSRC(i)

      If blnTableState Then
        Response.Write "</TD>"

      End If
         
      If sblnEditTableRowEnd(i) Then
        ShowHTMLCR "</TR>"

        blnTableState = False

      End If

      If sblnEditTableEnd(i) Then
        ShowHTMLCR "</TABLE>"

      End If

    Next
  End If

  FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditMountForm

REM -------------------------------------------------------------------------
REM Monta a estrutura de formularios do EDIT Script Java para validacao
REM -------------------------------------------------------------------------
Sub EditMountFormScript

  If sblnEditState Then
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditMountForm

REM -------------------------------------------------------------------------
REM Adiciona um item dentro do form em andamento
REM -------------------------------------------------------------------------
Sub EditAdd(intType, strText, strName, intSize, strValue, intMaxLength, strAlign, blnState, strSRC)

  sstrEditType(EditCounter)      = intType
  sstrEditText(EditCounter)      = strText
  sstrEditName(EditCounter)      = strName
  sintEditSize(EditCounter)      = intSize
  sstrEditValue(EditCounter)     = strValue
  sintEditMaxLength(EditCounter) = intMaxLength
  sstrEditAlign(EditCounter)     = strAlign
  sblnEditState(EditCounter)     = blnState
  sstrEditSRC(EditCounter)       = strSRC

  EditCounterPlus

  sblnEditTableBegin(EditCounter)    = Null
  sblnEditTableEnd(EditCounter)      = Null

  sblnEditTableRowBegin(EditCounter) = Null
  sblnEditTableRowEnd(EditCounter)   = Null

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAdd

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo Texto
REM -------------------------------------------------------------------------
Sub EditAddText(strText, strName, intSize, strValue)

  EditAdd conFormInputText, strText, strName, intSize, strValue, 0, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddText

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo Texto com o tamanho maximo definido
REM -------------------------------------------------------------------------
Sub EditAddTextMaxLength(strText, strName, intSize, strValue, intMaxLength)

  EditAdd conFormInputText, strText, strName, intSize, strValue, intMaxLength, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddTextMaxLength

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo Texto
REM -------------------------------------------------------------------------
Sub EditAddTextArea(strText, strName, intSize, strValue, intMaxLength, strAlign)

  EditAdd conFormInputTextArea, strText, strName, intSize, strValue, intMaxLength, strAlign, Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddText

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo Invisivel
REM -------------------------------------------------------------------------
Sub EditAddHidden(strText, strName, strValue)

  EditAdd conFormInputHidden, strText, strName, 0, strValue, 0, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddHidden

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo Invisivel
REM -------------------------------------------------------------------------
Sub EditAddPassword(strText, strName, intSize)

  EditAdd conFormInputPassword, strText, strName, intSize, "", 0, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddHidden

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo CheckBox
REM -------------------------------------------------------------------------
Sub EditAddCheckBox(strText, strName, strValue, strAlign, intState)
    
  EditAdd conFormInputCheckBox, strText, strName, 0, strValue, 0, strAlign, intState, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddCheckBox

REM -------------------------------------------------------------------------
REM Adiciona no EDIT Campo CheckBox
REM -------------------------------------------------------------------------
Sub EditAddRadio(strText, strName, strValue, strAlign, intState)

  EditAdd conFormInputRadio, strText, strName, 0, strValue, 0, strAlign, intState, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddCheckBox

REM -------------------------------------------------------------------------
REM Adiciona no EDIT o botao Submit
REM -------------------------------------------------------------------------
Sub EditAddSubmit(strValue, strName)

  EditAdd conFormInputSubmit, "", strName, 0, strValue, 0, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddSubmit

REM -------------------------------------------------------------------------
REM Adiciona no EDIT o botao Submit Image
REM -------------------------------------------------------------------------
Sub EditAddSubmitImage(strValue, strName, strImage)

  EditAdd conFormInputSubmit, "", strName, 0, strValue, 0, "", Null, strImage
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddSubmit

REM -------------------------------------------------------------------------
REM
REM -------------------------------------------------------------------------
Sub EditAddReset(strValue)
    
  EditAdd conFormInputReset, "", "", 0, strValue, 0, "", Null, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub

REM -------------------------------------------------------------------------
REM
REM -------------------------------------------------------------------------
Sub FormSetNoCrLf

  sblnFormSetNoCrLf = True

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub

REM -------------------------------------------------------------------------
REM
REM -------------------------------------------------------------------------
Sub FormClearNoCrLf

  sblnFormSetNoCrLf = False

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub


REM -------------------------------------------------------------------------
REM Inicio do Form
REM -------------------------------------------------------------------------
Sub FormBegin(txtAction, txtMethod, txtName, txtTarget, txtOnSubmit)
	Dim txtHTML

	txtHTML = "<form action=" & Chr(34) & txtAction & Chr(34)

	If txtMethod > "" Then
		txtHTML = txtHTML & " method=" & Chr(34) & txtMethod & Chr(34)

	End If

	If txtName > "" Then
		txtHTML = txtHTML & " name=" & Chr(34) & txtName & Chr(34)

	End If

	If txtTarget > "" Then
		txtHTML = txtHTML & " target=" & Chr(34) & txtTarget & Chr(34)

	End If

	If txtOnSubmit > "" Then
		txtHTML = txtHTML & " onsubmit=" & Chr(34) & txtOnSubmit & Chr(34)

  End If
  
  If sstrFormSetEncType > "" Then
    txtHTML = txtHTML & " ENCTYPE=""" & sstrFormSetEncType & """"
    
  End If

	ShowHTMLCR txtHTML & ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormBegin

REM -------------------------------------------------------------------------
REM Fim do Form
REM -------------------------------------------------------------------------
Sub FormEnd

	ShowHTMLCR "</form>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormEnd

REM -------------------------------------------------------------------------
REM Denine o Enctype
REM -------------------------------------------------------------------------
Sub FormSetEncType(ByVal strType)
   
  sstrFormSetEncType = strType

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetEncType

REM -------------------------------------------------------------------------
REM Entra Check Box
REM -------------------------------------------------------------------------
Sub FormInputCheckBox(txtName, txtValue, txtAlign, numState)
  
	FormInput "checkbox", txtName, 0, txtValue, 0, txtAlign, numState, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputCheckBox

REM -------------------------------------------------------------------------
REM Entra Radio Buttom
REM -------------------------------------------------------------------------
Sub FormInputRadio(ByVal txtName, ByVal txtValue, ByVal txtAlign, ByVal numState)
    
	FormInput "radio", txtName, 0, txtValue, 0, txtAlign, numState, ""
 
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputRadio

REM -------------------------------------------------------------------------
REM Entra botao Grafico
REM -------------------------------------------------------------------------
Sub FormInputImage(txtName, txtValue, txtImage, strAlt)
	FormSetAlt strAlt
	FormSetNoCrLf
	FormInput "image", txtName, 0, txtValue, 0, "", -1, txtImage
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputImage

REM -------------------------------------------------------------------------
REM Entra Text Area
REM -------------------------------------------------------------------------
Sub FormTextArea(ByVal txtName, ByVal intRows, ByVal intCols, ByVal txtContent)
   
  Response.Write "<TEXTAREA"
  
  If txtName > "" Then
    Response.Write " NAME=" & txtName
  
  End If

  If intRows > "" Then
    Response.Write " ROWS=" & intRows
    
  End If

  If intCols > "" Then
    Response.Write " COLS=" & intCols
    
  End If

  
	If sFormAccessKey > "" Then
		Response.Write " accesskey=""" & sFormAccessKey & """"
   
	End If
  
	If sFormInputonChange > "" Then
		Response.Write " onchange=""" & sFormInputonChange & """"
    
	End If
  
	If sFormTitle > "" Then
		Response.Write " title=""" & sFormTitle & """"
   
  End If         

	If sFormReadOnly > "" Then
		Response.Write " readonly"
   
  End If         
  
	If sFormDisabled > "" Then
		Response.Write " disabled"
    
  End If         

	If sFormClass > "" Then
		Response.Write " class=" & sFormClass
    
  End If         

  sFormTitle     = ""
  sFormAccessKey = ""

  ShowHTMLCR ">"

  If txtContent > "" Then
    ShowHTMLCR txtContent
    
  End If
  
  ShowHTMLCR "</TEXTAREA>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputTextArea

REM -------------------------------------------------------------------------
REM Inicio do Select
REM -------------------------------------------------------------------------
Sub FormSelectBegin(ByVal txtName, ByVal intSize, ByVal blnMultiple)
  
  Response.Write "<SELECT name=" & txtName
  
  If intSize > "" Then
    Response.Write " SIZE=" & intSize
    
  End If

  If blnMultiple Then
    Response.Write " MULTIPLE"
    
  End If
  
	If sFormAccessKey > "" Then
		Response.Write " accesskey=""" & sFormAccessKey & """"

	End If
  
	If sFormTitle > "" Then
		Response.Write " title=""" & sFormTitle & """"

  End If         

	If sFormReadOnly > "" Then
		Response.Write " readonly"

  End If         

	If sFormDisabled > "" Then
		Response.Write " disabled"

  End If         

  If sFormInputonChange > "" Then
		Response.Write " onchange=""" & sFormInputonChange & """"

	End If

	If sFormClass > "" Then
		Response.Write " class=" & sFormClass

  End If         

  sFormTitle     = ""
  sFormAccessKey = ""

  ShowHTMLCR ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectBegin

REM -------------------------------------------------------------------------
REM Entra Select
REM -------------------------------------------------------------------------
Sub FormSelectEnd

  ShowHTMLCR "</SELECT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectBegin

Sub FormId(strId)
  sFormId = strId

End Sub

REM -------------------------------------------------------------------------
REM Entra Select
REM -------------------------------------------------------------------------
Sub FormSelectOption(ByVal txtValue, ByVal txtText, ByVal blnSelect)

  Response.Write "<OPTION value=""" & txtValue & """"
  
  If blnSelect Then
    Response.Write " SELECTED"
    
  End If
  
  If sFormId > "" Then
    Response.Write " ID=""" & sFormId & """"
    
  End If
  
  ShowHTMLCR ">" & txtText

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectBegin

REM -------------------------------------------------------------------------
REM Executa um Select com as UF                
REM -------------------------------------------------------------------------
Public Sub FormSelectMonthName(ByVal strName, ByVal intSize, ByVal blnMultiple, ByVal intMonth)

  If intMonth = "" Then
    intMonth = Month(Now())
  
  End If

  FormSelectBegin strName, intSize, blnMultiple
  FormSelectChooseOption 1, "Janeiro",   intMonth
  FormSelectChooseOption 2, "Fevereiro", intMonth
  FormSelectChooseOption 3, "Março",     intMonth
  FormSelectChooseOption 4, "Abril",     intMonth
  FormSelectChooseOption 5, "Maio",      intMonth
  FormSelectChooseOption 6, "Junho",     intMonth
  FormSelectChooseOption 7, "Julho",     intMonth
  FormSelectChooseOption 8, "Agosto",    intMonth
  FormSelectChooseOption 9, "Setembro",  intMonth
  FormSelectChooseOption 10, "Outubro",  intMonth
  FormSelectChooseOption 11, "Novembro", intMonth
  FormSelectChooseOption 12, "Dezembro", intMonth
  FormSelectEnd                                 
                                              
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectMonthName

REM -------------------------------------------------------------------------
REM Executa um Select com os dias de um Mes
REM -------------------------------------------------------------------------
Public Sub FormSelectMonthDay(ByVal strName, ByVal intSize, ByVal blnMultiple, ByVal intDay)
  Dim i

  If intDay = "" Then
    intDay = Month(Now())
  
  End If

  FormSelectBegin strName, intSize, blnMultiple

  For i = 1 To 31
    FormSelectChooseOption i, "Dezembro", intDay

  Next

  FormSelectEnd                                 
                                               
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectMonthDay

REM -------------------------------------------------------------------------
REM Executa um Select com os dias da Semana
REM -------------------------------------------------------------------------
Public Sub FormSelectWeekDay(ByVal strName, ByVal intSize, ByVal blnMultiple, ByVal intWeekDay)
           
  If intWeekDay = "" Then
    intWeekDay = WeekDay(Now())
  
  End If

  FormSelectBegin txtName, intSize, blnMultiple
  FormSelectChooseOption 1, "Domingo",       intWeekDay
  FormSelectChooseOption 2, "Segunda-Feira", intWeekDay
  FormSelectChooseOption 3, "Terça-Feira",   intWeekDay
  FormSelectChooseOption 4, "Quarta-Feira",  intWeekDay
  FormSelectChooseOption 5, "Quinta-Feira",  intWeekDay
  FormSelectChooseOption 6, "Sexta-Feira",   intWeekDay
  FormSelectChooseOption 7, "Sábado",        intWeekDay
  FormSelectEnd                                 
                                               
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSelectWeekDay

REM -------------------------------------------------------------------------
REM Executa um Select com as UF                
REM -------------------------------------------------------------------------
Public Sub FormSelectUF(ByVal strName, ByVal intSize, ByVal blnMultiple, ByVal strUF)

  FormSelectBegin strName, intSize, blnMultiple
  If strUF = "" Then
    FormSelectChooseOption "", "", ""
  
  End If
  FormSelectChooseOption "AC", "Acre",                strUF
  FormSelectChooseOption "AP", "Amapá",               strUF
  FormSelectChooseOption "AM", "Amazonas",            strUF
  FormSelectChooseOption "BA", "Bahia",               strUF
  FormSelectChooseOption "CE", "Ceará",               strUF
  FormSelectChooseOption "DF", "Distrito Federal",    strUF
  FormSelectChooseOption "ES", "Espirito Santo",      strUF
  FormSelectChooseOption "GO", "Goiás",               strUF
  FormSelectChooseOption "MA", "Maranhão",            strUF
  FormSelectChooseOption "MT", "Mato Grosso",         strUF
  FormSelectChooseOption "MS", "Mato Grosso do Sul",  strUF
  FormSelectChooseOption "MG", "Minas Gerais",        strUF
  FormSelectChooseOption "PA", "Pará",                strUF
  FormSelectChooseOption "PB", "Paraiba",             strUF
  FormSelectChooseOption "PR", "Paraná",              strUF
  FormSelectChooseOption "PE", "Pernanbuco",          strUF
  FormSelectChooseOption "PI", "Piauí",               strUF
  FormSelectChooseOption "RJ", "Rio de Janeiro",      strUF
  FormSelectChooseOption "RN", "Rio Grande do Norte", strUF 
  FormSelectChooseOption "RS", "Rio Grande do Sul",   strUF
  FormSelectChooseOption "RO", "Rondônia",            strUF
  FormSelectChooseOption "RR", "Roraima",             strUF
  FormSelectChooseOption "SC", "Santa Catarina",      strUF
  FormSelectChooseOption "SP", "São Paulo",           strUF
  FormSelectChooseOption "SE", "Sergipe",             strUF
  FormSelectChooseOption "TO", "Tocantins",           strUF
  FormSelectEnd                                 
                                               
End Sub                                        
REM -------------------------------------------------------------------------
REM Fim da Sub FormSelectUF

REM -------------------------------------------------------------------------
REM Executa um Select com os Paises                
REM -------------------------------------------------------------------------
Public Sub FormSelectPais(ByVal strName, ByVal intSize, ByVal blnMultiple, ByVal strPais)
 
  FormSelectBegin strName, intSize, blnMultiple
  If strPais = "" Then
    FormSelectChooseOption "", "", ""
  
  End If
  FormSelectChooseOption "AF", "Afeganistão",                            strPais
  FormSelectChooseOption "AL", "Albânia",                                strPais
  FormSelectChooseOption "DZ", "Argélia",                                strPais
  FormSelectChooseOption "AS", "Samoa americano",                        strPais
  FormSelectChooseOption "AD", "Andorra",                                strPais
  FormSelectChooseOption "AO", "Angola",                                 strPais
  FormSelectChooseOption "AI", "Anguilla",                               strPais
  FormSelectChooseOption "AQ", "Antártica",                              strPais
  FormSelectChooseOption "AG", "Antígua",                                strPais
  FormSelectChooseOption "AR", "Argentina",                              strPais
  FormSelectChooseOption "AM", "Armênia",                                strPais
  FormSelectChooseOption "AW", "Aruba",                                  strPais
  FormSelectChooseOption "AU", "Austrália",                              strPais
  FormSelectChooseOption "AT", "Áustria",                                strPais
  FormSelectChooseOption "AZ", "Azerbeidjã",                             strPais
  FormSelectChooseOption "BS", "Bahamas",                                strPais
  FormSelectChooseOption "BH", "Bahrain",                                strPais
  FormSelectChooseOption "BD", "Bangladesh",                             strPais
  FormSelectChooseOption "BB", "Barbados",                               strPais
  FormSelectChooseOption "BE", "Bélgica",                                strPais
  FormSelectChooseOption "BZ", "Belize",                                 strPais
  FormSelectChooseOption "BJ", "Benin",                                  strPais
  FormSelectChooseOption "BM", "Ilhas Bermudas",                         strPais
  FormSelectChooseOption "BT", "Butão",                                  strPais
  FormSelectChooseOption "BO", "Bolívia",                                strPais
  FormSelectChooseOption "BA", "Bosnia e Herzegovina",                   strPais
  FormSelectChooseOption "BW", "Botsuana",                               strPais
  FormSelectChooseOption "BV", "Ilha de Bouvet",                         strPais
  FormSelectChooseOption "BR", "Brasil",                                 strPais
  FormSelectChooseOption "IO", "Território britânico no Oceano Índico",  strPais
  FormSelectChooseOption "VG", "Ilhas de Virgem britânicas",             strPais
  FormSelectChooseOption "BN", "Brunei",                                 strPais
  FormSelectChooseOption "BG", "Bulgária",                               strPais
  FormSelectChooseOption "BF", "Burkina Faso",                           strPais
  FormSelectChooseOption "BU", "Birmânia",                               strPais
  FormSelectChooseOption "BI", "Burundi",                                strPais
  FormSelectChooseOption "BY", "Byelorussian SSR",                       strPais
  FormSelectChooseOption "CA", "Canadá",                                 strPais
  FormSelectChooseOption "CT", "Cantão e Ilhas de Enderbury",            strPais
  FormSelectChooseOption "CV", "Cabo Verde",                             strPais
  FormSelectChooseOption "KY", "Ilhas de caimão",                        strPais
  FormSelectChooseOption "CF", "República africana central",             strPais
  FormSelectChooseOption "TD", "Chade",                                  strPais
  FormSelectChooseOption "CL", "Chile",                                  strPais
  FormSelectChooseOption "CN", "China",                                  strPais
  FormSelectChooseOption "CX", "Ilha de Natal",                          strPais
  FormSelectChooseOption "CC", "Ilhas de Cocos",                         strPais
  FormSelectChooseOption "CO", "Colômbia",                               strPais
  FormSelectChooseOption "KM", "Comores",                                strPais
  FormSelectChooseOption "CG", "Congo",                                  strPais
  FormSelectChooseOption "CK", "Cozinhe Ilhas",                          strPais
  FormSelectChooseOption "CR", "Costa Rica",                             strPais
  FormSelectChooseOption "HR", "Croácia (Hrvatska)",                     strPais
  FormSelectChooseOption "CU", "Cuba",                                   strPais
  FormSelectChooseOption "CY", "Chipre",                                 strPais
  FormSelectChooseOption "CS", "Tchecoslováquia",                        strPais
  FormSelectChooseOption "KH", "Campuchea democrático",                  strPais
  FormSelectChooseOption "YD", "Iêmen democrático",                      strPais
  FormSelectChooseOption "DK", "Dinamarca",                              strPais
  FormSelectChooseOption "DJ", "Djibouti",                               strPais
  FormSelectChooseOption "DM", "Dominica",                               strPais
  FormSelectChooseOption "DO", "República dominicana",                   strPais
  FormSelectChooseOption "NQ", "Dronning Maud Terra",                    strPais
  FormSelectChooseOption "TP", "Timor oriental",                         strPais
  FormSelectChooseOption "EC", "Equador",                                strPais
  FormSelectChooseOption "EG", "Egito",                                  strPais
  FormSelectChooseOption "SV", "El Salvador",                            strPais
  FormSelectChooseOption "GQ", "Guiné equatorial",                       strPais
  FormSelectChooseOption "ER", "Eritrea",                                strPais
  FormSelectChooseOption "EE", "Estônia",                                strPais
  FormSelectChooseOption "ET", "Etiópia",                                strPais
  FormSelectChooseOption "FO", "Ilhas de Faeroe",                        strPais
  FormSelectChooseOption "FK", "Ilhas de Falkland",                      strPais
  FormSelectChooseOption "FJ", "Fiji",                                   strPais
  FormSelectChooseOption "FI", "Finlândia",                              strPais
  FormSelectChooseOption "FR", "França",                                 strPais
  FormSelectChooseOption "FX", "França, Metropolita,",                   strPais
  FormSelectChooseOption "GF", "Guiana francesa",                        strPais
  FormSelectChooseOption "PF", "Polinésia francesa",                     strPais
  FormSelectChooseOption "TF", "Territórios Meridionais franceses",      strPais
  FormSelectChooseOption "GA", "Gabão",                                  strPais
  FormSelectChooseOption "GM", "Gâmbia",                                 strPais
  FormSelectChooseOption "GE", "Geórgia",                                strPais
  FormSelectChooseOption "DE", "Alemanha",                               strPais
  FormSelectChooseOption "GH", "Gana",                                   strPais
  FormSelectChooseOption "GI", "Gibraltar",                              strPais
  FormSelectChooseOption "GR", "Grécia",                                 strPais
  FormSelectChooseOption "GL", "Greenland",                              strPais
  FormSelectChooseOption "GD", "Granada",                                strPais
  FormSelectChooseOption "GP", "Guadalupe",                              strPais
  FormSelectChooseOption "GU", "Guam",                                   strPais
  FormSelectChooseOption "GT", "Guatemala",                              strPais
  FormSelectChooseOption "GN", "Guiné",                                  strPais
  FormSelectChooseOption "GW", "Guiné-bisseu",                           strPais
  FormSelectChooseOption "GY", "Guiana",                                 strPais
  FormSelectChooseOption "HT", "Haiti",                                  strPais
  FormSelectChooseOption "HM", "Ilhas Heard e Mc Donald",                strPais
  FormSelectChooseOption "HN", "Honduras",                               strPais
  FormSelectChooseOption "HK", "Hong Kong",                              strPais
  FormSelectChooseOption "HU", "Hungria",                                strPais
  FormSelectChooseOption "IS", "Islândia",                               strPais
  FormSelectChooseOption "IN", "Índia",                                  strPais
  FormSelectChooseOption "ID", "Indonésia",                              strPais
  FormSelectChooseOption "IR", "Irã",                                    strPais
  FormSelectChooseOption "IQ", "Iraque",                                 strPais
  FormSelectChooseOption "IE", "Irlanda",                                strPais
  FormSelectChooseOption "IL", "Israel",                                 strPais
  FormSelectChooseOption "IT", "Itália",                                 strPais
  FormSelectChooseOption "CI", "Costa do marfim",                        strPais
  FormSelectChooseOption "JM", "Jamaica",                                strPais
  FormSelectChooseOption "JP", "Japão",                                  strPais
  FormSelectChooseOption "JT", "Ilha de Johnston",                       strPais
  FormSelectChooseOption "JO", "Jordan",                                 strPais
  FormSelectChooseOption "KZ", "Kazakstan",                              strPais
  FormSelectChooseOption "KE", "Quênia",                                 strPais
  FormSelectChooseOption "KI", "Quiribati",                              strPais
  FormSelectChooseOption "KW", "Kuwait",                                 strPais
  FormSelectChooseOption "KG", "Kyrgyzstan",                             strPais
  FormSelectChooseOption "LA", "República Democrática de Lao",           strPais
  FormSelectChooseOption "LV", "Latvia",                                 strPais
  FormSelectChooseOption "LB", "Líbano",                                 strPais
  FormSelectChooseOption "LS", "Lesoto",                                 strPais
  FormSelectChooseOption "LR", "Libéria",                                strPais
  FormSelectChooseOption "LY", "Líbia Jamahiria árabe",                  strPais
  FormSelectChooseOption "LI", "Liechtenstein",                          strPais
  FormSelectChooseOption "LT", "Lituânia",                               strPais
  FormSelectChooseOption "LU", "Luxemburgo",                             strPais
  FormSelectChooseOption "MO", "Macau",                                  strPais
  FormSelectChooseOption "MK", "Macedônia",                              strPais
  FormSelectChooseOption "MG", "Madagáscar",                             strPais
  FormSelectChooseOption "MY", "Malásia",                                strPais
  FormSelectChooseOption "MW", "Malauí",                                 strPais
  FormSelectChooseOption "MV", "Maldivas",                               strPais
  FormSelectChooseOption "ML", "Mali",                                   strPais
  FormSelectChooseOption "MT", "Malta",                                  strPais
  FormSelectChooseOption "MQ", "Martinica",                              strPais
  FormSelectChooseOption "MR", "Mauritânia",                             strPais
  FormSelectChooseOption "MU", "Maurício",                               strPais
  FormSelectChooseOption "MX", "México",                                 strPais
  FormSelectChooseOption "MI", "Ilhas Midway",                           strPais
  FormSelectChooseOption "MC", "Mônaco",                                 strPais
  FormSelectChooseOption "MN", "Mongólia",                               strPais
  FormSelectChooseOption "MS", "Montserrat",                             strPais
  FormSelectChooseOption "MA", "Marrocos",                               strPais
  FormSelectChooseOption "MZ", "Moçambique",                             strPais
  FormSelectChooseOption "NA", "Namíbia",                                strPais
  FormSelectChooseOption "NP", "Nepal",                                  strPais
  FormSelectChooseOption "NR", "Nauru",                                  strPais
  FormSelectChooseOption "NL", "Países Baixos",                          strPais
  FormSelectChooseOption "AN", "Países Baixos Antilhas",                 strPais
  FormSelectChooseOption "NT", "Zona neutra",                            strPais
  FormSelectChooseOption "NC", "Calidônia novo",                         strPais
  FormSelectChooseOption "NZ", "Nova Zelândia",                          strPais
  FormSelectChooseOption "NI", "Nicarágua",                              strPais
  FormSelectChooseOption "NE", "Níger",                                  strPais
  FormSelectChooseOption "NG", "Nigéria",                                strPais
  FormSelectChooseOption "NU", "Niue",                                   strPais
  FormSelectChooseOption "NF", "Ilha de Norfolk",                        strPais
  FormSelectChooseOption "MP", "Ilhas Mariana do norte",                 strPais
  FormSelectChooseOption "NO", "Noruega",                                strPais
  FormSelectChooseOption "OM", "Omã",                                    strPais
  FormSelectChooseOption "PC", "Ilhas Pacíficas",                        strPais
  FormSelectChooseOption "PK", "Paquistão",                              strPais
  FormSelectChooseOption "PA", "Panamá",                                 strPais
  FormSelectChooseOption "PG", "Papua-Nova Guiné",                       strPais
  FormSelectChooseOption "PY", "Paraguai",                               strPais
  FormSelectChooseOption "KP", "República da Coréia",                    strPais
  FormSelectChooseOption "PE", "Peru",                                   strPais
  FormSelectChooseOption "PH", "Filipinas",                              strPais
  FormSelectChooseOption "PN", "Ilha de Pitcairn",                       strPais
  FormSelectChooseOption "PL", "Polônia",                                strPais
  FormSelectChooseOption "PT", "Portugal",                               strPais
  FormSelectChooseOption "TW", "Província da China Taiwan",              strPais
  FormSelectChooseOption "PR", "Porto Rico",                             strPais
  FormSelectChooseOption "QA", "Qatar",                                  strPais
  FormSelectChooseOption "KR", "República da Coréia",                    strPais
  FormSelectChooseOption "RE", "Reunião",                                strPais
  FormSelectChooseOption "RO", "Romênia",                                strPais
  FormSelectChooseOption "RU", "Federação russa",                        strPais
  FormSelectChooseOption "RW", "Ruanda",                                 strPais
  FormSelectChooseOption "GS", "Ilhas S. Geórgia e S. Intercale",        strPais
  FormSelectChooseOption "LC", "São Lúcia",                              strPais
  FormSelectChooseOption "VC", "São Vincent/Grenadines",                 strPais
  FormSelectChooseOption "WS", "Samoa",                                  strPais
  FormSelectChooseOption "SM", "San Marino",                             strPais
  FormSelectChooseOption "ST", "São Tomé e Principe",                    strPais
  FormSelectChooseOption "SA", "Arábia Saudita",                         strPais
  FormSelectChooseOption "SN", "Senegal",                                strPais
  FormSelectChooseOption "SC", "Seicheles",                              strPais
  FormSelectChooseOption "SL", "Sierra Leone",                           strPais
  FormSelectChooseOption "SG", "Cingapura",                              strPais
  FormSelectChooseOption "SB", "Ilhas de Salomão",                       strPais
  FormSelectChooseOption "SO", "Somália",                                strPais
  FormSelectChooseOption "ZA", "África do Sul",                          strPais
  FormSelectChooseOption "ES", "Espanha",                                strPais
  FormSelectChooseOption "LK", "Sri Lanka",                              strPais
  FormSelectChooseOption "SH", "St. Helena",                             strPais
  FormSelectChooseOption "KN", "St. Kitts Nevis Anguilla",               strPais
  FormSelectChooseOption "PM", "St. Pierre e Miquelon",                  strPais
  FormSelectChooseOption "SD", "Sudão",                                  strPais
  FormSelectChooseOption "SR", "Suriname",                               strPais
  FormSelectChooseOption "SJ", "Ilhas Svalbard e Jan Mayen",             strPais
  FormSelectChooseOption "SZ", "Suazilândia",                            strPais
  FormSelectChooseOption "SE", "Suécia",                                 strPais
  FormSelectChooseOption "CH", "Suíça",                                  strPais
  FormSelectChooseOption "SY", "Syran República árabe",                  strPais
  FormSelectChooseOption "TH", "Tailândia",                              strPais
  FormSelectChooseOption "TG", "Togo",                                   strPais
  FormSelectChooseOption "TK", "Tokelau",                                strPais
  FormSelectChooseOption "TO", "Tonga",                                  strPais
  FormSelectChooseOption "TT", "Trinidad e Tobago",                      strPais
  FormSelectChooseOption "TN", "Tunísia",                                strPais
  FormSelectChooseOption "TR", "Turquia",                                strPais
  FormSelectChooseOption "TM", "Turkmenistan",                           strPais
  FormSelectChooseOption "TC", "Turcos e Ilhas de Caicos",               strPais
  FormSelectChooseOption "TV", "Tuvalu",                                 strPais
  FormSelectChooseOption "UM", "EUA Ilhas Periféricas Secundárias",      strPais
  FormSelectChooseOption "UG", "Uganda",                                 strPais
  FormSelectChooseOption "UA", "SSR ucraniano",                          strPais
  FormSelectChooseOption "AE", "Emirados árabes unidos",                 strPais
  FormSelectChooseOption "GB", "Reino Unido",                            strPais
  FormSelectChooseOption "CM", "República unida de Cameroon",            strPais
  FormSelectChooseOption "TZ", "República unida de Tanzânia",            strPais
  FormSelectChooseOption "US", "Estados Unidos",                         strPais
  FormSelectChooseOption "PU", "Estados Unidos Ilhas do Pacífico",       strPais
  FormSelectChooseOption "VI", "Estados Unidos Ilhas Virgem",            strPais
  FormSelectChooseOption "HV", "Volta superior",                         strPais
  FormSelectChooseOption "UY", "Uruguai",                                strPais
  FormSelectChooseOption "UZ", "Uzbekistan",                             strPais
  FormSelectChooseOption "VU", "Vanuatu",                                strPais
  FormSelectChooseOption "VA", "Vaticano",                               strPais
  FormSelectChooseOption "VE", "Venezuela",                              strPais
  FormSelectChooseOption "VN", "Vietnã",                                 strPais
  FormSelectChooseOption "WK", "Ilha Wake",                              strPais
  FormSelectChooseOption "WF", "Wallis e Ilhas de Futuma",               strPais
  FormSelectChooseOption "EH", "Saara ocidental",                        strPais
  FormSelectChooseOption "YE", "Iêmen",                                  strPais
  FormSelectChooseOption "YU", "Iugoslávia",                             strPais
  FormSelectChooseOption "ZR", "Zaire",                                  strPais
  FormSelectChooseOption "ZM", "Zâmbia",                                 strPais
  FormSelectChooseOption "ZW", "Zimbabwee",                              strPais
  FormSelectEnd                                 
                                                                            
End Sub                                                                     
REM -------------------------------------------------------------------------
REM Fim da Sub FormSelectPais                                               
                                                                            
REM -------------------------------------------------------------------------
REM Executa um Select com as UF                                             
REM -------------------------------------------------------------------------
Public Sub FormSelectChooseOption(ByVal strOption, ByVal strDescription, ByVal strOptionFix)
                                                                            
  FormSelectOption strOption, strDescription, strOption = strOptionFix      
                                                                            
End Sub                                                                     
REM -------------------------------------------------------------------------
REM Fim da Sub FormSelectEstados                                            

REM -------------------------------------------------------------------------
REM Entra Botao Reset
REM -------------------------------------------------------------------------
Sub FormInputReset(txtValue)
    
	FormInput "reset", "", 0, txtValue, 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputReset

REM -------------------------------------------------------------------------
REM Entra Botao Submit
REM -------------------------------------------------------------------------
Sub FormInputSubmit(txtName, txtValue)

	FormInput "submit", txtName, 0, txtValue, 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputSubmit

REM -------------------------------------------------------------------------
REM Entra Botao Submit com Grafico
REM -------------------------------------------------------------------------
Sub FormInputSubmitImage(txtName, txtValue, txtImage)

	FormInput "image", txtName, 0, txtValue, 0, "", -1, txtImage

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputSubmitImage

REM -------------------------------------------------------------------------
REM Entra Campo Senha
REM -------------------------------------------------------------------------
Sub FormInputPassword(txtName, numSize)

	FormInput "password", txtName, numSize, "", 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputPassword

REM -------------------------------------------------------------------------
REM Entra Campo Texto
REM -------------------------------------------------------------------------
Sub FormInputText(txtName, numSize, txtValue)

	FormInput "text", txtName, numSize, txtValue, 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputText

REM -------------------------------------------------------------------------
REM Entra Campo File para download de arquivos
REM -------------------------------------------------------------------------
Sub FormInputFile(ByVal strName)

	FormInput "File", strName, 0, "", 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputText

REM -------------------------------------------------------------------------
REM Entra Campo Texto com tamanho maximo
REM -------------------------------------------------------------------------
Sub FormInputTextMaxLength(txtName, numSize, txtValue, numMaxLength)

	FormInput "text", txtName, numSize, txtValue, numMaxLength, "", -1, ""
                                                                                              
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputTextMaxLength

REM -------------------------------------------------------------------------
REM Entra Campo Invisivel
REM -------------------------------------------------------------------------
Sub FormInputHidden(txtName, txtValue)

	FormInput "hidden", txtName, 0, txtValue, 0, "", -1, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputHidden

REM -------------------------------------------------------------------------
REM Entra InputDefault
REM -------------------------------------------------------------------------
Sub FormInput(strType, strName, intSize, strValue, intMaxLength, strAlign, intState, strSRC)
	Dim strHTML

	strHTML = "<input type=""" & strType & """"

	If strName > "" Then
		strHTML = strHTML & " name=""" & strName & """"

	End If

	If intSize > 0 Then
		strHTML = strHTML & " size=""" & intSize & """"

	End If

	If strValue > "" Then
		 strHTML = strHTML & " value=""" & strValue & """"

	End If

	If intMaxLength > 0 Then
		strHTML = strHTML & " maxlength=""" & intMaxLength & """"

	ElseIf sFormMaxLength > "" Then
		strHTML = strHTML & " maxlength=""" & sFormMaxLength & """"
	
	End If
  
  If strAlign > "" Then
		strHTML = strHTML & " align=""" & strAlign & """"

	End If

	If strSRC > "" Then
		strHTML = strHTML & " src=""" & strSRC & """"

	End If

	If sFormInputonClick > "" Then
		strHTML = strHTML & " onclick=""" & sFormInputonClick & """"

	End If
  
	If sFormInputonChange > "" Then
		strHTML = strHTML & " onchange=""" & sFormInputonChange & """"

	End If
  
	If sFormAccessKey > "" Then
		strHTML = strHTML & " accesskey=""" & sFormAccessKey & """"

	End If

	If sFormAlt > "" Then
		strHTML = strHTML & " alt=""" & sFormAlt & """"

  End If     
  
	If sFormTitle > "" Then
		strHTML = strHTML & " title=""" & sFormTitle & """"

  End If         

	If sFormReadOnly > "" Then
		strHTML = strHTML & " readonly"

  End If         

	If sFormDisabled > "" Then
		strHTML = strHTML & " disabled"

  End If         

	If sFormClass > "" Then
		strHTML = strHTML & " class=" & sFormClass

  End If         

  sFormAlt       = ""
  sFormTitle     = ""
  sFormAccessKey = ""

  If intState = 1 Then
    strHTML = strHTML & " checked"

  End If  

  If sblnFormSetNoCrLf Then
	  Response.Write strHTML & ">"
  
  Else
	  ShowHTMLCR strHTML & ">"
	  
	End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInput

REM -------------------------------------------------------------------------
REM Altera o alt de um input do tipo image
REM -------------------------------------------------------------------------
Sub FormSetAlt(ByVal strAlt)

	sFormAlt = strAlt

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetAlt

REM -------------------------------------------------------------------------
REM Altera o Max Length dos forms
REM -------------------------------------------------------------------------
Sub FormSetMaxLength(ByVal strMaxLength)

	sFormMaxLength = strMaxLength

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetMaxLength

REM -------------------------------------------------------------------------
REM Altera o title de um input
REM -------------------------------------------------------------------------
Sub FormSetTitle(ByVal strTitle)

	sFormTitle = strTitle

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetTitle

REM -------------------------------------------------------------------------
REM Altera o disabled de um input
REM -------------------------------------------------------------------------
Sub FormSetDisabled(ByVal blnDisabled)
  
  If blnDisabled = "1" Then
    sFormDisabled = "1"
    
  Else
    sFormDisabled = ""
    
  End If  

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetDisabled

REM -------------------------------------------------------------------------
REM Altera o readonly de um input
REM -------------------------------------------------------------------------
Sub FormSetReadOnly(ByVal blnReadOnly)
  
  If blnReadOnly = "1" Then
    sFormReadOnly = "1"
    
  Else
    sFormReadOnly = ""
    
  End If  

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetReadOnly

REM -------------------------------------------------------------------------
REM Altera a tecla de atalho de um input
REM -------------------------------------------------------------------------
Sub FormSetAccessKey(ByVal strText)

	sFormAccessKey = strText

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetAccessKey

REM -------------------------------------------------------------------------
REM Altera a tecla de atalho de um input
REM -------------------------------------------------------------------------
Sub FormSetClass(ByVal strText)

	sFormClass = strText

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormSetAccessKey

REM -------------------------------------------------------------------------
REM Altera a chamada do onClick de um input
REM -------------------------------------------------------------------------
Sub FormInputSetOnClick(ByVal strText)
 
	sFormInputonClick = strText

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputSetOnClick

REM -------------------------------------------------------------------------
REM Altera a chamada do onChange de um input
REM -------------------------------------------------------------------------
Sub FormInputSetOnChange(ByVal strText)

	sFormInputonChange = strText

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputSetOnChange

REM -------------------------------------------------------------------------
REM Entra Botao Submit - Function
REM -------------------------------------------------------------------------
Function FunctionFormInputSubmit(txtName, txtValue)

	FunctionFormInputSubmit = FunctionFormInput("submit", txtName, 0, txtValue, 0, "", -1)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function FunctionFormInputSubmit

REM -------------------------------------------------------------------------
REM Entra InputDefault - Function
REM -------------------------------------------------------------------------
Function FunctionFormInput(strType, strName, inrSize, strValue, intMaxLength, strAlign, intState)
	Dim strHTML

	strHTML = "<input type=""" & strType & """"

	If strName > "" Then
		strHTML = strHTML & " name=""" & strName & """"

	End If
  
	If intSize > 0 Then
		strHTML = strHTML & " size=""" & intSize & """"

	End If

	If strValue > "" Then
		 strHTML = strHTML & " value=""" & strValue & """"

	End If

	If intMaxLength > 0 Then
		strHTML = strHTML & " maxlength=""" & intMaxLength & """"

	End If

	If strAlign > "" Then
		strHTML = strHTML & " align=""" & strAlign & """"

	End If

	If strSRC > "" Then
		strHTML = strHTML & " src=""" & strSRC & """"

	End If

	If Not intState = Null Then
		strHTML = strHTML & "checked"

	End If

	FunctionFormInput = txtHTML & ">"

End Function
REM -------------------------------------------------------------------------
REM Final da Function FormInput

REM -------------------------------------------------------------------------
REM Entra FormJavaBegin
REM -------------------------------------------------------------------------
Sub FormJavaBegin(ByVal strFunctionName, ByVal strFormName)

  ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
REM  ShowHTMLCR "<!" & "--"
  ShowHTMLCR "function " & strFunctionName & "() {"

  Session("FormJavaForm") = strFormName

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaBegin

REM -------------------------------------------------------------------------
REM Entra FormJavaEnd
REM -------------------------------------------------------------------------
Sub FormJavaEnd

  ShowHTMLCR "return true;"

  ShowHTMLCR "}"
REM  ShowHTMLCR "//-->"
  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaEnd

REM -------------------------------------------------------------------------
REM Entra FormJavaRequired
REM -------------------------------------------------------------------------
Sub FormJavaRequired(ByVal strField, ByVal strMessage)
  
  ShowHTMLCR "if (document." & Session("FormJavaForm") & "." & strField & ".value.length == 0) {"
  ShowHTMLCR "alert(""" & strMessage & """);"
  ShowHTMLCR "document." & Session("FormJavaForm") & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaRequired

REM -------------------------------------------------------------------------
REM Entra FormJavaCompare
REM -------------------------------------------------------------------------
Sub FormJavaCompare(ByVal strField,  ByVal strField2, ByVal strMessage)

  ShowHTMLCR "if (document." & Session("FormJavaForm") & "." & strField & ".value != document." & Session("FormJavaForm") & "." & strField2 & ".value) {"
  ShowHTMLCR "alert(""" & strMessage & """);"
  ShowHTMLCR "document." & Session("FormJavaForm") & "." & strField2 & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaCompare                

REM -------------------------------------------------------------------------
REM Entra FormJavaClearField
REM -------------------------------------------------------------------------
Private Sub FormJavaClearField(ByVal strField)
  ShowHTMLCR "document." & Session("FormJavaForm") & "." & strField & ".value = """""           
 
End Sub
REM -------------------------------------------------------------------------
REM Final da sub FormJavaClearField

REM -------------------------------------------------------------------------
REM Entra FormJavaRequired
REM -------------------------------------------------------------------------
Sub FormJavaFocus(ByVal strField)
  
  ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
REM  ShowHTMLCR "<!" & "--"
  ShowHTMLCR "{"
  ShowHTMLCR "document." & Session("FormJavaForm") & "." & strField & ".focus()"
  ShowHTMLCR "}"
REM  ShowHTMLCR "//-->"
  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaRequired

REM -------------------------------------------------------------------------
REM Fim do Edit.inc
%>
