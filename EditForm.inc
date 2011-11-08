<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 10:00h 16/12/1998
REM  Local       : Brasilia/DF                         
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------
%>
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/DataPars.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/FormPars.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Test.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Current.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Format.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/String.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/DataBase.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Form.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/List.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/URL.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Validate.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Summary.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Option.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm2/Others.inc"-->
<%

  Const conDataFloat          = "float"
  Const conDataReal           = "real"
  Const conDataVarChar        = "varchar"
  Const conDataChar           = "char"
  Const conDataText           = "text"
  Const conDataInt            = "int"
  Const conDataDateTime       = "datetime"
  Const conDataMoney          = "money"
  Const conDataTinyInt        = "tinyint"
  Const conDataSmallInt       = "smallint"
  Const conDataBit            = "bit"

  Const conDataNull           = True
  Const conDataNotNull        = False

  Const conRelation1to1       = 0
  Const conRelation1toN       = 1

  Const conConditionDisable   = True
  Const conConditionEnable    = False

  Const conOptionCascade      = 0
  Const conOptionRestricted   = 1
  Const conOptionNullifies    = 2

  Const conNCase              = 0
  Const conUCase              = 1
  Const conLCase              = 2

  Const conClientValidation   = True
  Const conServerValidation   = False

  Const conMoveFirst          = "0"
  Const conMovePrevious       = "1"
  Const conMoveNext           = "2"
  Const conMoveLast           = "3"

  Const conTextField          = 0
  Const conUFField            = 1
  Const conLookupField        = 2
  Const conSeparateDateField  = 3
  Const conCheckField         = 4
  Const conTextAreaField      = 5
  Const conRadioField         = 6
  Const conSelectField        = 7
  Const conPasswordField      = 8
  Const conHTTPField          = 9
  Const conEMailField         = 10
  Const conImageField         = 11
  Const conImageFileField     = 12
  Const conFileField          = 13
  Const conAtualizationField  = 14

  Const conValidNone          = 0
  Const conValidCGC           = 1
  Const conValidCPF           = 2
  Const conValidDate          = 3
  Const conValidSepDate       = 4
  Const conValidDateMToday    = 5
  Const conValidSepDateMToday = 6
  Const conValidTime          = 7
  Const conValidDateTime      = 8
  Const conValidEmail         = 9
  Const conValidCompareDates  = 10

  Const conPEditOption        = "EO"
  Const conPEditWhat          = "EW"
  Const conPEditMove          = "EM"
  Const conPEditEditable      = "EE"
  Const conPEditEditableStr   = "&EE=1"
  Const conPEditOrderField    = "EOF"
  Const conPEditOrderDesc     = "EOD"
  Const conPEditOrderDescStr  = "&EOD=1"
  Const conPEditFind          = "EF"
  Const conPEditFilter        = "ER"
  Const conPEditList          = "EL"
  Const conPEditDefaults      = "ED"
  Const conPEditTab           = "ET"
  Const conPEditGridOptionStr = "&EGO=1"
  Const conPEditGridOption    = "EGO"

  Const conOptionFirstPage    = "h00"
  Const conOptionExeFirstPage = "h01"
  Const conOptionAZEdit       = "h02"
  Const conOptionFPEdit       = "h03"
  Const conOptionEditor       = "h04"
  Const conOptionExeEdition   = "h05"
  Const conOptionAdd          = "h06"
  Const conOptionCopy         = "h07"
  Const conOptionEdit         = "h08"
  Const conOptionDelete       = "h09"
  Const conOptionSave         = "h10"
  Const conOptionSaveAdd      = "h11"
  Const conOptionSaveCopy     = "h12"

  Const conOptionGrid         = "h13"

  Const conOptionFindCommon   = "h14"
  Const conOptionFindAdvanced = "h15"
  Const conOptionExeFind      = "h16"
  Const conOptionSummary      = "h17"

  Const constrProximo         = "proximo "

  Const constrAdd             = "Incluir"
  Const constrEdit            = "Alterar"
  Const constrDelete          = "Excluir"
  Const constrSave            = "Salvar"
  Const constrFind            = "Localizar"
  Const constrList            = "Listagem"
  Const constrRefresh         = "Atualizar"
  Const constrCopy            = "Copiar"
  Const constrSaveAdd         = "Salvar e Incluir"
  Const constrSaveCopy        = "Salvar e Copiar"
  Const constrFirst           = " |<  "
  Const constrPrevious        = "  <  "
  Const constrNext            = "  >  "
  Const constrLast            = "  >| "
  Const constrTab             = "Tab"

  Const constrOK              = "   OK   "
  Const constrCancel          = "Cancelar"
  Const constrYes             = "   Sim   "
  Const constrNo              = "   Não   "

  Public sparEditOption
  Public sparEditWhat
  Public sparEditMove
  Public sparEditEditable
  Public sparEditFind
  Public sparEditFilter
  Public sparEditList
  Public sparEditDefaults
  Public sparEditTab
  Public sparEditGridOption
  
  Public sintCurrentTableNum
  Public sintCurrentRelation
  Public sstrSubText
  Public sstrGroupFieldLastValue
  Public sstrGroupField
  Public sstrFields
  Public sstrFields2
  Public sblnHaveSum
  Public sblnHaveGraph
  Public sintAbsolutePosition
  Public sintRecordCount
  Public sintLastNumber
  Public objUpload

  Public sobjForm
  Public sobjShow
  Public sobjTabela
  Public sobjBrowse

  Set sobjForm  = Server.CreateObject("ZTITools30.Form")
  Set sobjShow  = Server.CreateObject("ZTITools30.Show")
  Set sobjTabela = Server.CreateObject("ZTITools30.Table")
  Set sobjBrowse = Server.CreateObject("ZTITools30.Browse")

Dim sstrText, sstrText2, sstrText3, sstrLastValidate, sintSize, sintNumCol, sstrCheckAsRadio, sstrTab

  Public sstrCreateURL, sstrCreateURLMode, sstrCreateURLTarget

  Session("TableNumber") = 0
  sintSize = 0

REM -------------------------------------------------------------------------
REM Final do include EditForm.inc

%>