<%
REM ==================================================================================================================================================
REM  /TreeView.inc
REM --------------------------------------------------------------------------------------------------------------------------------------------------
REM  Descricao   : Sistema gerador de TreeView.
REM  Criação     : 22/12/99 07:56:00 PM
REM  Local       : Brasília/DF
REM  Elaborado   : Kelson da Costa Medeiros
REM  Versao      : 1.2.2
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM --------------------------------------------------------------------------------------------------------------------------------------------------

  ' Constantes do ambiente
  '-----------------------

  ' Constante com o nome do item raíz
  Const conRoot = "TreeViewRoot"

  ' Sufixo e extensão padrão para o nome das imagens
  Const conOpened    = "Open"
  Const conClosed    = "Close"
  Const conExtension = "gif"

  ' Constantes com os nomes das imagens padrão
  Const conLine            = "/ZTITools/img/folder/Line.gif"
  Const conLineBegin       = "/ZTITools/img/folder/LineBegin.gif"
  Const conLineBeginMinus  = "/ZTITools/img/folder/LineBeginMinus.gif"
  Const conLineBeginPlus   = "/ZTITools/img/folder/LineBeginPlus.gif"
  Const conLineEnd         = "/ZTITools/img/folder/LineEnd.gif"
  Const conLineEndMinus    = "/ZTITools/img/folder/LineEndMinus.gif"
  Const conLineEndPlus     = "/ZTITools/img/folder/LineEndPlus.gif"
  Const conLineMiddle      = "/ZTITools/img/folder/LineMiddle.gif"
  Const conLineMiddleMinus = "/ZTITools/img/folder/LineMiddleMinus.gif"
  Const conLineMiddlePlus  = "/ZTITools/img/folder/LineMiddlePlus.gif"
  Const conSpace           = "/ZTITools/img/folder/Space.gif"

  ' Constantes com os nomes "QueryStrings"
  Const conItemClicked  = "IC"
  Const conNewState     = "NS"
'  Const conThisLevel    = "TL"
'  Const conLookupFilter = "LF"

  ' Constantes com os nomes "QueryStrings"
  Const conStateOpen  = "O"
  Const conStateClose = "C"

  ' Constantes con os nomes dos tipos de itens
  Const conStatic = "Static"

  ' Constantes com os prefixos dos nomes das propriedades do item
  Const conItemType      = "ItemType"
  Const conItemParent    = "ItemParent"
  Const conItemIdent     = "ItemIdent"
  Const conItemText      = "ItemText"
  Const conItemURL       = "ItemURL"
  Const conItemOpened    = "ItemOpened"
  Const conItemIsNode    = "ItemIsNode"
  Const conItemLast      = "ItemLast"
  Const conItemChild     = "ItemChild"
  Const conItemImage     = "ItemImage"
  Const conItemClass     = "ItemClass"
  Const conItemBGColor   = "ItemBGColor"
  Const conItemFontName  = "ItemFontName"
  Const conItemFontColor = "ItemFontColor"
  Const conItemFontSize  = "ItemFontSize"
  Const conItemBold      = "ItemBold"
  Const conItemItalic    = "ItemItalic"
  Const conItemUnderline = "ItemUnderline"
  Const conItemHAlign    = "ItemHAlign"
  Const conItemVAlign    = "ItemVAlign"
  Const conItemTarget    = "ItemTarget"
  Const conItemNoWrap    = "ItemNoWrap"
  Const conItemHeight    = "ItemHeight"
  Const conItemWidth     = "ItemWidth"


  ' Variáveis que armazenam os valores das propriedades para os próximos
  ' items adicionados
  ' --------------------------------------------------------------------
  Dim sstrItemTarget, sstrItemBGColor, sstrItemWidth, sstrItemHeight
  Dim sbnlItemBold, sbnlItemItalic, sbnlItemUnderline, sbnlItemNoWrap
  Dim sstrItemFontColor, sstrItemFontName, sstrItemFontSize
  Dim sstrItemHAlign, sstrItemVAlign


  ' Variaveis globais
  ' -----------------

  ' Nome da imagem padrão para os links
  Dim sstrDefaultImage

  ' Cor de fundo do "TreeView"
  Dim sstrTreeBGColor

  ' Pacote com os nomes dos nós
  Dim sdicNodes

  ' Pacote com os itens adicionados
  Dim sdicItems

  ' Pacote com as imagens adicionados
  Dim sdicImages

  ' Quantidade de itens adicionados
  Dim sintItemsCount

  ' Quantidade de imagens adicionadas
  Dim sintImagesCount

  ' Quantidade de nós definidos para o próximo item que for adicionado
  Dim sintMaxNodes

  ' Variáveis que armazenam os valos dos "QueryStrings"
  Dim sstrItemClicked
  Dim sstrNewState

  ' Criação dos Objetos
  ' -------------------

  ' Pacote com os nomes dos nós
  Set sdicNodes = Server.CreateObject( "Scripting.Dictionary" )

  ' Pacote com as propriedades dos itens adicionados
  Set sdicItems = Server.CreateObject( "Scripting.Dictionary" )

  ' Pacote com as propriedades das imagens adicionados
  Set sdicImages = Server.CreateObject( "Scripting.Dictionary" )


  ' Configuração dos "QueryStrings"
  ' -------------------------------

  ' Armazena o valor do "QueryString" "conTA", tipo de ação a ser executada
  sstrItemClicked = Request.QueryString( conItemClicked )

  ' Armazena o valor do "QueryString" "conTI", item que executou a ação
  sstrNewState = Request.QueryString( conNewState )


REM =========================================================================
REM Inicializa a definição do "TreeView"
REM -------------------------------------------------------------------------
Public Sub TreeViewBeginTree

  ' Inicializa os contadores
  sintItemsCount  = 0
  sintImagesCount = 0
  sintMaxNodes    = 0

  ' Inicializa as propriedades
  sstrItemBGColor   = ""
  sstrItemTarget    = ""
  sstrItemHAlign    = ""
  sstrItemVAlign    = ""
  sbnlItemBold      = False
  sbnlItemItalic    = False
  sbnlItemUnderline = False
  sbnlItemNoWrap    = False
  sstrItemFontColor = ""
  sstrItemFontName  = ""
  sstrItemFontSize  = ""
  sstrItemHeight    = ""
  sstrItemWidth     = ""
  sstrTreeBGColor   = ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewBeginTree
REM =========================================================================

REM =========================================================================
REM Adiciona um nó a este "TreeView".
REM -------------------------------------------------------------------------
Public Sub TreeViewBeginNode( ByVal strText, ByVal strURL, ByVal blnOpened, ByVal strImage )

  ' Adiciona o item no pacote
  TreeViewAddStatic strText, strURL, blnOpened, True, strImage

  ' Adiciona este item na lista de nós abertos
  sdicNodes.Add conItemParent & sdicNodes.Count + 1, sintItemsCount

  ' Se a quantidade de nós para o próximo item for maior que a quantidade máxima de nós
  If sdicNodes.Count > sintMaxNodes Then

     ' Armazena a nova quantidade máxima de nós
     sintMaxNodes = sdicNodes.Count

  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewBeginNode
REM =========================================================================

REM =========================================================================
REM Adicina um item estático ao "TreeView".
REM -------------------------------------------------------------------------
Public Sub TreeViewEndNode

  ' Remove o último item da lista de nós abertos
  sdicNodes.Remove( conItemParent & sdicNodes.Count )

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewBeginStatic
REM =========================================================================

REM =========================================================================
REM Adiciona um nó a este "TreeView".
REM -------------------------------------------------------------------------
Public Sub TreeViewAddItem( ByVal strText, ByVal strURL, ByVal strImage )

  ' Adiciona o item no pacote
  TreeViewAddStatic strText, strURL, False, False, strImage

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewBeginNode
REM =========================================================================

REM =========================================================================
REM Finaliza o "TreeView" e inicia o código que o constrói.
REM -------------------------------------------------------------------------
Public Sub TreeViewEndTree

  ' Prepara os items para depois incluílos.
  TreeViewPrepareTree

  ' Mostra todos os items do "TreeView".
  TreeViewMakeTree

  ' Destrói o objeto que armazena os nomes dos nós
  Set sdicNodes = Nothing

  ' Destrói o objeto que armazena os itens
  Set sdicItems = Nothing

  ' Destrói o objeto que armazena as imagens
  Set sdicImages = Nothing

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewShowTree
REM =========================================================================

REM =========================================================================
REM Adiciona uma imagem à lista de imagens disponíveis aos itens. Desta forma
REM ao se adicionar o nome os itens não é necessário informar o nome da
REM imagem completo novamente, bastando informar o nome chave da imagem
REM -------------------------------------------------------------------------
Public Sub TreeViewAddImage( ByVal strInitial, ByVal strAddress, ByVal strOpen, ByVal strClose, ByVal strExtension )

  ' Se o usuário tentar adicionar uma imagem que já exista
  If sdicImages.Exists( strInitial & conOpened ) Then

     ' Informa ao mesmo que esta imagem já existe
     Show.Message "Já foi adicionada uma imagem com o nome """ & strInitial & """."

     ' Não continua com a execução do procedimento
     Exit Sub

  End If

  ' Incrementa o contador de items adicionados
  sintImagesCount = ( sintImagesCount + 1 )

  ' Se o usuário não informou o sufixo para o nome da imagem aberta
  If Not strOpen > "" Then

     ' Informa o sufixo padrão
     strOpen = conOpened

  End If

  ' Se o usuário não informou o sufixo para o nome da imagem aberta
  If Not strClose > "" Then

     ' Informa o sufixo padrão
     strClose = conClosed

  End If

  ' Se o usuário não informou o sufixo para o nome da imagem aberta
  If Not strExtension > "" Then


     ' Informa o sufixo padrão
     strExtension = conExtension

  End If

  ' Adiciona no pacote as propriedades deste item
  sdicImages.Add strInitial & conOpened, strAddress & strInitial & strOpen & "." & strExtension
  sdicImages.Add strInitial & conClosed, strAddress & strInitial & strClose & "." & strExtension

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewAddImage
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a imagem padrão a ser
REM utilizado quando o usuário não informar o nome chave da imagem
REM -------------------------------------------------------------------------
Public Sub TreeViewDefaultImage( ByVal strImageInitial )

  ' Armazena o nome da imagem padrão
  sstrDefaultImage = strImageInitial

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewDefaultImage
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a cor de fundo para os
REM próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemBGColor( ByVal strColor )

  ' Armazena a cor de fundo para os próximos itens adicionados
  sstrItemBGColor = strColor

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemBGColor
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a cor da fonte para os
REM próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemFontColor( ByVal strColor )

  ' Armazena a cor da fonte para os próximos itens adicionados
  sstrItemFontColor = strColor

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemFontColor
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é o nome da fonte para os
REM próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemFontName( ByVal strName )

  ' Armazena o nome da fonte para os próximos itens adicionados
  sstrItemFontName = strName

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemFontName
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é o tamanho da fonte para os
REM próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemFontSize( ByVal strSize )

  ' Armazena o tamanho da fonte para os próximos itens adicionados
  sstrItemFontSize = strSize

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemFontSize
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define se os próximos itens adicionados
REM aparecerão com negrito
REM -------------------------------------------------------------------------
Public Sub TreeViewItemBold( ByVal bnlBold )

  ' Armazena se o usuário deseja ou não utilizar negrito no texto dos
  ' próximos itens que serão adicionados
  sbnlItemBold = bnlBold

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemBold
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define se os próximos itens adicionados
REM aparecerão com itálico
REM -------------------------------------------------------------------------
Public Sub TreeViewItemItalic( ByVal bnlItalic )

  ' Armazena se o usuário deseja ou não utilizar itálico no texto dos
  ' próximos itens que serão adicionados
  sbnlItemItalic = bnlItalic

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemItalic
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define se os próximos itens adicionados
REM aparecerão com subscrito
REM -------------------------------------------------------------------------
Public Sub TreeViewItemUnderline( ByVal bnlUnderline )

  ' Armazena se o usuário deseja ou não utilizar subscrito no texto dos
  ' próximos itens que serão adicionados
  sbnlItemUnderline = bnlUnderline

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemUnderline
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é o alinhamento horizontal
REM para os próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemHAlign( ByVal strHAlign )

  ' Armazena o tipo de alinhamento horizontal que o usuário deseja utilizar
  sstrItemHAlign = strHAlign

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemHAlign
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é o alinhamento vertical
REM para os próximos itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemVAlign( ByVal strVAlign )

  ' Armazena o tipo de alinhamento vertical que o usuário deseja utilizar
  sstrItemVAlign = strVAlign

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemVAlign
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é o alvo para os  próximos
REM itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemTarget( ByVal strItemTarget )

  ' Armazena o nome do alvo que o usuário deseja nos próximos itens
  sstrItemTarget = strItemTarget

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemTarget
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define se o texto do item poderá ou não
REM aparecer quebrado em multiplas linhas
REM -------------------------------------------------------------------------
Public Sub TreeViewItemNoWrap( ByVal bnlNoWrap )

  ' Informa que os próximos item poderão ou ser ter seu conteúdo quebrado
  sbnlItemNoWrap = bnlNoWrap

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemNoWrap
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a largura para os próximos
REM itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemWidth( ByVal strWidth )

  ' Armazena a largura de apresentação dos próximos itens
  sstrItemWidth = strWidth

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemWidth
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a altura para os próximos
REM itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewItemHeight( ByVal strHeight )

  ' Armazena a largura de apresentação dos próximos itens
  sstrItemHeight = strHeight

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemHeight
REM =========================================================================

REM =========================================================================
REM Este método (futura propriedade) define qual é a altura para os próximos
REM itens adicionados
REM -------------------------------------------------------------------------
Public Sub TreeViewTreeBGColor( ByVal strColor )

  ' Armazena a cor de fundo do "TreeView"
  sstrTreeBGColor = strColor

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewItemHeight
REM =========================================================================

REM =========================================================================
REM Faz o Igual para parametros
REM -------------------------------------------------------------------------
Public Function TreeViewEqual( ByVal strOption, ByVal strValue )

  ' Retorno da função
  TreeViewEqual = "&" & strOption & "=" & strValue

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ShowMenuItem
REM =========================================================================

REM =========================================================================
REM Adicina um item estático ao "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewAddStatic( ByVal strText, ByVal strURL, ByVal blnOpened, ByVal bnlIsNode, ByVal strImage )

  ' Variável para controle do "Loop" "For"
  Dim i

  ' Incrementa o contador de items adicionados
  sintItemsCount = ( sintItemsCount + 1 )

  ' Adiciona no pacote as propriedades deste item
  sdicItems.Add conItemType   & sintItemsCount, conStatic
  If sdicNodes.Count > 0 Then
     sdicItems.Add conItemParent & sintItemsCount, sdicNodes( conItemParent & sdicNodes.Count )
  Else
     sdicItems.Add conItemParent & sintItemsCount, 0
  End If
  sdicItems.Add conItemIdent      & sintItemsCount, sdicNodes.Count
  sdicItems.Add conItemText       & sintItemsCount, strText
  sdicItems.Add conItemURL        & sintItemsCount, strURL
  sdicItems.Add conItemIsNode     & sintItemsCount, bnlIsNode
  sdicItems.Add conItemOpened     & sintItemsCount, blnOpened
  sdicItems.Add conItemChild      & sintItemsCount, False
  sdicItems.Add conItemImage      & sintItemsCount, strImage
  sdicItems.Add conItemBGColor    & sintItemsCount, sstrItemBGColor
  sdicItems.Add conItemFontName   & sintItemsCount, sstrItemFontName
  sdicItems.Add conItemFontColor  & sintItemsCount, sstrItemFontColor
  sdicItems.Add conItemFontSize   & sintItemsCount, sstrItemFontSize
  sdicItems.Add conItemBold       & sintItemsCount, sbnlItemBold
  sdicItems.Add conItemItalic     & sintItemsCount, sbnlItemItalic
  sdicItems.Add conItemUnderline  & sintItemsCount, sbnlItemUnderline
  sdicItems.Add conItemHAlign     & sintItemsCount, sstrItemHAlign
  sdicItems.Add conItemVAlign     & sintItemsCount, sstrItemVAlign
  sdicItems.Add conItemTarget     & sintItemsCount, sstrItemTarget
  sdicItems.Add conItemNoWrap     & sintItemsCount, sbnlItemNoWrap
  sdicItems.Add conItemWidth      & sintItemsCount, sstrItemWidth
  sdicItems.Add conItemHeight     & sintItemsCount, sstrItemHeight

  ' Informa ao pai deste item que ele agora tem filhos
  sdicItems( conItemChild & sdicItems( conItemParent & sintItemsCount ) ) = True

  ' Percorre todos os item
  For i = 1 To sintItemsCount - 1

    ' Verifica qual tem o mesmo pai do item que eu vou adicionar agora
    If sdicItems( conItemParent & i ) = sdicItems( conItemParent & sintItemsCount ) Then

       ' Indica que ele já não é mais o último
       sdicItems( conItemLast & i ) = False

    End If

  Next

  ' Informa que este novo item é o último deste pai
  sdicItems.Add conItemLast   & sintItemsCount, True

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewAddStatic
REM =========================================================================

REM =========================================================================
REM Prepara os items para depois incluílos.
REM -------------------------------------------------------------------------
Private Sub TreeViewPrepareTree

  ' Variável para controle do "Loop" "For"
  Dim i, j

  ' Percorre todos os item
  For i = 1 To sintItemsCount

    ' Verifica se existe um nó sem filho
    If ( sdicItems( conItemIsNode & i ) ) And ( Not sdicItems( conItemChild & i ) ) Then

       ' Já que este nó não tem filho, tornarei este nó fechado
       sdicItems( conItemOpened & i ) = False

    End If


    ' Se este é o item clicado
    If i = cint( sstrItemClicked ) Then

       ' Copia o índice deste item para a variável "j"
       j = i

       ' Percorro todos os pais deste item
       Do While sdicItems( conItemParent & j ) > 0

          ' Informa que este item está sendo fechado
          sdicItems( conItemOpened & j ) = True

          ' Informa que este item está aberto
          sdicItems( conItemOpened & sdicItems( conItemParent & j ) ) = True

          ' Aponta o índice para o pai deste item
          j = sdicItems( conItemParent & j )

       Loop

       ' Se o usuário esta abrindo este item
       If sstrNewState = conStateOpen Then

          ' Informa que este item está sendo fechado
          sdicItems( conItemOpened & i ) = True

       ' Se o usuário está fechando este item
       Else

          ' Informa que este item está sendo fechado
          sdicItems( conItemOpened & i ) = False

       End If

    End If

  Next

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewMakeTree
REM =========================================================================

REM =========================================================================
REM Mostra todos os items do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewMakeTree

  ' Variável para controle do "Loop" "For"
  Dim intIndex

  ' Retirando o espaço entre as celulas da tabela
  Table.Padding = 0
  Table.Spacing = 0

  ' Configurando a cor de fundo do "TreeView"
  Table.Style.Color2 = sstrTreeBGColor

  ' Inicializa uma nova tabela onde será formatado o "TreeView"
  Table.BeginTable "100%"

  ' Percorre todos os itens do "TreeView" para construílo
  For intIndex = 1 To sintItemsCount

      ' Se o item for do tipo "Static"
      If sdicItems( conItemType & intIndex ) = conStatic Then

         ' Desenha um item estático do "TreeView".
         TreeViewShowStaticItem( intIndex )

      End If

  Next

  ' Finaliza a tabela utilizada na construção do "TreeView"
  Table.EndTable

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewMakeTree
REM =========================================================================

REM =========================================================================
REM Desenha um item estático do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewShowStaticItem( ByVal intIndex )

  ' Se o pai não estiver aberto
  If Not sdicItems( conItemOpened & sdicItems( conItemParent & intIndex ) ) And ( sdicItems( conItemParent & intIndex ) > 0 ) Then

     ' Informa que este item está sendo fechado
     sdicItems( conItemOpened & intIndex ) = False

     ' Caio fora do procedimento
     Exit Sub

  End If

  ' Adiciona um novo item, ou seja, inicializa uma nova linha na tabela
  Table.BeginRow

    ' Desenha a indentação do item do "TreeView".
    TreeViewShowIdent( intIndex )

    ' Mostra a image que indica o estado do item
    TreeViewStateImage( intIndex )

    ' Mostra o ícone do item
    TreeViewShowIcon( intIndex )

    ' Mostra o título do item do "TreeView".
    TreeViewShowCaption( intIndex )

  ' Finaliza a linha de desenho do item
  Table.EndRow

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewShowStaticItem
REM =========================================================================

REM =========================================================================
REM Desenha a indentação do item do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewShowIdent( ByVal intIndex )

  ' Variáveis para controle dos "Loop"
  Dim i, j

  ' Pacote que armazena o nomes das imagens que deverão ser mostradas
  Dim dicParents

  ' Define o tamanho das colunas anteriores à que aparece o texto para que fique com tamanho igua a 1%
  Table.ColumnWidth = "1%"

  ' Copia o índice deste item para a variável "j"
  j = intIndex

  ' Iniciliza a variável será o que armazenará o valor
  i = 1

  ' Iniciliza o pacote
  Set dicParents = Server.CreateObject( "Scripting.Dictionary" )

  ' Percorre todos os itens pai deste item verificando a necessidade de desenhar a linha
  ' de continuação
  Do While sdicItems( conItemParent & j ) > 0

     ' Verifica se este item pai é o último item do nó
     If sdicItems( conItemLast & sdicItems( conItemParent & j ) ) Then

        ' Já que este item é o último item deste nó, não desenha a linha de continuação
        dicParents.Add "Parent" & dicParents.Count + 1, conSpace

     ' Já que este item ainda não é o último item deste nó
     Else

        ' Desenha a linha de continuação
        dicParents.Add "Parent" & dicParents.Count + 1, conLine

     End If

     ' Aponta o índice para o pai deste item
     j = sdicItems( conItemParent & j )

  Loop

  ' Percorre o pacote com os nomes dos objetos de traz para frente, desta forma
  For i = dicParents.Count To 1 Step - 1

      ' Adiciona a coluna responsável pelo incremento
      Table.Column "<img src=""" & dicParents( "Parent" & i ) & """ border=0>"

  Next

  ' Destrói o pacote
  Set dicParents = Nothing

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewShowIdent
REM =========================================================================

REM =========================================================================
REM Desenha o estado do item do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewStateImage( ByVal intIndex )

  ' Inicializa a coluna onde será mostrado o estado do item
  Table.BeginColumn

    ' Se este item for um nó
    If ( sdicItems( conItemIsNode & intIndex ) ) And ( sdicItems( conItemChild & intIndex )  ) Then

       ' Começa a construção do link que irá me reinvocar executando a ação a mim relacionada
       URL.BeginURL Initializer.ScriptName

         ' Adiciona no link o ponteiro para o item clicado
         URL.Add conItemClicked, intIndex

         ' Se este item estiver aberto
         If sdicItems( conItemOpened & intIndex ) Then

            ' Informa no link o novo estado deste item
            URL.Add conNewState, conStateClose

            ' Se este for o último item da arvore
            If sdicItems( conItemLast & intIndex ) Then

               ' Adiciono a imagem [-] sem continuação de linha
               URL.Show "<img src=""" & conLineEndMinus & """ border=0>"

            ' Se este item estiver fechado
            Else

               ' Adiciono a imagem [-] com continuação de linha
               URL.Show "<img src=""" & conLineMiddleMinus & """ border=0>"

            End If

         ' Se este não for o último item da lista
         Else

            ' Informa no link o novo estado deste item
            URL.Add conNewState, conStateOpen

            ' Se este for o último item da arvore
            If sdicItems( conItemLast & intIndex ) Then

               ' Adiciono a imagem [+] sem continuação de linha
               URL.Show "<img src=""" & conLineEndPlus & """ border=0>"

            ' Se este item estiver fechado
            Else

               ' Adiciono a imagem [+] com continuação de linha
               URL.Show "<img src=""" & conLineMiddlePlus & """ border=0>"

            End If

         End If

      ' Finaliza a construção da URL
      URL.EndURL

    ' Se este for um item comum
    Else

       ' Se este for o primeiro item
       If ( intIndex = 1 ) Then

          ' Adiciono uma coluna onde ficará apenas o traço no meio com continuação de linha
          Show.Image conLineBegin

       ' Se este for um item do meio da lista
       ElseIf Not sdicItems( conItemLast & intIndex ) Then

          ' Adiciono uma coluna onde ficará apenas o traço no meio com continuação de linha
          Show.Image conLineMiddle

       ' Se este for o último item da arvore
       Else

          ' Adiciono uma coluna onde ficará apenas o traço no meio sem continuação de linha
          Show.Image conLineEnd

       End If

    End If

  ' Finaliza esta coluna
  Table.EndColumn

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewStateImage
REM =========================================================================

REM =========================================================================
REM Desenha o ícone do item do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewShowIcon( ByVal intIndex )

  ' Variável que armazena o nome da imagem que deverá ser aberta
  Dim strImageName

  ' Se este item estiver aberto
  If sdicItems( conItemOpened & intIndex ) Then

    ' Configura o nome da imagem que deverá ser mostrada para a padrão aberta
    strImageName = sstrDefaultImage & conOpened

  ' Se este item estiver fechado
  Else

    ' Configura o nome da imagem que deverá ser mostrada para a padrão fechada
    strImageName = sstrDefaultImage & conClosed

  End If

  ' Se foi informado o nome da imagem para este item
  If sdicItems( conItemImage & intIndex ) > "" Then

     ' Se este item estiver aberto
     If sdicItems( conItemOpened & intIndex ) Then

        ' Configura o nome da imagem que deverá ser mostrada a imagem aberta
        strImageName = sdicItems( conItemImage & intIndex ) & conOpened

     ' Se este item estiver fechado
     Else

        ' Configura o nome da imagem que deverá ser mostrada a imagem fechada
        strImageName = sdicItems( conItemImage & intIndex ) & conClosed

     End If

  End If

  ' Se esta imagem existir na lista de imagens
  If sdicImages.Exists( strImageName ) Then

     ' Adiciono uma coluna onde ficará a imagem de estado
     Table.Column "<img src=""" & sdicImages( strImageName ) & """ border=0>"

  End If


End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewShowIcon
REM =========================================================================

REM =========================================================================
REM Mostra o título do item do "TreeView".
REM -------------------------------------------------------------------------
Private Sub TreeViewShowCaption( ByVal intIndex )

  ' Como vou começar a definição do texto do item, redefino a largura da coluna para a padrão
  Table.ColumnWidth = ""

  ' Espande a coluna para que se espanda até que ocupe a quantidade de colunas restantes da tabela
  Table.ColumnColSpan = ( ( sintMaxNodes * 2 ) - sdicItems( conItemIdent & intIndex ) )

  ' Configura a cor de fundo deste item
  Table.ColumnColor = sdicItems( conItemBGColor & intIndex )

  ' Informa que está coluna não pode quebrar
  Table.ColumnNoWrap = sdicItems( conItemNoWrap & intIndex )

  ' Informa qual será o alinhamento horizontal desta coluna
  Table.ColumnAlign = sdicItems( conItemHAlign & intIndex )

  ' Informa qual será o alinhamento vertical desta coluna
  Table.ColumnVAlign = sdicItems( conItemVAlign & intIndex )

  ' Informa qual será a altura desta coluna
  Table.ColumnHeight = sdicItems( conItemHeight & intIndex )

  ' Informa qual será a largura desta coluna
  Table.ColumnWidth = sdicItems( conItemWidth & intIndex )

  ' Inicializa a coluna onde fica o texto do item
  Table.BeginColumn

    ' Inicializa as "tags" necessárias para mostrar o texto formatado do
    ' modo definido pelo usuário
    TreeViewBeginShowCaption( intIndex )

    ' Mostra o texto
    Show.HTML sdicItems( conItemText & intIndex )

    ' Finaliza as "tags" que foram abertas
    TreeViewEndShowCaption( intIndex )

  ' Finaliza a coluna onde fica o texto do item
  Table.EndColumn

  ' Configura o valor padrão para a propriedade "ColumnHeight"
  Table.ColumnHeight = ""

  ' Configura o valor padrão para a propriedade "ColumnWidth"
  Table.ColumnWidth = ""

  ' Configura o valor padrão para a propriedade "ColumnColSpan"
  Table.ColumnColSpan = ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewShowCaption
REM =========================================================================

REM =========================================================================
REM Inicia a configuração do texto.
REM -------------------------------------------------------------------------
Private Sub TreeViewBeginShowCaption( ByVal intIndex )

   ' Mostra o espaço separador antes do texto do item
   Show.HTML "&nbsp;"

  ' Se este item deve aparecer em negrito
  If sdicItems( conItemBold & intIndex ) Then

     ' Inicializa o negrito
     Show.Bold

  End If

  ' Se este item deve aparecer em italico
  If sdicItems( conItemItalic & intIndex ) Then

     ' Inicializa o italico
     Show.Italic

  End If

  ' Se este item deve aparecer em subscrito
  If sdicItems( conItemUnderline & intIndex ) Then

     ' Inicializa o subscrito
     Show.Underline

  End If

  ' Se alguma das propriedades da fonte foi configurada
  If ( sdicItems( conItemFontName  & intIndex ) > "" ) Or _
     ( sdicItems( conItemFontColor & intIndex ) > "" ) Or _
     ( sdicItems( conItemFontSize  & intIndex ) > "" ) Then

     ' Inicializa a "tag" de fonte
     Show.HTML "<font"

     ' Se o nome da fonte foi informado
     If sdicItems( conItemFontName & intIndex ) > "" Then

        ' Configura o nome da fonte
        Show.HTML " face=""" & sdicItems( conItemFontName & intIndex ) & """"

     End If

     ' Se a cor da fonte foi informada
     If sdicItems( conItemFontColor & intIndex ) > "" Then

        ' Configura a cor da fonte
        Show.HTML " color=""" & sdicItems( conItemFontColor & intIndex ) & """"

     End If

     ' Se o tamanho da fonte foi informado
     If sdicItems( conItemFontSize & intIndex ) > "" Then

        ' Configura o tamanho da fonte
        Show.HTML " size=""" & sdicItems( conItemFontSize & intIndex ) & """"

     End If

     ' Adiciona ao HTML o sinal ">" da tag "<font>"
     Show.HTML ">"

  End If

  ' Se o "link" não foi informado
  If sdicItems( conItemURL & intIndex ) > "" Then

    ' Inicializa "link" do texto
    Show.HTML "<a href=" & sdicItems( conItemURL & intIndex )

    ' Se houver um alvo configurado
    If sdicItems( conItemTarget & intIndex ) > "" Then

       ' Mostra o texto com o "link"
       Show.HTML " target=""" & sdicItems( conItemTarget & intIndex ) & """"

    End If

    ' Adiciona ao HTML o sinal ">" da tag "<a>"
    Show.HTML ">"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewBeginShowCaption
REM =========================================================================

REM =========================================================================
REM Finaliza a configuração do texto.
REM -------------------------------------------------------------------------
Private Sub TreeViewEndShowCaption( ByVal intIndex )

  ' Se o "link" não foi informado
  If sdicItems( conItemURL & intIndex ) > "" Then

    ' Inicializa "link" do texto
    Show.HTML "</a>"

  End If

  If ( sdicItems( conItemFontName  & intIndex ) > "" ) Or _
     ( sdicItems( conItemFontColor & intIndex ) > "" ) Or _
     ( sdicItems( conItemFontSize  & intIndex ) > "" ) Then

     Show.HTML "</font>"

  End If


  ' Se este item deve aparecer em subscrito
  If sdicItems( conItemUnderline & intIndex ) Then

     ' Inicializa o subscrito
     Show.UnderlineEnd

  End If

  ' Se este item deve aparecer em italico
  If sdicItems( conItemItalic & intIndex ) Then

     ' Inicializa o italico
     Show.ItalicEnd

  End If

  ' Se este item deve aparecer em negrito
  If sdicItems( conItemBold & intIndex ) Then

     ' Inicializa o negrito
     Show.BoldEnd

  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do TreeViewEndShowCaption
REM =========================================================================


REM --------------------------------------------------------------------------------------------------------------------------------------------------
REM Fim do TreeView.inc
REM ==================================================================================================================================================
' TreeView.BeginTree
'
'    TreeView.AddImage( strInitial, strAddress, [strOpen], [strClose], [strExtension ] )
'    TreeView.DefaultImage( strImageInitial )
'
'    Tree.BeginNode( strText, strURL, [blnOpened], [strImageInitial] )
'
'      TreeView.ItemTarget       = strTarget
'      TreeView.ItemNoWrap       = bnlNoWrap
'      TreeView.ItemBackColor    = strBackColor
'      TreeView.ItemBold         = bnlBold
'      TreeView.ItemItalic       = bnlItalic
'      TreeView.ItemUnderline    = bnlUnderline
'      TreeView.ItemHAlign       = strHAlign
'      TreeView.ItemVAlign       = strVAlign
'      TreeView.ItemWidth        = strWidth
'      TreeView.ItemHeight       = strHeight
'      TreeView.ItemFontColor    = strColor
'      TreeView.ItemFontName     = strFace
'      TreeView.ItemFontSize     = strSize
'      TreeView.TreeBGColor      = strColor
'
'      ? TreeView.AddLookup ( strTable, strKeyField, strDisplayField, intMaxSize, intMaxLines, bnlWordWrap, strOptionURL )
'      ? TreeView.LookupFilterAdd
'      ? TreeView.AddColumn
'      * TreeView.ItemStyle = strStyle
'      * TreeView.ItemClass = strClass
'      * TreeView.ItemHover = strHover
'
'      TreeView.AddItem( strText, strURL, [strImageInitial] )
'
'    TreeView.EndNode
'
' TreeView.EndTree
%>