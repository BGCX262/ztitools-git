<script language=JavaScript>
<!--
// --------------------------------------------------------------------------
// Body.js
// --------------------------------------------------------------------------
// Descricao   : Gerenciador da navegacao do corpo
// Criacao     : 11:23h 23/2/1998
// Local       : Brasilia/DF
// Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
// Versao      : 1.0.0
// Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
// --------------------------------------------------------------------------

HRef = "";

// --------------------------------------------------------------------------
// Muda o estado de um botao
// --------------------------------------------------------------------------
function cmdChangeStatus(cmdName, blnStatus, blnState) {
    var argv = cmdChangeStatus.arguments;
    var argc = cmdChangeStatus.arguments.length;
    HRef = (argc > 3) ? argv[3] : HRef;
    
  if (cmdName.substring(0, 3) == "cmd") {
    cmdAName = "a" + cmdName
    cmdNumber = cmdName.substring(3, 99);

    parent.ToolBar.Global.cmdstatus[cmdNumber] = blnStatus;
    parent.ToolBar.Global.cmdstate[cmdNumber] = blnState;

    if (blnStatus) {
      if (blnState)
        Result = parent.ToolBar.Global.cmdout[cmdNumber].src;

      else
        Result = parent.ToolBar.Global.cmddown[cmdNumber].src;
    }
    else
      {
      parent.ToolBar.Global.cmdstate[cmdNumber] = false;
      
      if (blnState)
        Result = parent.ToolBar.Global.cmdoff[cmdNumber].src;

      else
        Result = parent.ToolBar.Global.cmdoff[cmdNumber].src;
      }

    parent.ToolBar.Global.cmdhref[cmdNumber] = HRef;
    
    parent.ToolBar.document[cmdName].src = Result;
  }
}

-->
</SCRIPT>
