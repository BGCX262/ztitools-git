<script language=JavaScript>
<!--
// --------------------------------------------------------------------------
// Toolbar
// --------------------------------------------------------------------------
// Descricao   : Gerenciador de menus do Navegador de Publicacao
// Criacao     : 11:23h 23/2/1998
// Local       : Brasilia/DF
// Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
// Versao      : 1.0.0
// Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
// --------------------------------------------------------------------------

// --------------------------------------------------------------------------
// Verificacao do tipo de browser
// --------------------------------------------------------------------------
bName = navigator.appName;
bVer = navigator.appVersion;

if (((bName == "Netscape" && bVer > "3.0") && (bName == "Netscape" && bVer  < "4.01"))
  ||
  ((bName == "Netscape" && bVer >= "4.03"))
  ||
  (bName == "Microsoft Internet Explorer" && bVer >= "4.0"))

  br = "n3";

else br = "n2";

var AHREF;
var Status;
var ATarget;
var IMGWidth;
var IMGHeight;
var IMGRoot;

var cmdover = new Array();
var cmdout = new Array();
var cmddown = new Array();
var cmdoff = new Array();
var cmdcounter = 0;

// --------------------------------------------------------------------------
// Adciona item no toolbar
// --------------------------------------------------------------------------

function AddButton(IMGSrc, IMGAlt) {
  if (br == "n3") {
    var strStatus = "Active.gif"
    var argv = AddButton.arguments;
    var argc = AddButton.arguments.length;
    AHREF = (argc > 2) ? argv[2] : AHREF;
    Status = (argc > 3) ? argv[3] : Status;
    ATarget = (argc > 4) ? argv[4] : ATarget;   
    IMGWidth = (argc > 5) ? argv[5] : IMGWidth; 
    IMGHeight = (argc > 6) ? argv[6] : IMGHeight; 
    IMGRoot = (argc > 6) ? argv[7] : IMGRoot; 

    if (cmdcounter == 0)
      document.write("<img src='" + IMGRoot + "inicio.gif' width=9 height=26>");

    cmdover[cmdcounter] = new Image();
    cmdover[cmdcounter].src = IMGRoot +IMGSrc + "ActiveColor.gif"

    cmdout[cmdcounter] = new Image();
    cmdout[cmdcounter].src = IMGRoot +IMGSrc + "Active.gif"

    cmddown[cmdcounter] = new Image();
    cmddown[cmdcounter].src = IMGRoot +IMGSrc + "Down.gif"

    cmdoff[cmdcounter] = new Image();
    cmdoff[cmdcounter].src = IMGRoot +IMGSrc + "Inactive.gif"

    if (Status == "A") 
      strStatus = "Active.gif";

    if (Status == "B") 
      strStatus = "Down.gif";

    if (Status == null) 
      strStatus = "Inactive.gif";

    imgsrc = "<a href='" + AHREF + "' target=" + ATarget + " name=acmd" + cmdcounter + "></a>"
    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"

    SetCookie("Cookiecmd" + cmdcounter, Status);
    
    document.write(imgsrc);

    cmdcounter++

  }
}

// --------------------------------------------------------------------------
// Adiciona um espaco
// --------------------------------------------------------------------------
function AddSpace() {
  if (br == "n3") {
    document.write("<img src='" + IMGRoot + "space.gif' width=8 height=26>");
  }
}

// --------------------------------------------------------------------------
// Inicializa os cookies dos butoes
// --------------------------------------------------------------------------
function cmdInit() {
  if (br == "n3") {
    for (x = 0; x < cmdcounter; x++) {
      if (GetCookie("Cookiecmd" + x) == null) {
        SetCookie(("Cookiecmd" + x), "A");
        
      }
    }
  }
}

// --------------------------------------------------------------------------
// Atualiza o estado dos botoes
// --------------------------------------------------------------------------
function cmdRefresh() {
  if (br == "n3") {
    for (x = 0; x < cmdcounter; x++) {
      cmdName = "cmd" + x

      if (GetCookie("Cookiecmd" + x) == "A")
        document[cmdName].src = cmdout[x].src;

      if (GetCookie("Cookiecmd" + x) == "B")
        document[cmdName].src = cmddown[x].src;

      if (GetCookie("Cookiecmd" + x) == null)        
        document[cmdName].src = cmdoff[x].src;

    }
  }
}
// --------------------------------------------------------------------------
// Gerencia o evento MouseOver
// --------------------------------------------------------------------------
function cmdOver() {
  if (br == "n3") {
    if (window.event.srcElement.tagName == "IMG") {
      cmdName = window.event.srcElement.name;
  
      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);
       
        if (GetCookie("Cookiecmd" + cmdNumber) == "A"){
          document[cmdName].src = cmdover[cmdNumber].src;
        }
      }
    }
  }
}

// --------------------------------------------------------------------------
// Gerencia o evento MouseOut
// --------------------------------------------------------------------------
function cmdOut() {
  if (br == "n3") {
    if (window.event.srcElement.tagName == "IMG") {
      cmdName = window.event.srcElement.name;
  
      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);

        if (GetCookie("Cookiecmd" + cmdNumber) == "A"){
          document[cmdName].src = cmdout[cmdNumber].src;
        }
      }
    }
  }
}
// --------------------------------------------------------------------------
// Gerencia o evento MouseDown
// --------------------------------------------------------------------------
function cmdDown() {
  if (br == "n3") {
    if (window.event.srcElement.tagName == "IMG") {
      cmdName = window.event.srcElement.name;

      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);
        cmdAName = "acmd" + cmdNumber;

        if (GetCookie("Cookiecmd" + cmdNumber) == "A" || GetCookie("Cookiecmd" + cmdNumber) == "B") {
          document[cmdName].src = cmddown[cmdNumber].src;
          this.parent[this[cmdAName].target].location = this[cmdAName].href;
        }
      }
    }
  }
}

// --------------------------------------------------------------------------
// Pega o valor de um cookie
// --------------------------------------------------------------------------
function getCookieVal (offset) {
  var endstr = document.cookie.indexOf (";", offset);
  if (endstr == -1)
	endstr = document.cookie.length;
  return unescape(document.cookie.substring(offset, endstr));
}
// --------------------------------------------------------------------------
// Pega um cookie
// --------------------------------------------------------------------------
function GetCookie (name) {
  var arg = name + "=";
  var alen = arg.length;
  var clen = document.cookie.length;
  var i = 0;
  while (i < clen) {
	var j = i + alen;
	if (document.cookie.substring(i, j) == arg)
	  return getCookieVal (j);
	i = document.cookie.indexOf(" ", i) + 1;
	if (i == 0) break;
  }
  return null;
}

var CookieExpires = null;
var CookiePath = null;
var CookieDomain = null;
var CookieSecure = false;
// --------------------------------------------------------------------------
// Define um cookie
// --------------------------------------------------------------------------
function SetCookie (name, value) {
  var argv = SetCookie.arguments;
  var argc = SetCookie.arguments.length;
  CookieExpires = (argc > 2) ? argv[2] : CookieExpires;
  CookiePath = (argc > 3) ? argv[3] : CookiePath;
  CookieDomain = (argc > 4) ? argv[4] : CookieDomain;
  CookieSecure = (argc > 5) ? argv[5] : CookieSecure;
  
  document.cookie = name + "=" + escape (value) +
	((CookieExpires == null) ? "" : ("; expires=" + CookieExpires.toGMTString())) +
	((CookiePath == null) ? "" : ("; path=" + CookiePath)) +
	((CookieDomain == null) ? "" : ("; domain=" + CookieDomain)) +
	((CookieSecure == true) ? "; CookieSecure" : "");
	
}
// --------------------------------------------------------------------------
// Deleta um cookie
// --------------------------------------------------------------------------
function DeleteCookie (name) {
  var argv = DeleteCookie.arguments;
  var argc = DeleteCookie.arguments.length;
  path = (argc > 1) ? argv[1] : CookiePath;
  domain = (argc > 2) ? argv[2] : CookieDomain;
  secure = (argc > 3) ? argv[3] : CookieSecure;

  var exp = new Date();
  exp.setTime (exp.getTime());  // This cookie is history
  var cval = GetCookie(name);

  document.cookie = name + "=; expires=" + exp.toGMTString() +
	((path == null) ? "" : ("; path=" + path)) +
	((domain == null) ? "" : ("; domain=" + domain)) +
	((secure == true) ? "; secure" : "");

}

// var expdate = new Date();

// expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365));

// SetCookie("Teste", "Zevallos", expdate, "/", "www.projects.zevallos2.com.br", false);

// DeleteCookie("Teste");
  
// alert(GetCookie("Teste"));
//-->
</script>
