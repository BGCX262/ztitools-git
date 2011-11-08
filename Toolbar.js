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
  (bName == "Microsoft Internet Explorer" && bVer >= "4.0")) {

    br = "n3";

  	Global = new globalVars();

    var AHREF;
    var Status;
    var ATarget;
    var IMGWidth;
    var IMGHeight;
    var IMGRoot;

    var cmdcounter = 0;
  }

else br = "n2";

function globalVars(){
  // Sets the global variables for the script. 
  // These may be changed to quickly customize the tree's apperance

  // Botoes  
  this.cmdstatus = new Array();
  this.cmdstate = new Array();
  this.cmdhref = new Array();
  this.cmdover = new Array();
  this.cmdout = new Array();
  this.cmddown = new Array();
  this.cmdoff = new Array();
  this.cmdCounter = 0;

  // Fonts
  this.face="Helv,Arial";
  this.fSize=1;

  // Spacing
  this.vSpace=2;
  this.hSpace=4;
  this.tblWidth=500;
  this.selTColor="#FFCC00";
  this.selFColor="#000000";
  this.selUColor="#CCCCCC";

  // Images
  this.imagedir="/ZTITools/img/folder/";
  this.appIcon = "app";    
  this.spaceImg=this.imagedir  + "space.gif";
  this.lineImg=this.imagedir  + "line.gif";
  this.plusImg=this.imagedir  + "plus.gif";
  this.minusImg=this.imagedir  + "minus.gif";
  this.emptyImg=this.imagedir + "blank.gif";
  this.plusImgLast=this.imagedir  + "plusl.gif";
  this.minusImgLast=this.imagedir  + "minusl.gif";
  this.emptyImgLast=this.imagedir + "blankl.gif";

  //Help
  this.helpFileName="bhelp.htm";
  this.helpDir="http://www.admin.zevallos2.com.br/iishelp/iis/htm/core/"
  

  // Other Flags
  this.showState = false;  
  this.dontAsk = false;
  this.updated = false;
  this.homeurl = top.location.href;
  this.siteProperties = false;
  this.working = false;
  
}

// --------------------------------------------------------------------------
// Adiciona item no toolbar
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

    Global.cmdover[cmdcounter] = new Image();
    Global.cmdover[cmdcounter].src = IMGRoot +IMGSrc + "ActiveColor.gif"

    Global.cmdout[cmdcounter] = new Image();
    Global.cmdout[cmdcounter].src = IMGRoot +IMGSrc + "Active.gif"

    Global.cmddown[cmdcounter] = new Image();
    Global.cmddown[cmdcounter].src = IMGRoot +IMGSrc + "Down.gif"

    Global.cmdoff[cmdcounter] = new Image();
    Global.cmdoff[cmdcounter].src = IMGRoot +IMGSrc + "Inactive.gif"

    Global.cmdhref[cmdcounter] = AHREF;

    Global.cmdstatus[cmdcounter] = true;

    Global.cmdstate[cmdcounter] = Status

    if (Status) 
      strStatus = "Active.gif";

    if (Status) 
      strStatus = "Down.gif";

    if (Status == null) {
      strStatus = "Inactive.gif";
      Global.cmdstatus[cmdcounter] = false;
    }

    imgsrc = "<a href='" + AHREF + "' target=" + ATarget + " name=acmd" + cmdcounter + "></a>"
    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"

    SetCookie("Cookiecmd" + cmdcounter, Status);
    
    document.write(imgsrc);

    cmdcounter++

    Global.cmdCounter = cmdcounter
  }
}

//AddButton.prototype = new GlobalVars;

// --------------------------------------------------------------------------
// Adiciona um espaco
// --------------------------------------------------------------------------
function AddSpace() {
  if (br == "n3") {
    document.write("<img src='" + IMGRoot + "space.gif' width=8 height=26>");
  }
}

//AddSpace.prototype = new GlobalVars;

// --------------------------------------------------------------------------
// Inicializa os cookies dos butoes
// --------------------------------------------------------------------------
function cmdInit() {
  if (br == "n3") {
    for (x = 0; x < cmdcounter; x++) {
      if (Global.cmdstate[x] == null) {
        Global.cmdstate[x] = true;
        
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

      if (Global.cmdstatus[x]) {
        if (Global.cmdstate[x])
          document[cmdName].src = Global.cmdout[x].src;
        
        else 
          document[cmdName].src = Global.cmddown[x].src;
        }
        
      else
        document[cmdName].src = Global.cmdoff[x].src;

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
       
        if (Global.cmdstatus[cmdNumber] && Global.cmdstate[cmdNumber]){
          document[cmdName].src = Global.cmdover[cmdNumber].src;
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

        if (Global.cmdstatus[cmdNumber] && Global.cmdstate[cmdNumber]){
          document[cmdName].src = Global.cmdout[cmdNumber].src;
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

        if (Global.cmdstatus[cmdNumber]) {
          if (Global.cmdstatus[cmdNumber])
            document[cmdName].src = Global.cmddown[cmdNumber].src;
            
          else
            document[cmdName].src = Global.cmddown[cmdNumber].src;
            
          
          this[cmdAName].href = Global.cmdhref[cmdNumber];

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

// --------------------------------------------------------------------------
// Abre a janela do help
// --------------------------------------------------------------------------
function helpBox(){
	if (Global.helpFileName == null){
		alert("Desculpe, a ajuda nao esta disponivel!");
	}
	else{
		thefile = Global.helpDir +Global.helpFileName+".htm";
    window.open(thefile ,"Help","toolbar=no,scrollbars=yes,directories=no,menubar=no,width=375,height=500");
  }
}

// --------------------------------------------------------------------------
// Abre a janela Sobre
// --------------------------------------------------------------------------
function aboutBox() {
  popbox = window.open("/admin/about.asp","about","toolbar=no,scrollbars=no,directories=no,menubar=no,width="+525+",height="+300);

  if(popbox !=null){
    if (popbox.opener == null){
      popbox.opener = self;
    }
  }
}

// --------------------------------------------------------------------------
// Abre a janela de acordo com os parametros
// --------------------------------------------------------------------------
function popBox(title, width, height, filename){
  thefile=(filename + ".asp");
  thefile="pop.asp?pg="+thefile;

  popbox=window.open(thefile,title,"toolbar=no,scrollbars=yes,directories=no,menubar=no,width="+width+",height="+height);
  if(popbox !=null){
    if (popbox.opener==null){
      popbox.opener=self;
    }
  }
}

// var expdate = new Date();

// expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365));

// SetCookie("Teste", "Zevallos", expdate, "/", "www.projects.zevallos2.com.br", false);

// DeleteCookie("Teste");
  
// alert(GetCookie("Teste"));
//-->
</script>
