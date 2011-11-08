<script language=JavaScript>
<!--
// --------------------------------------------------------------------------
// Menu
// --------------------------------------------------------------------------
// Descricao   : Gerenciador de menus genericos
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

if (bName == "Microsoft Internet Explorer" && bVer >= "4.0")
  br = "n3";


else br = "n2";

var intLast;
var strLast = "";
  
var AHREF;
var ATarget;
var IMGWidth;
var IMGHeight;
var IMGRoot;

var Init = 1;

var IMGExt; 
var IMGUp;  
var IMGOver;
var IMGOut; 
var IMGDown;

var cmdover = new Array();
var cmdout = new Array();
var cmddown = new Array();
var cmdcounter = 0;

// --------------------------------------------------------------------------
// Inicializa o ambiente dos butoes
// --------------------------------------------------------------------------
// InitButton(IMGExt, IMGUp, IMGOver, IMGOut, IMGDown)
function InitButton() {
  var argv = InitButton.arguments;
  var argc = InitButton.arguments.length;
  IMGExt  = (argc > 0) ? argv[0] : ".jpg";
  IMGUp   = (argc > 1) ? argv[1] : "-Up";
  IMGOver = (argc > 2) ? argv[2] : "-Over";
  IMGOut  = (argc > 3) ? argv[3] : "-Out";
  IMGDown = (argc > 4) ? argv[4] : "-Down";
  
  IMGUp   = IMGUp + IMGExt;
  IMGOver = IMGOver + IMGExt;
  IMGOut  = IMGOut + IMGExt;
  IMGDown = IMGDown + IMGExt;

}


// --------------------------------------------------------------------------
// Adciona texto HTML dentro da area dos butoes
// --------------------------------------------------------------------------
function HTMLButton(strHTML) {
  document.write(strHTML);
}

// --------------------------------------------------------------------------
// Adciona item no menu
// --------------------------------------------------------------------------
function AddButton(IMGSrc, IMGAlt) {
  var argv = AddButton.arguments;
  var argc = AddButton.arguments.length;
  AHREF = (argc > 2) ? argv[2] : AHREF;
  ATarget = (argc > 3) ? argv[3] : ATarget;
  IMGWidth = (argc > 4) ? argv[4] : IMGWidth;
  IMGHeight = (argc > 5) ? argv[5] : IMGHeight;
  IMGRoot = (argc > 6) ? argv[6] : IMGRoot;

  if (Init == 1) {
    InitButton();
    
    Init = 2;
  }
  
    AddNewButton(IMGSrc, IMGAlt, AHREF, ATarget, IMGWidth, IMGHeight, IMGRoot);

    document.write("<BR>");

}

// --------------------------------------------------------------------------
// Adciona item no menu nova versao
// --------------------------------------------------------------------------
function AddNewButton(IMGSrc, IMGAlt) {
  var strStatus = IMGUp;
  var argv = AddNewButton.arguments;
  var argc = AddNewButton.arguments.length;
  AHREF = (argc > 2) ? argv[2] : AHREF;
  ATarget = (argc > 3) ? argv[3] : ATarget;
  IMGWidth = (argc > 4) ? argv[4] : IMGWidth;
  IMGHeight = (argc > 5) ? argv[5] : IMGHeight;
  IMGRoot = (argc > 6) ? argv[6] : IMGRoot;

  if (br == "n3") {
    cmdover[cmdcounter] = new Image();
    cmdover[cmdcounter].src = IMGRoot + IMGSrc + IMGOver;
  
    cmdout[cmdcounter] = new Image();
    cmdout[cmdcounter].src = IMGRoot + IMGSrc + IMGUp;

    cmddown[cmdcounter] = new Image();
    cmddown[cmdcounter].src = IMGRoot + IMGSrc + IMGDown;

    imgsrc = "<a href='" + AHREF + "'"
    
    if (ATarget != "null") {
      imgsrc = imgsrc + " target=" + ATarget
    }

    imgsrc = imgsrc + " name=acmd" + cmdcounter + "></a>"

    if (AHREF.substring(0, 1) != "/")
      AHREF = "/" + AHREF
      
    if (parent.Index != null){
      if ((parent.Index.location.pathname + parent.Index.location.search).toLowerCase() == (AHREF).toLowerCase() || (parent.Index.location.pathname + parent.Index.location.search).toLowerCase() == ("/" + AHREF).toLowerCase()){
        imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + IMGDown + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"
        strLast = "cmd" + cmdcounter;
        intLast = cmdcounter;
      }
      else
      {
        imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"
      }
    }
    else
    {
      imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"
    }
    document.write(imgsrc);

    cmdcounter++

  }
  else {
    imgsrc = "<a href='" + AHREF + "' target=" + ATarget + " name=acmd" + cmdcounter + ">"
    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + " border=0></a>"

    document.write(imgsrc);

  }
}


// --------------------------------------------------------------------------
// Gerencia o evento MouseOver
// --------------------------------------------------------------------------
function cmdOver() {
  if (br == "n3") {
    if (event.srcElement.tagName == "IMG") {
      cmdName = event.srcElement.name;
  
      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);
       
        if (cmdNumber != intLast){
          document[cmdName].src = cmdover[cmdNumber].src;
        }
      }
    }
    event.returnValue = false;
    event.cancelBubble = true;
  
  }
}

// --------------------------------------------------------------------------
// Gerencia o evento MouseOut
// --------------------------------------------------------------------------
function cmdOut() {
  if (br == "n3") {
    if (event.srcElement.tagName == "IMG") {
      cmdName = event.srcElement.name;
  
      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);

        if (cmdNumber != intLast){
          document[cmdName].src = cmdout[cmdNumber].src;
        }  
      }
    }
    event.returnValue = false;
    event.cancelBubble = true;
  }
}
// --------------------------------------------------------------------------
// Gerencia o evento MouseDown
// --------------------------------------------------------------------------
function cmdDown() {
  if (br == "n3") {
    if (event.srcElement.tagName == "IMG") {
      if (strLast != ""){
        document[strLast].src = cmdout[intLast].src;
      }      
      cmdName = event.srcElement.name;

      if (cmdName.substring(0, 3) == "cmd") {
        cmdNumber = cmdName.substring(3, 99);
        cmdAName = "a" + cmdName;
        strLast = cmdName;
        intLast = cmdNumber;

        document[cmdName].src = cmddown[cmdNumber].src;
        
        ahref = eval(cmdAName);

        if (ahref.target) {
          if (ahref.target == "__Top") {
            parent.location = ahref.href;
            
          } else {
            eval("parent." + ahref.target + ".location = ahref.href");
          
          }
        } else {
          parent.location = ahref.href;
        }
      }
    }
    event.returnValue = false;
    event.cancelBubble = true;
  }
}
  document.onmouseover = cmdOver;
  document.onmouseout  = cmdOut;
  document.onmousedown = cmdDown;
  document.onmouseup   = cmdOut;

//-->
</script>

