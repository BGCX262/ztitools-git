<script language=JavaScript>
<!--
// --------------------------------------------------------------------------
// Menu
// --------------------------------------------------------------------------
// Descricao   : Gerenciador de menus genericos com Image Splash Alive
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
var IMGSplash;
var IMGSplashExt;

var animationtimeout;

var splashdefault;

var cmdover = new Array();
var cmdout = new Array();
var cmddown = new Array();
var cmdsplash = new Array();
var cmdcounter = 0;

var cmdimage = new Array();
var cmdtype = new Array();
var cmdextention = new Array();
var animationlast = 0;

// --------------------------------------------------------------------------
// Inicializa o ambiente dos butoes
// --------------------------------------------------------------------------
// InitButton(IMGExt, IMGUp, IMGOver, IMGOut, IMGDown, IMGSplash)
function InitButton() {
  var argv = InitButton.arguments;
  var argc = InitButton.arguments.length;
  IMGExt  = (argc > 0) ? argv[0] : ".jpg";
  IMGUp   = (argc > 1) ? argv[1] : "-Up";
  IMGOver = (argc > 2) ? argv[2] : "-Over";
  IMGOut  = (argc > 3) ? argv[3] : "-Out";
  IMGDown = (argc > 4) ? argv[4] : "-Down";
  IMGSplash = (argc > 5) ? argv[5] : "-Splash";
  IMGSplashExt = (argc > 6) ? argv[6] : IMGExt;
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
    cmdout[cmdcounter] = new Image();
    cmdover[cmdcounter] = new Image();
    cmddown[cmdcounter] = new Image();
    cmdsplash[cmdcounter] = new Image();

    cmdimage[cmdcounter] = IMGRoot + IMGSrc;
    cmdtype[cmdcounter] = "";
    cmdextention[cmdcounter] = IMGExt;

    imgsrc = "<a href='" + AHREF + "'"

    if (ATarget != "null") {
      imgsrc = imgsrc + " target=" + ATarget
    }

    imgsrc = imgsrc + " name=acmd" + cmdcounter + "></a>"

    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + IMGExt + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + ">"

    if (AHREF.substring(0, 1) != "/")
      AHREF = "/" + AHREF

    if (parent.Index != null){
      if ((parent.Index.location.pathname + parent.Index.location.search).toLowerCase() == (AHREF).toLowerCase()){
        strLast = "cmd" + cmdcounter;
        intLast = cmdcounter;
      }
    }
    document.write(imgsrc);

    cmdcounter++

  }
  else {
    imgsrc = "<a href='" + AHREF + "' target=" + ATarget + " name=acmd" + cmdcounter + ">"
    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + IMGExt + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=cmd" + cmdcounter + " border=0></a>"

    document.write(imgsrc);

  }
}


// --------------------------------------------------------------------------
// Define o Splash padrao
// --------------------------------------------------------------------------
function AddSplashDefault(IMGSrc, IMGAlt) {
  var argv = AddSplashDefault.arguments;
  var argc = AddSplashDefault.arguments.length;
  IMGWidth = (argc > 2) ? argv[2] : IMGWidth;
  IMGHeight = (argc > 3) ? argv[3] : IMGHeight;
  IMGRoot = (argc > 4) ? argv[4] : IMGRoot;

  if (br == "n3") {
    splashdefault = new Image();
    splashdefault.src = IMGRoot + IMGSrc

    imgsrc = "<img src='" + IMGRoot + IMGSrc + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=splash>"

    document.write(imgsrc);

  }
  else {
    imgsrc = "<img src='" + IMGRoot + IMGSrc + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=splash></a>"

    document.write(imgsrc);

  }

}

// --------------------------------------------------------------------------
// Adciona item no menu nova versao
// --------------------------------------------------------------------------
function AddNewSplashButton(IMGSrc, IMGAlt) {
  var strStatus = IMGUp;
  var argv = AddNewSplashButton.arguments;
  var argc = AddNewSplashButton.arguments.length;
  AHREF = (argc > 2) ? argv[2] : AHREF;
  ATarget = (argc > 3) ? argv[3] : ATarget;
  IMGWidth = (argc > 4) ? argv[4] : IMGWidth;
  IMGHeight = (argc > 5) ? argv[5] : IMGHeight;
  IMGRoot = (argc > 6) ? argv[6] : IMGRoot;

  if (br == "n3") {
    cmdout[cmdcounter] = new Image();
    cmdover[cmdcounter] = new Image();
    cmddown[cmdcounter] = new Image();
    cmdsplash[cmdcounter] = new Image();

    cmdimage[cmdcounter] = IMGRoot + IMGSrc;
    cmdtype[cmdcounter] = "s";
    cmdextention[cmdcounter] = IMGExt;

    imgsrc = "<a href='" + AHREF + "'"

    if (ATarget != "null") {
      imgsrc = imgsrc + " target=" + ATarget
    }

    imgsrc = imgsrc + " name=ascmd" + cmdcounter + "></a>"

    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + IMGExt + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=scmd" + cmdcounter + ">"

    if (parent.Index != null){
      if ((parent.Index.location.pathname + parent.Index.location.search).toLowerCase() == ("/" + AHREF).toLowerCase()){
        strLast = "scmd" + cmdcounter;
        intLast = cmdcounter;
      }
    }
    document.write(imgsrc);

    cmdcounter++

  }
  else {
    imgsrc = "<a href='" + AHREF + "' target=" + ATarget + " name=acmd" + cmdcounter + ">"
    imgsrc = imgsrc + "<img src='" + IMGRoot + IMGSrc + strStatus + IMGExt + "' width=" + IMGWidth + " height=" + IMGHeight + " alt='" + IMGAlt + "' name=scmd" + cmdcounter + " border=0></a>"

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
          if (cmdout[cmdNumber].complete) {
            document[cmdName].src = cmdover[cmdNumber].src;
          }
        }

      }

      if (cmdName.substring(0, 4) == "scmd") {
        cmdNumber = cmdName.substring(4, 99);

        if (cmdNumber != intLast){
          if (cmdout[cmdNumber].complete) {
            document[cmdName].src = cmdover[cmdNumber].src;
          }
        }

        document.splash.src = cmdsplash[cmdNumber].src;
      }

      resetAnimation();
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
          if (cmdout[cmdNumber].complete) {
            document[cmdName].src = cmdout[cmdNumber].src;
          }
        }
      }

      if (cmdName.substring(0, 4) == "scmd") {
        cmdNumber = cmdName.substring(4, 99);

        if (cmdNumber != intLast){
          if (cmdout[cmdNumber].complete) {
            document[cmdName].src = cmdout[cmdNumber].src;
          }
        }

        document.splash.src = splashdefault.src;
      }

      resetAnimation();
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

        if (cmdout[cmdNumber].complete) {
          document[cmdName].src = cmddown[cmdNumber].src;

        }

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

      if (cmdName.substring(0, 4) == "scmd") {
        cmdNumber = cmdName.substring(4, 99);
        cmdAName = "a" + cmdName;
        strLast = cmdName;
        intLast = cmdNumber;

        if (cmdout[cmdNumber].complete) {
          document[cmdName].src = cmddown[cmdNumber].src;

        }

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

      resetAnimation();
    }
    event.returnValue = false;
    event.cancelBubble = true;
  }
}

var loadedtimeout;
var loadedcounter = 0;
var blnloaded;

// --------------------------------------------------------------------------
// Cancela a tentativa de ler todas as imagens
// --------------------------------------------------------------------------
function doLoadImages() {
  if (loadedcounter == 0 ) {
    blnloaded = true;

  }

  strobject = cmdtype[loadedcounter] + "cmd" + loadedcounter;

  if (!document[strobject] == null) {
    if (!document[strobject].complete) {
      blnloaded = false;

    }
  }

  if (++loadedcounter >= cmdcounter) {
    loadedcounter = 0;

    if (blnloaded) {
      window.clearTimeout(loadedtimeout);

      doLoadRestImages();

    }

  } else {
    window.clearTimeout(loadedtimeout);
    loadedtimeout = window.setTimeout("doLoadImages()", 100);

  }

}

// --------------------------------------------------------------------------
// Gerencia a carga automatica das imagens
// --------------------------------------------------------------------------
function doLoadRestImages() {

  for (x = 0; x < 4; x++) {
    for (i = 0; i < cmdcounter; i++) {
      strobject = cmdtype[i] + "cmd" + i;

      switch (x) {
        case 0 :
          cmdout[i].src = document[strobject].src;

          if (i == intLast) {
            cmddown[i].src = cmdimage[i] + IMGDown + cmdextention[i];
            document[strLast].src = cmddown[intLast].src;
          }

          break;

        case 1 :
          cmdover[i].src = cmdimage[i] + IMGOver + cmdextention[i];

          break;

        case 2 :
          cmddown[i].src = cmdimage[i] + IMGDown + cmdextention[i];

          break;

        case 3 :
          if (cmdtype[i] == "s") {
            cmdsplash[i].src = cmdimage[i] + IMGSplash + IMGSplashExt;

          }

      }
    }
  }

  resetAnimation();

}

var animationcounterx = 0;
var animationcounteri = 0;
var animationloadedtimeout;
var animationadder = 1;

// --------------------------------------------------------------------------
// Gerencia a animacao das opcoes
// --------------------------------------------------------------------------
function doAnimation() {

  if (cmdcounter > 0) {
    window.clearTimeout(animationtimeout);

    strobject = cmdtype[animationcounteri] + "cmd" + animationcounteri;

    strobject = cmdtype[animationcounteri] + "cmd" + animationcounteri;

  //  alert(strobject + "\nX=" + animationcounterx  + "\nI=" + animationcounteri  + "\nCounter=" + cmdcounter);

    blnsplash = false;

    blnslow = false;

    switch (animationcounterx) {
      case 0 :
        document[strobject].src = cmdover[animationcounteri].src;

  //  alert(document[strobject].src + "\n" + strobject + "\n" + animationcounterx  + "\n" + animationcounteri  + "\n" + cmdcounter);
  //  alert(cmdover[animationcounteri].src + "\n" + document[strobject].src);

        if (cmdtype[animationcounteri] == "s") {
          document.splash.src = cmdsplash[animationcounteri].src;

          blnsplash = true;

        }
        break;

      case 1 :
        document[strobject].src = cmddown[animationcounteri].src;

        if (cmdtype[animationcounteri] == "s") {
          document.splash.src = cmdsplash[animationcounteri].src;

          blnsplash = true;

        }
        break;

      case 2 :
        document[strobject].src = cmdout[animationcounteri].src;

        if (cmdtype[animationcounteri] == "s") {
          document.splash.src = cmdsplash[animationcounteri].src;

          blnsplash = true;

        }

        break;
    }

    animationcounteri = animationcounteri + animationadder;

    if (animationcounteri >= cmdcounter || animationcounteri < 0) {
      if (blnsplash) {
        document.splash.src = splashdefault.src;

      }

      if (++animationcounterx >= 3) {
        animationcounterx = 0;

        animationadder = animationadder * -1;

        blnslow = true;
      }

      if (animationadder == 1) {
        animationcounteri = 0;

      } else {
        animationcounteri = cmdcounter - 1;

      }
    }

    window.clearTimeout(animationloadedtimeout);

    if (blnslow) {
      animationloadedtimeout = window.setTimeout("doAnimation()", 5000);

    } else {
      animationloadedtimeout = window.setTimeout("doAnimation()", 700);

    }

  }
}

// --------------------------------------------------------------------------
// Inicializa a animacao
// --------------------------------------------------------------------------
function resetAnimation() {

  window.clearTimeout(animationloadedtimeout);
  window.clearTimeout(animationtimeout);

  if (animationcounteri != 0 || animationcounterx != 0) {
    for (i = 0; i < cmdcounter; i++) {
      strobject = cmdtype[i] + "cmd" + i;

      if (i != intLast){
        document[strobject].src = cmdout[i].src;
      }
      else {
        document[strobject].src = cmddown[i].src;
        }

      if (cmdtype[i] == "s") {
        document.splash.src = splashdefault.src;

      }
    }

    animationcounteri = 0;
    animationcounterx = 0;
  }

  animationtimeout = window.setTimeout("doAnimation()", 16000)

}

  document.onmouseover = cmdOver;
  document.onmouseout  = cmdOut;
  document.onmousedown = cmdDown;
  document.onmouseup   = cmdOut;

  window.onload   = doLoadImages;

//-->
</script>

