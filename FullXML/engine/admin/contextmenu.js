// Define global script variables
var bContextKey=false;

// The fnDetermine function performs most of the work


// Cette fonction renvoie la valeur de l'attribut "contextmenu" de l'element html qui a ete clicke
function fnGetContextID(el) {
  while (el!=null) {
    if (el.contextmenu) return el.contextmenu
    el = el.parentElement
  }
  return ""
}

// Cette fonction envoie la valeur de l'attribut "ID" de l'element qui a ete clicke droit.
// cet element doit avoir un attribut "contextmenu"
function fnGetTarget(el) 
{
	while (el!=null)
	{
		if (el.id && el.contextmenu) 
			return el.id;
		el = el.parentElement;
	}
	return "";
}


// Cette fonction envoie l'element qui a ete clicke droit.
// cet element doit avoir un attribut "contextmenu"
function fnGetTargetObject(el) 
{
	while (el!=null)
	{
		if (el.id && el.contextmenu) 
			return el;
		el = el.parentElement;
	}
	return null;
}


// This function return the actual menuitem
function fnGetMenuItem(el) 
{
	while (el!=null)
	{
		if (el.id && el.component=="menuitem") 
			return el;
		el = el.parentElement;
	}
	return null;
}


function fnDetermine()
{
	oWorkItem=event.srcElement;
	//oWorkItem = fnGetMenuItem(event.srcElement);
			
  // Proceed if the desired keyboard key is pressed.
  if(bContextKey==true){    
    // If the menu STATUS is false, continue.
    if(oContextMenu.getAttribute("status")=="false"){
      // Give the menu mouse capture so it can interact better with the page.
      oContextMenu.setCapture();

     // Relocate the menu to an offset from the mouse position.
     //oContextMenu.style.top=event.clientY + document.body.scrollTop + 1;
     //oContextMenu.style.left=event.clientX + document.body.scrollLeft +  1;
               

      oContextMenu.innerHTML="";
      // Set its STATUS to true.
      var sContext = fnGetContextID(event.srcElement)
      if (sContext!="") {
        fnPopulate(sContext, fnGetTargetObject(oWorkItem))
        oContextMenu.setAttribute("status","true");
        event.returnValue=false;
      }
      else
        event.returnValue=true;
        
        // ---------------------------------------------
        //Correct positionning of the menu        
	    //alert(document.body.scrollTop);
	    
	    //mouse position
	    var y = event.clientY; //+ document.body.scrollTop + 1;
	    var x = event.clientX; //+ document.body.scrollLeft + 1;
	    
	    //body size
	    var L	= document.body.clientWidth;
	    var H	= document.body.clientHeight;
	      
	    //Menu size
	    var ll = 200;
	    var hh = oContextMenu.style.height;
	    
	    hh = eval(hh.substring(0, hh.indexOf('px')));
	    	    
	    
	    //adjust x position
	    if ((x+ll) < L)
	    	oContextMenu.style.left = x + 'px';
	    else
	    	oContextMenu.style.left = (L-ll+document.body.scrollLeft) + 'px';
	    	    
	    
	    //adjust y position
	    if (hh + y < H)
	    	oContextMenu.style.top = y + 'px';
	    else
	    	oContextMenu.style.top = (H-hh+document.body.scrollTop) + 'px';
	    
	    	
	    /*
	    //adjust x position
	    if ((x+ll) < L)
	    	oContextMenu.style.left = x + 'px';
	    else
	    	oContextMenu.style.left = (L-ll) + 'px';
	    	    
	    
	    //adjust y position
	    if (hh + y < H)
	    	oContextMenu.style.top = y + 'px';
	    else
	    	oContextMenu.style.top = (H-hh) + 'px';
	    */
	    
	    // ---------------------------------------------
    }
  }
  else{
    
    
	// If the keyboard key was not pressed and the menu status is true, continue.
    if(oContextMenu.getAttribute("status")=="true")
	{
      	// try to get the SPAN that is clicked in the menu 
		oWorkItem = fnGetMenuItem(event.srcElement);
		if(oWorkItem && oWorkItem.parentElement.id=="oContextMenu" && oWorkItem.getAttribute("component")=="menuitem")
	  		fnFireContext(oWorkItem);
      
		// Reset the context menu, release mouse capture, and hide it.  
		oContextMenu.style.display="none";
		oContextMenu.setAttribute("status","false");
		oContextMenu.releaseCapture();
		oContextMenu.innerHTML="";
		event.returnValue=false;
    }
  }
}
    
// create the content of the menu
// note if a span has a disabled attribute, it looks disabled !
// i also use the class 'checked' to disaply BOLD an pre-selected item
function fnPopulate(sID, srcObj)
{
	var srcId = srcObj.id;
	var str=""
	var elMenuRoot = document.all.contextDef.XMLDocument.childNodes(0).selectSingleNode('contextmenu[@id="' + sID + '"]')
  if (elMenuRoot) {
    for(var i=0;i<elMenuRoot.childNodes.length;i++)
    {
      switch(elMenuRoot.childNodes[i].getAttribute("type"))
      {
		case "title":
			str+='<big>'+elMenuRoot.childNodes[i].getAttribute("value")+'</big>';
			break;
      
		case "separator":
			str+='<div align=right><table cellspacing=0 cellpadding=0 class=separator><tr><td height=1></td></tr></table></div>';
			break;
			
		case "content":
			var icon = "&nbsp;";
			var disabled = "";
			
			// in the case of the moves entry:
			// disable moves that target objet forbig
			if (elMenuRoot.childNodes[i].getAttribute("cmd")=="moveupcontent" && srcObj.moveup=="false") 
				disabled = " disabled='true'";

			if (elMenuRoot.childNodes[i].getAttribute("cmd")=="movedowncontent" && srcObj.movedown=="false") 
				disabled = " disabled='true'";


			// in the case of the box change menu entry:
			// if the box node of the menu is equal to the box node of the clicked content, we display it as checked
			if (elMenuRoot.childNodes[i].getAttribute("cmd")=="changebox" && elMenuRoot.childNodes[i].getAttribute("id")==srcObj.box) 
				icon = "<img " + disabled + " src=engine/admin/media/checked.png>&nbsp;";
			
			if (elMenuRoot.childNodes[i].getAttribute("icon"))
				icon = "<img " + disabled + " src=" + elMenuRoot.childNodes[i].getAttribute("icon")  + ">&nbsp;";
				
				
			

			str+='<span pageid="' + elMenuRoot.getAttribute("pageid") + '" component="menuitem" ' + ' elmenuid="' + elMenuRoot.childNodes[i].getAttribute("id") + '" ' +
				' websiteid="' + elMenuRoot.getAttribute("websiteid") + '" ' +
				' contentid="' + srcId + '" ' +
				' cmd="' + elMenuRoot.childNodes[i].getAttribute("cmd") + '" ' +
				' module="' + elMenuRoot.childNodes[i].getAttribute("module") + '" ' +
				' box="' + elMenuRoot.childNodes[i].getAttribute("id") + '" ' +
				' id=oMenuItem' + i + ' class=""' + disabled + '>';
				
				str += "<table width=100% cellspacing=0 cellpadding=0><tr><td class=icon>" + icon + "</td><td class=text>" + elMenuRoot.childNodes[i].getAttribute("value")  + "</td></tr></table>" +  "</span><BR>";
			break;
			
		case "placeholder":
			var icon = "<img width=16 height=16 src=modules/" + elMenuRoot.childNodes[i].getAttribute("module") + "/media/contenttype_" + elMenuRoot.childNodes[i].getAttribute("contenttype") + ".png>";
			str+='<span pageid="' + elMenuRoot.getAttribute("pageid") + '" component="menuitem" ' + ' elmenuid="' + elMenuRoot.childNodes[i].getAttribute("id") + '" ' +
				' websiteid="' + elMenuRoot.getAttribute("websiteid") + '" ' +
				' placeholder="' + srcId + '" ' +
				' cmd="' + elMenuRoot.childNodes[i].getAttribute("cmd") + '" ' +
				' module="' + elMenuRoot.childNodes[i].getAttribute("module") + '" ' +
				' contenttype="' + elMenuRoot.childNodes[i].getAttribute("id") + '" ' +
				' class=""' +
				' id=oMenuItem' + i + '>'+  
				"<table width=100% cellspacing=0 cellpadding=0><tr><td class=icon>" + icon + "</td><td class=text>" + elMenuRoot.childNodes[i].getAttribute("value")  + "</td></tr></table>" +  "</span><BR>";
			
				//' id=oMenuItem' + i + '>' + elMenuRoot.childNodes[i].getAttribute("value") + 
				//"</span><BR>";
			break;
		
		case "page":
			str+='<span pageid="' + elMenuRoot.getAttribute("pageid") + '" component="menuitem" ' + 'elmenuid="' + elMenuRoot.childNodes[i].getAttribute("id") + '" ' +
				' srcId="' + srcId + '" ' +
				' projectid="' + elMenuRoot.getAttribute("projectid") + '" ' +
				' cmd="' + elMenuRoot.childNodes[i].getAttribute("cmd") + '" ' +
				' module="' + elMenuRoot.childNodes[i].getAttribute("module") + '" ' +
				' class=""' +
				' id=oMenuItem' + i + '>'+  
				"<table width=100% cellspacing=0 cellpadding=0><tr><td class=icon>" + icon + "</td><td class=text>" + elMenuRoot.childNodes[i].getAttribute("value")  + "</td></tr></table>" +  "</span><BR>";
			break;
		
      }      
    }
    oContextMenu.innerHTML=str;
    oContextMenu.style.display="block";
    oContextMenu.style.pixelHeight = oContextMenu.scrollHeight    
  }
}

function fnChirpOn(){
  if((event.clientX>0) &&
     (event.clientY>0) &&
     (event.clientX<document.body.offsetWidth) &&
     (event.clientY<document.body.offsetHeight)){
    oWorkItem = fnGetMenuItem(event.srcElement);
	if (oWorkItem)
	  oWorkItem.className = "selected";
    
  }
}
function fnChirpOff(){
  if((event.clientX>0) &&
     (event.clientY>0) &&
     (event.clientX<document.body.offsetWidth) &&
    (event.clientY<document.body.offsetHeight)){
    oWorkItem = fnGetMenuItem(event.srcElement);
	if (oWorkItem)
	  oWorkItem.className = "";
    
  }
}

function fnInit(){
  if (oContextMenu) {
    //oContextMenu.style.height=document.body.offsetHeight/2;
    oContextMenu.style.zIndex=2;
    // Setup the basic styles of the context menu.
    document.oncontextmenu=fnSuppress;
  }
}

function fnInContext(el) {
  while (el!=null) {
    if (el.id=="oContextMenu") return true
    el = el.offsetParent
  }
  return false
}

function fnSuppress(){
  if (!(fnInContext(event.srcElement))) { 
    oContextMenu.style.display="none";
    oContextMenu.setAttribute("status","false");
    oContextMenu.releaseCapture();
    bContextKey=true;
  }

  fnDetermine();
  bContextKey=false;
}


// Customize this function based on your context menu
function fnFireContext(oItem)
{	
	// if the span element has a disabled attribute, then we're not executig the action
	if (oItem.disabled) return;

	switch (oItem.cmd)
	{
	case "source":
		location.href = "view-source:" + location.href;
		break;
	
	case "insertcontent":
		window.open('popup.asp?webform=webform_insert_content&pID=' + oItem.pageid + '&contenttype=' + oItem.contenttype + '&placeholder=' + oItem.placeholder, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
		
	case "editcontent":
		window.open('popup.asp?webform=webform_update_content&pID=' + oItem.pageid + '&id=' + oItem.contentid, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;

	case "deletecontent":
		window.open('popup.asp?webform=webform_delete_content&pID=' + oItem.pageid + '&id=' + oItem.contentid, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
	
	case "refreshcontent":
		window.open('popup.asp?process=Do_Refresh_content&sid=' + oItem.websiteid + '&pID=' + oItem.pageid + '&id=' + oItem.contentid, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
	
	case "moveupcontent":
		window.open('popup.asp?webform=webform_moveup_content&sid=' + oItem.websiteid + '&pID=' + oItem.pageid + '&contentID=' + oItem.contentid, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
		
	case "movedowncontent":
		window.open('popup.asp?webform=webform_movedown_content&sd=' + oItem.websiteid + '&pID=' + oItem.pageid + '&contentID=' + oItem.contentid, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
		
	case "changebox":
		window.open('popup.asp?webform=webform_changebox_content&sid=' + oItem.websiteid + '&pID=' + oItem.pageid + '&contentID=' + oItem.contentid + '&box=' + oItem.box, 'fx4popup', 'width=730, height=400, scrollbars=0, status=0, resizable=1');
		break;
	
	/*		
	case "removebox":
		RemoveBox(oItem.contentid, oItem.menuid);
		break;
	
	// --> the page contextual menu
	case "editpage":
		document.location = 'VisualPortalManager/default.asp?projectID=' + oItem.projectid + '&action=editpage&id=' + oItem.pageid;
		break
		
	case "edittemplate":
		document.location = 'VisualPortalManager/default.asp?projectID=' + oItem.projectid + '&action=edittemplate&id=' + oItem.pageid;
		break		
	// <-- 
		
	*/
		
	case "back":
		history.back()
		break;
		
	
	}
}
