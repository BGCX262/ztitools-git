<%
	'-- GLOBAL QUERYSTRING WRAPPERS --------------------------------------------
	Public g_sServerName	: g_sServerName = Request.ServerVariables("SERVER_NAME")
	Public g_sScriptName	: g_sScriptName = Request.ServerVariables("SCRIPT_NAME")
	Public g_sURL			: g_sURL		= g_sScriptName & "?" & Request.QueryString
	Public g_sProcess		: g_sProcess	= IFF(len(Request("process"))>0, lcase(Request("process")), "")
	Public g_sWebform		: g_sWebform	= IFF(left(Request("webform"), 8)="webform_", lcase(Request("webform")), "webform_default")
	Public message			: message		= Request("message")
	Public g_sBaseUrl		: g_sBaseUrl	= "http://" & request.servervariables("SERVER_NAME") & request.servervariables("PATH_INFO")
	'---------------------------------------------------------------------------
	
	
	'-- RENDERING VARIABLES ----------------------------------------------------
	Public g_sPageID		: g_sPageID = Request.QueryString("pID")
	Public g_sMenuID		: g_sMenuID = Request.QueryString("mID")
	Public g_sContentID		: g_sContentID = IFF(len(Request.QueryString("cID"))>0, Request.QueryString("cID"), "")
	Public g_sSkin
	Public g_sTemplate
	Public g_sTheme	
	'---------------------------------------------------------------------------
	
	
	'-- XML VARIABLES ----------------------------------------------------------
	Public g_oWebSiteXML				' website.xml
	Public g_oWebPageXML				' webpage.xml
	Public g_oLanguagesXML				' languages.xml
	Public g_oLocalVersion				' VERSION
	
	
	'-- LOCALISATION VARIABLES -------------------------------------------------
	Public g_sCulture
	Public g_iLCID
	Public g_sEncoding	
	'---------------------------------------------------------------------------
	
	
	'-- APPLICATION VARIABLES NAME ---------------------------------------------
	CONST APPVAR_SEPARATOR	= "SEPARATOR_CHARACTER"
	CONST APPVAR_MAPPATH	= "SERVER_MAPPATH"
	CONST APPVAR_DOM_MODULES= "DOM_MODULES"
	CONST APPVAR_DOM_SKINS	= "DOM_SKINS"
	'---------------------------------------------------------------------------
	
	
	'-- MSXML3 PROGID
	CONST DOMDOCUMENT_PROGID				= "MSXML2.DOMDocument.3.0"
	CONST FreeThreadedDOMDOCUMENT_PROGID	= "MSXML2.FreeThreadedDOMDocument.3.0"
	CONST XSLTEMPLATE_PROGID				= "MSXML2.XSLTemplate.3.0"
	
	'-- MSXML4 PROGID
	'CONST DOMDOCUMENT_PROGID				= "MSXML2.DOMDocument.4.0"
	'CONST FreeThreadedDOMDOCUMENT_PROGID	= "MSXML2.FreeThreadedDOMDocument.4.0"
	'CONST XSLTEMPLATE_PROGID				= "MSXML2.XSLTemplate.4.0"
	'---------------------------------------------------------------------------
	
	
	'-- CREATE A GLOBAL FSO ----------------------------------------------------
	Dim g_oFSO
	Set g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
	'---------------------------------------------------------------------------
	
	
	'-- GET AND SET THE SERVER.MAPPATH VALUE -----------------------------------
	Dim g_sServerMappath
	If len(Application(APPVAR_MAPPATH))=0 then 
		g_sServerMappath = Server.MapPath(".")
		Application(APPVAR_MAPPATH) = g_sServerMappath
	Else
		g_sServerMappath = Application(APPVAR_MAPPATH)
	End If
	'---------------------------------------------------------------------------
	
	
	'-- TO ASSURE THAT THE WEB.CONFIG IS LOADED --------------------------------
	LoadWebConfig true
	'---------------------------------------------------------------------------
	
	
	'-- READ SOME VALUES FROM THE WEBCONFIG ------------------------------------
	DIM APPLICATION_NAME		: APPLICATION_NAME = appSettings("APPLICATION_NAME")	
	DIM APPLICATION_VERSION		: APPLICATION_VERSION = appSettings("APPLICATION_VERSION")
	
	DIM DEFAULT_PUBLICATION_STATE	: DEFAULT_PUBLICATION_STATE = cbool(appSettings("DEFAULT_PUBLICATION_STATE"))
	DIM USE_NEWPAGE_TEMPLATE		: USE_NEWPAGE_TEMPLATE = cbool(AppSettings("USE_NEWPAGE_TEMPLATE"))
	DIM USE_CACHE					: USE_CACHE = cBool(appSettings("USE_CACHE"))
	
	'-- THESE LINES DEFINE THE NAME OF SOME CORE FOLDER
	DIM ADMIN_FOLDER	: ADMIN_FOLDER		= "\engine\admin\"'appSettings("ADMIN_FOLDER")
	DIM MODULES_FOLDER	: MODULES_FOLDER	= g_sServerMapPath & "\" & appSettings("MODULES_FOLDER") & "\"
	DIM SKINS_FOLDER	: SKINS_FOLDER		= g_sServerMapPath & "\" & appSettings("SKINS_FOLDER") & "\"
	
	'-- name of the media folder (upload and display fro web, must be under the root)
	DIM MEDIA_FOLDER	: MEDIA_FOLDER	= appSettings("MEDIA_FOLDER")
	
	DIM DATA_FOLDER		: DATA_FOLDER	= CheckFileName(appSettings("DATA_FOLDER"))
	DIM PAGES_FOLDER	: PAGES_FOLDER	= appSettings("PAGES_FOLDER")
	DIM USERS_FOLDER	: USERS_FOLDER	= appSettings("USERS_FOLDER")
	DIM LOGS_FOLDER		: LOGS_FOLDER	= appSettings("LOGS_FOLDER")
	DIM STATS_FOLDER	: STATS_FOLDER	= appSettings("STATS_FOLDER")
		
	DIM XMLFILE_EXTENSION : XMLFILE_EXTENSION = appSettings("XMLFILE_EXTENSION")
	
	DIM WEBSITES_FILE : WEBSITES_FILE = appSettings("WEBSITES_INDEXFILE")
	DIM MODULES_FILE : MODULES_FILE = appSettings("MODULES_INDEXFILE")
	DIM SKINS_FILE : SKINS_FILE = appSettings("SKINS_INDEXFILE")
	DIM GROUPS_FILE : GROUPS_FILE = appSettings("GROUPS_INDEXFILE")
	DIM USERS_FILE : USERS_FILE = appSettings("USERS_INDEXFILE")
	DIM LINKS_FILE : LINKS_FILE = appSettings("LINKS_INDEXFILE")
	DIM REDIRECTS_FILE : REDIRECTS_FILE = appSettings("REDIRECTS_INDEXFILE")
	DIM WEBSITE_FILE : WEBSITE_FILE = appSettings("WEBSITE_FILE")	
	'---------------------------------------------------------------------------
		
		
	'-- The data files ---------------------------------------------------------
	Public Function website_xml
		website_xml = DATA_FOLDER & WEBSITE_FILE & XMLFILE_EXTENSION
	End Function	
	
	Public Function webpage_xml
		webpage_xml = DATA_FOLDER & PAGES_FOLDER & g_sPageID & XMLFILE_EXTENSION
	End Function
	
	Public Function skins_xml
		skins_xml = DATA_FOLDER & SKINS_FILE & XMLFILE_EXTENSION
	End Function
	
	Public Function modules_xml
		modules_xml = DATA_FOLDER & MODULES_FILE & XMLFILE_EXTENSION
	End Function
	
	Public Function groups_xml
		groups_xml = DATA_FOLDER & GROUPS_FILE & XMLFILE_EXTENSION
	End Function
	
	Public Function users_xml
		users_xml = DATA_FOLDER & USERS_FILE & XMLFILE_EXTENSION
	End Function
	
	Public Function links_xml
		links_xml = DATA_FOLDER & LINKS_FILE & XMLFILE_EXTENSION
	End Function
    
    Public Function redirects_xml
		redirects_xml = DATA_FOLDER & REDIRECTS_FILE & XMLFILE_EXTENSION
	End Function
    	
	Public Function modules_xml
		modules_xml = DATA_FOLDER & MODULES_FILE & XMLFILE_EXTENSION
	End Function
	
	Public Function upload_path
		upload_path = g_sServerMapPath & "\" & MEDIA_FOLDER
	End Function
	
	Public Function version_path
		version_path = g_sServerMapPath & "\VERSION"
	End Function
	
	Public Function languages_path
		languages_path = g_sServerMapPath & "\engine\languages.xml"
	End Function
	'----------------------------------------------------------------------------
				
	
	'-- ACCESS LEVEL -----------------------------------------------------------
	CONST CONST_ACCESS_LEVEL_DENIED = 0
	CONST CONST_ACCESS_LEVEL_VIEWER = 1
	CONST CONST_ACCESS_LEVEL_CONTRIBUTOR = 2
	CONST CONST_ACCESS_LEVEL_AUTHOR = 3
	CONST CONST_ACCESS_LEVEL_MODERATOR = 4
	CONST CONST_ACCESS_LEVEL_ADMINISTRATOR = 5
		
	Dim g_arrAccessLevel
	g_arrAccessLevel = Array("denied", "viewer", "contributor", "author", "moderator", "administrator")
	'---------------------------------------------------------------------------


	'-- WEBSITE MODEL ----------------------------------------------------------
	CONST CONST_WEBSITE_MODEL_PUBLIC = 0
	CONST CONST_WEBSITE_MODEL_SIGNUP = 1
	CONST CONST_WEBSITE_MODEL_PRIVATE = 2
	
	Dim g_arrModel : g_arrModel = array("publicmodel", "signupmodel", "privatemodel")
	'---------------------------------------------------------------------------
	
	
	'-- SUPPORTED EMAIL COM COMPONENTS -----------------------------------------
	Dim g_arrEmailCom : g_arrEmailCom = array("", "CDOSYS", "CDONTS", "Jmail", "AspEmail", "AspMail")
	'---------------------------------------------------------------------------


	'-- Load the doms of website and page -> init some global variables
	LoadDOMInMemory false
		
	'-- Load Modules
	LoadModulesInMemory false
		
	'-- Load skins
	LoadSkinsInMemory false
	
	
	'-- The current user -------------------------------------------------------
	Dim g_oUser
	Set g_oUser = New Cuser
	Do_Authentication_Cookie
	'---------------------------------------------------------------------------
%>