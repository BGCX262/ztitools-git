<%
	const SUCCESS	= "SUCCESS"
	const INFO		= "INFO"
	const WARNING	= "WARNING"
	const ERROR		= "ERROR"
	const FATAL		= "FATAL ERROR"
	
	'----------------------------------------------------
	'-- Write a message into the log file of the application
	'----------------------------------------------------
	Sub LogEvent(pLevel, pAppName, pTitle, pMessage)
		'-- error trapping
		on error resume next
		
		Dim test
		Set test = New LogFile
			test.TemplateFileName = server.MapPath("Logs\%y-%m-%d.log")
			test.Log(array(pTitle, pMessage))
		Set test = nothing
		
		'-- clear and remove the error trapping
		if err<>0 then
			err.Clear
		end if
		on error goto 0
	End Sub
	
	
	'------------------------------------------------------------------------------------------
	' cette classe permet de gerer un fichier de log quotidien, mensuel ou annuel du genre IIS.
	' On peut par ailleurs parametrer le chemin du fichier, et le separateur des champs 
	'------------------------------------------------------------------------------------------
	Class LogFile
	
		Private m_oFSO
		Private m_oOutStream
		Private m_sFilename
		Private m_sLogSeparator
		
		Private m_arrColumns

		'------------------------------------------------------------------------------
		'  CONSTRUCTEUR : instance de FSO
		'------------------------------------------------------------------------------
		Private Sub Class_Initialize()
		  	Set m_oFSO = Server.CreateObject("Scripting.FileSystemObject") 
			m_sLogSeparator = Chr(9)  '' par defaut une tabulation
			m_arrColumns = null
		End Sub
		
		'------------------------------------------------------------------------------
		' DESSTRUCTEUR : supprimes les objets
		'------------------------------------------------------------------------------
		Private Sub Class_Terminate()
			Set m_oFSO = Nothing
			Set m_oOutStream = Nothing
		End Sub
	
	
		' Affecte un template de nom de fichier de Log :
		' %y : année, %m, %d
 		Property Let TemplateFileName (sValue)
 			
 			sValue = Replace(sValue, "%d", Right("0" & Day(Now), 2)) 
 			sValue = Replace(sValue, "%m", Right("0" & Month(Now), 2))
 			sValue = Replace(sValue, "%y", Year(Now))
 			m_sFilename = sValue
 			
 			
 			Dim addcols
 			addcols = false
 			
 			'-- check for first line
 			if IsArray(m_arrColumns) and (Not m_oFSO.FileExists(m_sFilename)) Then
				addcols = true
			end if
 			
 			'-- on ouvre/crée le fichier
 			on error resume next
 			Set m_oOutStream = m_oFSO.OpenTextFile(m_sFilename, 8, True)
 			if err.number<>0 then
 				Response.Write "Administrator, check that the data's folder (and its subfolders) have the good security settings. Please refer to the FX4 quick start guide for more detailed informations.<br>"
 				Response.write "{system error message: " & err.Description & "}"
 				err.Clear
 				response.End
 			end if
 			
 			on error goto 0
 			
 			'-- on ajoute la premiere ligne, si besoin
 			if addcols then
 				Dim i, tmp
 				tmp = m_arrColumns(lbound(m_arrColumns))
 				For i = lbound(m_arrColumns)+1 to ubound(m_arrColumns)
					tmp = tmp & m_sLogSeparator & replace(replace(m_arrColumns(i), vbCr, ""), vblf, "")
				Next
			
				m_oOutStream.WriteLine tmp
 			end if
 			
 			
 		End Property
 		
 		' Recupere le séparateur de champ
 		Property Let FieldSeparator(sValue)
 			m_sLogSeparator = sValue
 		End Property
 		
 		'Affecte le séparateur de champs
 		Property Get FieldSeparator()
 			FieldSeparator = m_sLogSeparator
 		End Property

		' recuepere le nom du fichier de log
		Property Get FileName()
			FileName = m_sFilename
		End Property
		
		Property Let Columns(arrayNames)
			m_arrColumns = arrayNames
		End Property

		' Ecrit une ligne dans le fichier de log
		Public Sub LogLine(sText)
			m_oOutStream.WriteLine Now() & m_sLogSeparator & replace(sText, vbcrlf, " ")
		End Sub
		

		'--------------------------------------------
		' -- Write an array of value as a new line --
		'--------------------------------------------
		Public Sub Log(arrayText)
			Dim tmp, i			
			tmp = Now() 
			
			For i = lbound(arrayText) to ubound(arrayText)
				tmp = tmp & m_sLogSeparator & replace(replace(arrayText(i), vbCr, ""), vblf, "")
			Next
			
			m_oOutStream.WriteLine tmp
		End Sub
		
	'	Private Sub CheckFileExistence()
	'		
	'	End Sub
		
	End Class
%>