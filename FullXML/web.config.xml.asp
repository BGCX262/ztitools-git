<!-- <% Response.end %> -->
<configuration>
	<appSettings>
				
		<!-- Publication settings -->
		<key add="DEFAULT_PUBLICATION_STATE" value="2" /><!-- 0: offline, 1: in progress, 2: online-->
		<key add="MULTIPLE_WEBSITE" value="true" />
		<key add="USE_NEWPAGE_TEMPLATE" value="true" />
		
		<!-- Cache settings  -->
		<key add="USE_CACHE" value="false" />
		
		<!-- RC4 encrypt key -->
		<key add="RC4_KEY" value="1d2t8RGy1e1x5" />
								
		<!-- Data Folder settings -->		
		<key add="DATA_FOLDER" value="\db\" />
		<key add="PAGES_FOLDER" value="pages\" />
		<key add="USERS_FOLDER" value="users\" />
		<key add="LOGS_FOLDER" value="logs\" />
		<key add="STATS_FOLDER" value="stats\" />
		
		<!-- Media folder name -> used for upload and richmedia display -->
		<key add="MEDIA_FOLDER" value="media" />	
		
		<!-- Registry key to JMicrosoft ET Text driver settings -->
		<key add="JET_REGISTRY_KEY" value="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text" />
		
		<!-- Core folders name :: Should not be changed -->
		<key add="SKINS_FOLDER" value="skins"/>
		<key add="MODULES_FOLDER" value="modules"/>
		<!--<key add="ADMIN_FOLDER" value="\engine\admin\"/>-->
		
		<!-- XML Files Options -->
		<key add="XMLFILE_EXTENSION" value=".xml.asp"/>
		
		<!-- Index Files Name -->
		<key add="WEBSITES_INDEXFILE" value="websites"/>
		<key add="MODULES_INDEXFILE" value="modules"/>
		<key add="SKINS_INDEXFILE" value="skins"/>
		<key add="GROUPS_INDEXFILE" value="groups"/>
		<key add="USERS_INDEXFILE" value="users"/>
		<key add="LINKS_INDEXFILE" value="links"/>
		<key add="REDIRECTS_INDEXFILE" value="redirects"/>	
        	
		<!-- XML File name -->
		<key add="WEBSITE_FILE" value="website"/>
				
		<!-- External modules may have to add KEY NODES here -->
		<!-- Toolbar module -->
		<key add="TOOLBAR_FILE" value="toolbar"/>
		
		<!-- Polls module -->
		<key add="POLLS_FOLDER" value="polls\" />
		<key add="POLLS_INDEXFILE" value="polls"/>
	</appSettings>
	
	<permissions>
		<!-- Default permission level for each buildin groups, grouped by website model -->
		<GetDefautLevel>
			<model id="0">
				<group id="administrator" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="webmaster" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="member" level="CONST_ACCESS_LEVEL_VIEWER"/>
				<group id="anonymous" level="CONST_ACCESS_LEVEL_VIEWER"/>
			</model>
			<model id="1">
				<group id="administrator" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="webmaster" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="member" level="CONST_ACCESS_LEVEL_VIEWER"/>
				<group id="anonymous" level="CONST_ACCESS_LEVEL_VIEWER"/>
			</model>
			<model id="2">
				<group id="administrator" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="webmaster" level="CONST_ACCESS_LEVEL_ADMINISTRATOR"/>
				<group id="member" level="CONST_ACCESS_LEVEL_VIEWER"/>
				<group id="anonymous" level="CONST_ACCESS_LEVEL_DENIED"/>
			</model>
		</GetDefautLevel>
		
		<!-- list of authorized action/process for each level -->
		<ActionOnContent>
			
			<level id="2" name="contributor">
				<webform id="content_insert"/>
				<webform id="content_owner_update"/>
				
				<process id="do_insert_content"/>
			</level>
			
			
			<!--
			<level id="3" name="author">
				
				<process id="do_update_self_content"/>
				<process id="do_delete_self_content"/>
				<webform id="deletecontent"/>
			</level>
			
			
			<level id="4" name="moderator">
				<process id="do_update_content"/>
				<process id="do_delete_content"/>
				
				<process id="do_moveup_content"/>
				<process id="do_movedown_content"/>
				<process id="do_changebox_content"/>
				<process id="do_refresh_content"/>
				<webform id="editcontent"/>
				<webform id="deletecontent"/>
				<webform id="moveupcontent"/>
				<webform id="movedowncontent"/>
				<webform id="changebox"/>
			</level>
			<level id="5" name="administrator" >
				<process id="do_insert_content"/>
				<process id="do_update_content"/>
				<process id="do_delete_content"/>
				<process id="do_moveup_content"/>
				<process id="do_movedown_content"/>
				<process id="do_changebox_content"/>
				<process id="do_refresh_content"/>
				<webform id="editcontent"/>
				<webform id="deletecontent"/>
				<webform id="moveupcontent"/>
				<webform id="movedowncontent"/>
				<webform id="changebox"/>
			</level>
			-->
		</ActionOnContent>
				
		
		
	</permissions>
	
</configuration>