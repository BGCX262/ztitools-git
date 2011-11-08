<%
REM -------------------------------------------------------------------------
REM  /Connection.inc
REM -------------------------------------------------------------------------
REM Descricao   : ScripLet ASP Object create ConnectionString
REM Criacao     : 1/30/99 2:23AM
REM Local       : Brasilia/DF/Brazil
REM Elaborado   : Ruben Zevallos Jr. - zevallos@zevallos.com.br
REM Versao      : 1.0.0
REM Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM             : http://www.zevallos.com.br
REM -------------------------------------------------------------------------
REM The ZTITools.Connection ScriptLet is part of ZTITools 2.5 a complet tool
REM that search to abstract the HTML to ASP programers.
REM You can use the following code as you want, but we will glad if you 
REM notice us and/or save some information about us.
REM The author is not responsible for any damage or problem related with 
REM part or the entire source of ZTItools.Connection.
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM System constants
REM -------------------------------------------------------------------------
Const conConnSQL     = 1
Const conConnDBase   = 2
Const conConnAccess  = 3
Const conConnExcel   = 4
Const conConnFoxPro  = 5
Const conConnText    = 6
Const conConnParadox = 7
      
Const conConnDriverAccess   = 25 
Const conConnDriverdBaseIII = 21 
Const conConnDriverdBaseIV  = 277
Const conConnDriverdBase5   = 533
Const conConnDriverExcel3   = 534
Const conConnDriverExcel4   = 278
Const conConnDriverExcel5   = 22 
Const conConnDriverExcel7   = 22 
Const conConnDriverExcel97  = 790
Const conConnDriverFoxPro20 = 24 
Const conConnDriverFoxPro25 = 280
Const conConnDriverFoxPro26 = 536
Const conConnDriverParadox3 = 26 
Const conConnDriverParadox4 = 282
Const conConnDriverParadox5 = 538
Const conConnDriverText     = 27 

REM -------------------------------------------------------------------------
REM ConnectionString format
REM -------------------------------------------------------------------------
REM DSN	     = The name of the data source as returned by SQLDataSources. The 
REM            DSN keyword is not used if DRIVER is used.
REM DRIVER 	 = The name of the driver as returned by SQLDrivers. The DRIVER 
REM            keyword is not used if DSN is used.The SQL Server driver name
REM            is {SQL Server}.
REM SERVER	 = The name of the server on the network on which the data source 
REM            resides. On a Microsoft Windows NT computer, "(local)" can be 
REM            entered as the server, in which case a local copy of SQL Server
REM            can be used, even when this is a non-networked version. Note 
REM            that when the 16-bit SQL Server driver is using "(local)" without
REM            a network, the "MS Loopback Adapter" must be installed.
REM UID	     = The user login ID.
REM PWD	     = The user-specified password.
REM APP	     = The name of the application calling SQLDriverConnect (optional).
REM WSID	   = The workstation ID. Typically, this is the network name of the 
REM            computer on which the application resides (optional).
REM DATABASE = The name of the SQL Server database (optional).
REM LANGUAGE = The national language to be used by SQL Server (optional).
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Direct Data Access
REM -------------------------------------------------------------------------
REM Driver	            Keywords Required	Examples
REM -------------------------------------------------------------------------
REM Microsoft Access	  Driver, DBQ	      Driver={Microsoft Access Driver (*.mdb)};
REM                                       DBQ=c:\temp\sample.mdb
REM dBASE	              Driver, DriverID	Driver={Microsoft dBASE Driver (*.dbf)};
REM                                       DBQ=c:\temp; DriverID=277
REM Microsoft Excel 3.0	Driver, DriverID	Driver={Microsoft Excel Driver (*.xls)};
REM or 4.0                                DBQ=c:\temp; DriverID=278
REM Microsoft Excel	    Driver, DriverID, Driver={Microsoft Excel Driver (*.xls)};
REM  5.0/7.0            DBQ	              DBQ=c:\temp\sample.xls; DriverID=22
REM Microsoft Excel 97	Driver, DriverID, Driver={Microsoft Excel Driver (*.xls)};
REM                     DBQ               DBQ=c:\temp\sample.xls; DriverID=790
REM Microsoft FoxPro	  Driver, DriverID	Driver={Microsoft FoxPro Driver (*.dbf)};
REM                                       DBQ=c:\temp; DriverID=536
REM Paradox	            Driver, DriverID	Driver={Microsoft Paradox Driver (*.db )};
REM                                       DBQ=c:\temp; DriverID=26
REM Text	              Driver	          Driver={Microsoft Text Driver (*.txt;*.csv)};
REM                                       DefaultDir=c:\temp
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM DRIVERID	An integer ID for the driver.
REM -------------------------------------------------------------------------
REM 25        (Microsoft Access)
REM 21        (dBASE III)
REM 277       (dBASE IV)
REM 533       (dBASE 5.0)
REM 534       (Microsoft Excel 3.0)
REM 278       (Microsoft Excel 4.0)
REM 22        (Microsoft Excel 5.0/7.0)
REM 790       (Microsoft Excel 97)
REM 24        (Microsoft FoxPro 2.0)
REM 280       (Microsoft FoxPro 2.5)
REM 536       (Microsoft FoxPro 2.6)
REM 26        (Paradox 3.x)
REM 282       (Paradox 4.x)
REM 538       (Paradox 5.x)
REM 27        (Text)

%>