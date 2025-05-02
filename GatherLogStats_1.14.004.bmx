' -----------------------------------------------------------------------------------------------------------------
'BlitzMax program : 10 Feb 2014 11:50:50
'Template         : GatherLogStats.bmx
'
' last updated 25/05/2017
'
'	11/07/2016	1.10.002		Add date_time_idx2
'	26/07/2016	1.10.003		Swap sort order of 'Published counts by the hour' and increase to 48 hours
'											Increase daily count from 7 to 14 days. 
'	28/07/2016	1.10.004		Add code to generate output for graphs.
'											Replace Instr() with strFind()
'	10/08/2016	1.10.005
'	06/09/2016	1.10.006		Change Publication report to HTML format
'	08/09/2016	1.10.007		Reformat Publication report
'											Cleanup code
'	21/09/2016	1.10.008		Update count report format
'											Move Reports to \Reports folder
'	24/10/2016	1.10.009		Change sequence of report of count by server
'	11/11/2016	1.10.010		Added code to create new .csv files to be fed to dashboard
'	01/12/2016	1.10.011		Update gen_data_files() to properly create a result when no data available for given times
'	16/12/2016	1.11 			Switch updates to transaction "db_queueCommand()" instead of writing out to csvfile.
'											Update Publication_Report.html. Add Error count to report.
'	09/01/2017	1.12				Add new table for file writes. 
'											Combine Success & Errors into one table, add status flag to publications
'											Move program to bin and config file to conf	
'	15/05/2017	1.12.001		Bug fix, getting filsize on live files before they were opened meant size may be wrong and cause it to loop.
'	25/05/2017	1.12.002		Add system table to store datestamp of last cleanup and last reindex
'											Add code to perform reindex
'											Add paramter to config to specify how often to perfrom reindex
'	05/06/2017	1.12.003		Bug fix, t_output.WriteOutput(outpath) was being called twice for Pub error counts. 
'	19/06/2017	1.12.004		Bug fix, DB_CleanupDB() was not clear old dates from published table
'	21/06/2017	1.12.005		Bug fix, mismatch in doc_name & file name
'											Write db error messages to log files
'	15/08/2019	1.12.006	
'	15/10/2019	1.13			Update daily total to get totals for last 14 days.	
'	11/03/2020	1.14			Reformat tables, remove date. Change starttime & endtime format from ddmmss to YYYYDDMMhhmmss
'	02/04/2020	1.14.001		Extract time taken to publish new documents, send details to dashboard in file pub_perf_all.csv
'	06/04/2020	1.14.002		Bug fix, exclude duration <0 for pub_perf_all.csv
'	09/02/2021	1.14.003
'	02/05/2025	1.14.004		Reinstate CleanupDB()		
'
'
' -----------------------------------------------------------------------------------------------------------------
SuperStrict

Import MaxGui.Drivers
Import MaxGUI.ProxyGadgets
Import tm.sqlite
Include "..\..\..\..\..\public\bmx_extensions.bmx"
Include "..\..\..\..\..\public\bmx_gui_extensions.bmx"
Include "..\..\..\..\..\public\bmx_t_logfile.bmx"
Include "..\..\..\..\..\public\bmx_SQLite_extensions.bmx"   'Use with tm.sqlite
Global recsPerTransaction:Int = 250
db_commit_buffer:Int = recsPerTransaction * 512
db_max_queue_items:Int = recsPerTransaction
db_header_on = False
' ----------------------------------------------------------------------------------
' Global variables & constants
' ----------------------------------------------------------------------------------
AppTitle ="Gather Log Stats"
Const progname:String = "GatherLogStats"
Const version:String = "1.14.004"
Const builddate:String = "02/05/2025"

Global rc:Int
Global trace:Int = False

Global rebuild_db:Int = False
Global refresh_rows:Int = False
Global days2load:Int = 0
Global days2keep:Int = 60
Global ServerList:String[] '= ["CFR2WEBPUB02","CFR2WEBPUB03","CFR2WEBPUB04","CFR2WEBPUB05"]
Global sourcePath:String '= "C:\Databases\Report2Web\Router_logs\CFR2WEBPUB03\"
Global savePath:String '= sourcePath+"save\"
Global db_path:String ' = "database\Report2Web_stats.db"
Global db_copy_path:String
Global read_file_path:String = "c:\temp\tempfile.txt"
Global user_wait:Int = False;
Global hide_window:Int = False;
Global reports_path:String = "..\Reports\"
Global data_path:String = "..\Reports\"
Global reindex_flags:String = Null
'Local logpath:String = "\c$\Program Files (x86)\Report2Web Client Tools\Logs\"
'Local logpath_saved:String = "\d$\ControlD_RouterLog\"
Global logpath_current:String = "\c$\Program Files (x86)\Report2Web Client Tools\Logs\"'"\d$\ControlD_RouterLog_temp\"
'Global filecount:Int = 0
Global files:String[]
Global inpath:String
Global bytessize:Float  '= FileSize(location)			
Global bytesCount:Float '= 0
Global command:String
Global rows:String[]
Global seekSize:Long
Global records_added:Int = False
Load_settings()

' ----------------------------------------------------------------------------------
' Open log file
' ----------------------------------------------------------------------------------
t_logfile.set_logpath("logs\gatherLogStats",,10) ',True,10)
t_logfile.set_CleanupInterval(5)
't_logfile.displayON()
t_logfile.writelog("Starting")
'Global csvfile:TStream
' ----------------------------------------------------------------------------------
' Open GUI
' ----------------------------------------------------------------------------------
Global WindowMain:TGadget
Global teaTextArea0:TGadget
Global teaTextArea1:TGadget
Global btnPROCESS:tgadget
Global btnCANCEL:tgadget
Global progbar:tgadget
Global progtext:tgadget
Global panel1:tgadget
'Global label1:tgadget
WindowMain=CreateWindow("GatherLogStats version "+version,0,0,800,400,Desktop(),WINDOW_TITLEBAR|WINDOW_RESIZABLE|WINDOW_CENTER)
	If hide_window = True
		MinimizeWindow WindowMain
	End If
	panel1 = CreatePanel(0,0,ClientWidth(windowmain),ClientHeight(windowmain)-40,windowmain,PANEL_GROUP)
	SetGadgetLayout panel1,1,1,1,1
	'teaTextArea0=CreateTextArea(4,4,ClientWidth(panel1)-8,ClientHeight(panel1)-120,panel1,0)
	'SetTextAreaColor teaTextArea0,255,255,255,1
	'SetTextAreaColor teaTextArea0,0,0,0,0
	'SetGadgetLayout teaTextArea0,1,1,1,1
	'SetGadgetFont(teaTextArea0, LookupGuiFont(GUIFONT_MONOSPACED,12) )
	'teaTextArea1=CreateTextArea(4,GadgetHeight(teaTextArea0)+8,ClientWidth(panel1)-8,ClientHeight(panel1)-GadgetHeight(teaTextArea0),panel1,0)
	teaTextArea1=CreateTextArea(4,4,ClientWidth(panel1)-8,ClientHeight(panel1),panel1,0)
	SetTextAreaColor teaTextArea1,255,255,255,1
	SetTextAreaColor teaTextArea1,0,0,0,0
	SetGadgetLayout teaTextArea1,1,1,0,1
	SetGadgetFont(teaTextArea1, LookupGuiFont(GUIFONT_MONOSPACED,10) )

	progbar = CreateProgBar(10, ClientHeight(WindowMain) - 30, 250, 20, WindowMain)
	SetGadgetLayout progbar,1,0,0,1
	'label1=CreateLabel("",220,ClientHeight(windowMain)-30,200,20,windowmain) ',LABEL_FRAME)
	'SetGadgetLayout label1,1,0,0,1
	progtext = CreateLabel("", 270, ClientHeight(WindowMain) - 30, 350, 20, WindowMain)', LABEL_FRAME)
	
	btnCANCEL = CreateButton("Exit", ClientWidth(WindowMain) - 100, ClientHeight(WindowMain) - 30, 80, 24, WindowMain)
	SetGadgetLayout btnCANCEL,0,1,0,1

' ----------------------------------------------------------------------------------
' Load settings
' ----------------------------------------------------------------------------------
If db_path = "" 
	'Notify("DB_PATH not defined")
	End
End If
'DebugStop
message(1,"*** Starting ****")
CloseFile(WriteFile("gatherlogstats_started.txt"))
message(1,"db_path = "+db_path)
message(1,"reports_path = "+reports_path)
message(1,"data_path = "+data_path)
' ----------------------------------------------------------------------------------
' Re-Build database if required
' ----------------------------------------------------------------------------------
If rebuild_db = True
	'End
	message(1,"Delete Database")
	refresh_rows = True
	DeleteFile(db_path)
EndIf
If FileType(db_path) = 0
	rebuild_db = True
	t_database.BuildDB()
EndIf
' ----------------------------------------------------------------------------------
' Open Database and start processing
' ----------------------------------------------------------------------------------
rc = db_openDB(db_path,db_pntr)
If rc <> 0
	t_logfile.writeLog("Error opening database")
	End
Else
	' ----------------------------------------------------------------------------------
	' Read log files and update statistics.
	'Update_Stats()
	' ----------------------------------------------------------------------------------
	' Clean out old entries based on days2keep value
	If days2keep > 0
		t_database.CleanupDB(days2keep)
	End If
	' ----------------------------------------------------------------------------------
	' Generate reports and data files for Dashboard
	gen_reports()
	Gen_data_files()
	' ----------------------------------------------------------------------------------
	' Database house keeping such as reindexing to improve performance
	HouseKeeping()
	' ----------------------------------------------------------------------------------
	' Close database and exit 
	message(1,copies("-",80))
	message(1, "*** Completed ***")
	db_closedb(db_pntr)
End If
t_logfile.writelog("Completed")
If user_wait
	wait()
End If
DeleteFile("gatherlogstats_started.txt")
FreeGadget windowmain

End

Function Update_Stats()
	'If newvals[0] <> "null"	
	Local searchdate:String
	Local current_date:String = GetCurrentDate("DDMMYYYY")
	Local logdate:String
	Local starttime:String
	Local endtime:String
	Local docname:String
	Local write_filename:String
	Local server:String
	Local detail:String
	Local bytesSize:Long 
	Local bytesCount:Long 
	Local status:Int 
	For Local i:Int = -days2load To 0 
		searchDate=CalcNewDate(current_Date,i)
		logdate = ReformatDate(CalcNewDate(GetCurrentDate("DDMMYYYY"),i),"DDMMYYYY","YYYY-MM-DD") 		
		Local log_filename:String = "Burster-"+logdate+".txt"
		For Local server:String = EachIn ServerList
			Local log_source_path:String
			Local date:String 
			log_source_path = "\\"+server+logpath_current+log_filename
			'Print log_source_path
			'message(1,"Processing "+log_source_path)
			If FileType(log_source_path)=1
				message(1, "Opening " + log_source_path)
				Local infile:TStream = ReadFile(log_source_path)
				If infile = Null
					message(1,"  *** Error opening file ***")
				Else
					' =================================================================================================================
					bytesSize = FileSize(log_source_path)			
					bytesCount = 0
					' ----------------------------------------------------------------------------------
					' Retrieve log file size, delete deatils if required
					' ----------------------------------------------------------------------------------
					date = reformatdate(searchDate,"DDMMYYYY","YYYYMMDD")
					seekSize = 0
					command = "SELECT size_bytes FROM log_size WHERE DATE = '"+date+"' and SERVER = '"+server+"';"
					'message(1,command)
					rows = t_database.QueryDB(command, rc)
					'message(1,"rc="+rc)
					If rc=False
						message(1,db_last_message)
						t_logfile.writelog(db_last_command)
						t_logfile.writelog(db_last_message)
						End
					Else
						If Len(rows)=0
								command = "INSERT INTO log_size VALUES ('"+server+"','"+date+"',0);"
								'message(1,command)
								t_database.QueryDB(command,rc)
								'message(1,"rc="+rc)
								If rc=False
									message(1,db_last_message)
									t_logfile.writelog(db_last_command)
									t_logfile.writelog(db_last_message)
									End
								EndIf
						Else
							'message(1,rows[0])
							SeekSize = Long(getString(rows[0],"","|"))
						End If
					End If
					If Seeksize>BytesSize
						message(1,"Size mismatch")
						t_logfile.writelog("Size mismatch on "+log_source_path)
						SeekSize = 0
					End If
					If rebuild_db = False And (refresh_rows = True Or seekSize = 0)
						Local date:String = reformatdate(searchDate,"DDMMYYYY","YYYYMMDD")
						message(1, "Delete entries for " + server + " entries dated " + searchdate)
						command = "DELETE FROM published WHERE ENDTIME LIKE '" + date + "%' and SERVER = '" + server + "';"
						If trace = True Then message(1,command)
						t_database.QueryDB(command,rc)
						If rc=False
							message(1,db_last_message)
							t_logfile.writelog(db_last_command)
							t_logfile.writelog(db_last_message)
							End
						End If
						command = "DELETE FROM file_writes WHERE ENDTIME LIKE '" + date + "%' and SERVER = '" + server + "';"
						If trace = True Then message(1,command)
						t_database.QueryDB(command,rc)
						If rc=False
							message(1,db_last_message)
							t_logfile.writelog(db_last_command)
							t_logfile.writelog(db_last_message)
							End
						End If
						SeekSize = 0
					End If
									
					' =================================================================================================================
					Local count:Int = 0
					Local doc_name:String = ""
					Local doc_id:String = ""
					Local load_date:String
					Local load_time:String
					Local load_time_AMPM:String
					Local vals:String[]
					Local datestamp:String
					Local timestamp:String
					If SeekSize>0
						If SeekStream(infile,SeekSize)<>-1
							bytesCount=SeekSize
						EndIf
					End If
					'rint "begin read"
					'*************************************************************************************************************
					'  Begin read loop until EOF detected
					'*************************************************************************************************************
					Repeat 
						Local linedata:String = ReadLine(infile)
						If Eof(infile) Then Exit
						bytesCount:+Len(linedata)
						'rint bytesCount+" "+bytesSize
						If Trim(linedata)= "" Then Continue
						If (MilliSecs() Mod 30) = 0
							UpdateProgBar ProgBar,(Float(bytesCount/1024)/Float(bytesSize/1024))
							SetGadgetText(progtext, "Read " + formatbytes(bytesCount, "KB") + " out of " + formatBytes(bytessize, "KB") + " transcount " + db_queue_items_count + "/" +db_max_queue_items + "  buffer " + Len(db_queue_items)+"/"+db_commit_buffer)
							RedrawGadget progbar
						EndIf
						Select PollEvent()
							Case EVENT_GADGETACTION						' interacted with gadget
							Case EVENT_WINDOWCLOSE						' close gadget
								If Confirm("Exit Program") 
									End
								End If
						End Select
						
						'rint LSet(linedata,80)
						If strFind(linedata, "Started processing") > 0
							''rint linedata		
							docname=StripDir(getstring(linedata,"Started processing "," with"))
							docname = Replace(docname,"'","@")
							starttime = checktime(Trim(Left(linedata,11)))
							endtime="HHMMSS"
							'Print docname
						End If
						If strFind(linedata,"Wrote out file") >0 
							'rint linedata		
							write_filename=StripDir(getstring(linedata,"Wrote out file: "," with"))
							write_filename = Replace(write_filename,"'","@")
							endtime = checktime(Trim(Left(linedata,11)))
							'If strFind(docname,"FROMPE")>0 And strFind(write_filename,Mid(docname,8,6))=0
							'	Print "====> "+logdate+"  "+docname+"  "+write_filename
							'End If
							'detail = "'"+server+"','"+reformatDate(logdate,"YYYY-MM-DD","YYYYMMDD")+"','"+endtime+"','"+docname+"','"+write_filename+"'"
							Local writeDatestamp:String = reformatDate(logdate, "YYYY-MM-DD", "YYYYMMDD")
							detail = "'" + server + "','" + writeDatestamp + endtime + "','" + docname + "','" + write_filename + "'"
							command = "INSERT INTO file_writes VALUES("+detail+");"
							'rint command
							rc = db_queueCommand(db_pntr,command)
							If rc<>0
								t_logfile.writelog(db_last_command)
								t_logfile.writelog(db_last_message)
								End
							End If
						End If
						If strFind(linedata,"Published on server")>0 Or strFind(linedata,"File in Error")>0  
							If strFind(linedata,"Published on server") >0
								status = 1
								Local entries:Int = Int((db_count(db_pntr,"published where doc_name = '"+docname+"' and status='0'")))
								'Print db_last_command
								If entries>0
									status = 2
								End If
							Else
								status = 0
								command = "DELETE FROM file_writes WHERE doc_name = '"+docname+"';"
								'rint command
								rc = db_queueCommand(db_pntr,command)
								If rc<>0
									t_logfile.writelog(db_last_command)
									t_logfile.writelog(db_last_message)
									End
								End If
							End If
							endtime = checktime(Trim(Left(linedata, 11)))
							Local writeDatestamp:String = reformatDate(logdate, "YYYY-MM-DD", "YYYYMMDD")
							detail = "'" + server + "','" + writeDatestamp + starttime + "','" + writeDatestamp + endtime + "','" + docname + "','" + status + "'"
							'rint detail
							command = "INSERT INTO PUBLISHED VALUES("+detail+");"
							'rint command
							rc = db_queueCommand(db_pntr,command)
							If rc<>0
								t_logfile.writelog(db_last_command)
								t_logfile.writelog(db_last_message)
								End
							End If
							records_added = True
							count:+1
							'If count=500 Then Exit
						End If
						'RedrawGadget teaTextArea1
						RedrawGadget progbar
						'rint "wend"
						'rint GCMemAlloced()+"  "+GCCollect()
						If bytesCount>=BytesSize Then Exit
					Forever
					CloseFile infile
					db_queueCommit(db_pntr)
					command = "UPDATE log_size SET size_bytes = '"+bytessize+"' WHERE DATE = '"+date+"' and SERVER = '"+server+"';"
					If trace=True Then message(1,command)
					t_database.QueryDB(command,rc)
					If rc=False
						message(1,db_last_message)
						t_logfile.writelog(db_last_command)
						t_logfile.writelog(db_last_message)
						End
					EndIf
					SetGadgetText(progtext, "")
					'Input
				EndIf
				' -----------------------------------------------------------------------------------
			End If
		Next
	Next
	't_database.QueryDB("VACUUM;",RC)
	'If rc <> True
	'	Print db_last_message
	'End If
'	CloseFile csvfile
'	message(1,"process published.csv")
'	Local infile:TStream = ReadFile("published.csv")
'	bytesSize = FileSize("published.csv")			
'	bytesCount = 0
'	While Not Eof(infile)
'		Local command:String = ReadLine(infile)
'		bytesCount:+Len(command)
'		db_queuecommand(db_pntr,command)
'		If (MilliSecs() Mod 30) = 0
'			UpdateProgBar ProgBar,(Float(bytesCount/1024)/Float(bytesSize/1024))
'			'SetGadgetText(label1,"Read "+formatbytes(bytesCount,"KB")+" out of "+formatBytes(bytesSize,"KB"))
'			RedrawGadget progbar
'		EndIf
'		Select PollEvent()
'			Case EVENT_WINDOWCLOSE						' close gadget
'				If Confirm("Exit Program") 
'					End
'				End If
'		End Select
'	Wend
'	CloseFile infile
'	RedrawGadget progbar
	db_queueCommit(db_pntr)
	'wait()
End Function

Function HouseKeeping()
	Local command:String, datecheck:String
	If reindex_flags<> ""
		Local days:Int = Int(getString(reindex_flags,"","|"))
		Local time:String = getString(reindex_flags,"|","")
		command = "SELECT VALUE FROM SYSTEM WHERE PARM='last_reindex';"
		datecheck = getcurrentdate("YYYYMMDD",-days)+time
		rows = t_database.QueryDB(command,rc)
		If rc=False
			t_logfile.writelog(command)
			t_logfile.writelog(db_last_message)
			message(1,command+"~n"+db_last_message)
		Else
			If Len(rows)=0
				t_logfile.writelog(command)
				t_logfile.writelog(db_last_message)
				message(1,command+"~nParm not found")
			Else
				If getString(rows[0],"","|") < datecheck
					t_logfile.writelog("Rebuilding Indexes")
					message(1,"Rebuilding Indexes")
					command = "REINDEX;"
					t_database.QueryDB(command,rc)
					If rc = False
						t_logfile.writelog(command)
						t_logfile.writelog(db_last_message)
						message(1,command+"~n"+db_last_message)
					Else
						message(1,"Index Rebuild Completed")
						command = "UPDATE SYSTEM SET VALUE = '"+getCurrentDate("YYYYMMDDhhmmss")+"' WHERE PARM = 'last_reindex';"
						t_database.QueryDB(command,rc)
					End If
				End If
			End If
		End If
	End If
End Function

Function Load_settings()
	Local cfgFile:TStream = ReadFile("..\conf\GatherLogStats.ini")
	If cfgfile = Null
		'Notify("unable to load ini file")
		End
	Else	
		While Not Eof(cfgfile)
			Local linedata:String = ReadLine(cfgfile)
			'''rint linedata
			If Left(linedata,1) = "*" Or Left(linedata,1) = "'" Then Continue
			If strFind(linedata,"SERVER=")=1
				Local count:Int = Len(ServerList)
				ServerList = ServerList[..count+1]
				ServerList[count] = getString(linedata,"SERVER=",";")
			End If
			If strFind(linedata,"REFRESH=")=1
				If GetString(linedata,"REFRESH=",";") = "YES"
					refresh_rows=True			
				End If
			End If
			If strFind(linedata,"REBUILD_DB=")=1
				If GetString(linedata,"REBUILD_DB=",";") = "YES"
					rebuild_db=True			
				End If
			End If
			If strFind(linedata,"WAIT=")=1
				If GetString(linedata,"WAIT=",";") = "YES"
					user_wait=True			
				End If
			End If
			If strFind(linedata,"HIDE=")=1
				If GetString(linedata,"HIDE=",";") = "YES"
					hide_window=True			
				End If
			End If
			If strFind(linedata,"DB_PATH=")=1
				db_path=GetString(linedata,"DB_PATH=",";")
				db_copy_path = ExtractDir(db_path)+"\Report2Web_stats.db"
			End If
			If strFind(linedata,"REPORTS_PATH=")=1
				reports_path=GetString(linedata,"REPORTS_PATH=",";")
			End If
			If strFind(linedata,"DATA_PATH=")=1
				data_path=GetString(linedata,"DATA_PATH=",";")
			End If
			If strFind(linedata,"DAYS2LOAD=")=1
				days2load = Int(getstring(linedata,"DAYS2LOAD=",";"))
			End If
			If strFind(linedata,"DAYS2KEEP=")=1
				days2keep = Int(getstring(linedata,"DAYS2KEEP=",";"))
				If days2keep<0 Then days2keep=0
			End If
			If strFind(linedata,"REINDEX=")=1
				reindex_flags = getstring(linedata,"REINDEX=",";")
			End If
			If strFind(linedata,"TRACE=")=1
				trace = Int(getstring(linedata,"TRACE=",";"))
			End If
			If strFind(linedata, "recsPerTransaction=")
				Local temp:Int = Int(getString(linedata, "recsPerTransaction=", ";"))
				If temp > 0
					recsPerTransaction = temp
					db_commit_buffer = recsPerTransaction * 512
					db_max_queue_items = recsPerTransaction
				End If
			End If
			If days2load > days2keep
				days2load = days2keep
			End If
		Wend
		CloseFile cfgfile
		If hide_window
			user_wait = False
		End If
	End If
End Function

Function wait()
	Repeat
		Select WaitEvent()
			Case EVENT_GADGETACTION						' interacted with gadget
				DoGadgetAction()
			Case EVENT_WINDOWCLOSE						' close gadget
				If Confirm("Exit Program") 
					Exit
				End If
		End Select
	Forever
End Function

Function DoGadgetAction()
	Select EventSource()

		Case btnCANCEL	' user pressed button
			If Confirm("Exit Program")
				End
			End If
	End Select
End Function

Function Gen_Reports()
	Local rows:String[] 
	Local outfile:TStream
	Local vals:String[]
	Local command:String
	Local searchdate:String
	Local fromdate:String
	Local todate:String
	If FileType(reports_path)<>2
		CreateDir(reports_path)
	End If
	' ************************************************************************************************************
	' Write out to trace file
	' ************************************************************************************************************
	'message(1,"Generating trace file")
	'rows = t_database.QueryDB("select date,substr(time,1,2),count(*) from published group by date||substr(time,1,2) order by date||substr(time,1,2) desc;",rc) 
	'If rc = True And Len(rows)>0
	'	outfile = WriteFile(reports_path+"\TRACE_FILE.txt")		
	'	For Local text:String = EachIn rows
	'		''rint text		
	'		WriteLine(outfile,text)
	'	Next
	'	CloseFile outfile
	'End If
	' ************************************************************************************************************
	' Generate Publication report
	' ************************************************************************************************************
	Local count:Int = 0
	Local col1:String
	Local col2:String
	Local col3:String
	Local ralign:String = " align="+Chr(34)+"right"+Chr(34)
	Local lalign:String = " align="+Chr(34)+"left"+Chr(34)
	outfile = WriteFile(reports_path+"\Publication_Report.html")		
	message(1,"Reporting files published")
	fromdate = GetCurrentDate("YYYYMMDD",-14)
	todate = GetCurrentDate("YYYYMMDD")
	WriteLine(outfile,"<!DOCTYPE html>")
	WriteLine(outfile,"<html>")
	WriteLine(outfile,"<head>")
	WriteLine(outfile,"<title>R2W Counts</title>")
	WriteLine outfile,"<style>"
	WriteLine outfile,"* {margin:20;}"
	WriteLine outfile,"table {"	
	WriteLine outfile,"  font-family: arial, sans-serif;"
	WriteLine outfile,"  font-size: 12px;"
	WriteLine outfile,"  border-collapse: collapse;"
	WriteLine outfile,"  width: 400px;"
	WriteLine outfile,"  margin: 20px;"
	WriteLine outfile,"}"
	WriteLine outfile,"td, th {"
	WriteLine outfile,"  border: 1px solid #bbbbbb;"
	'WriteLine outfile,"  text-align: left;"
	WriteLine outfile,"  padding: 8px;"
	WriteLine outfile,"}"
	WriteLine outfile,"tr:nth-child(odd) {"
	WriteLine outfile,"  background-color: #eeeeee;"
	WriteLine outfile,"}"
	WriteLine outfile,"p {"
	WriteLine outfile,"  padding-left: 25px;"
	WriteLine outfile,"}"
	WriteLine outfile,"</style>"
	WriteLine(outfile,"</head>")
	WriteLine(outfile,"<body>")
	WriteLine(outfile,"<h2>Report2Web Publication Counts</h2>")
	WriteLine(outfile,"<h3>Published counts by day (last 14 days)</h3>")
	WriteLine(outfile, "<table style=" + Chr(34) + "width:25%" + Chr(34) + ">")
	WriteLine(outfile,"<tr><th>Date</th><th>Processed</th><th>Errors</th><th>Created</th></tr>")
	'rows = t_database.QueryDB("select date,count(*) from published where date >= '"+fromdate+"' and date <= '"+todate+"'  group by date order by date desc;",rc) 
	command = "select substr(endtime,1,8),count(*) from published where status = '1' and endtime >= '" + fromdate + "' group by substr(endtime,1,8) order by endtime desc;"
	'Print command
	rows = t_database.QueryDB(command,rc) 

	If rc = False
		t_logfile.writelog(db_last_command)
		t_logfile.writelog(db_last_message)
		Return
	Else
		If Len(rows) > 0
			For Local text:String = EachIn rows
				WriteLine(outfile,"<tr>")
				vals = parseVar(text,"|")
				Local errorCount:Int = Int(db_count(db_pntr, "published where status = '0' and endtime like '" + vals[0] + "%'"))
				'Print db_last_command
				'Local fileCount:Int = Int(db_count(db_pntr,"published where status = '0' and date = '"+vals[0]+"'"))
				Local fileWrites:Int = Int(db_count(db_pntr, "file_writes where endtime like '" + vals[0] + "%'"))
				If count = 0
					'text = LSet("Today",12)+RSet(vals[1],10)	
					col1 = "Today"
				Else
					'text = LSet(reformatdate(vals[0],"YYYYMMDD","DD-MM-YYYY"),12)+RSet(vals[1],10)
					col1 = reformatdate(vals[0],"YYYYMMDD","DD-MM-YYYY")
				EndIf
				text = "<td>"+col1+"</td><td"+ralign+">"+trunc(vals[1],0,True)+"</td><td"+ralign+">"+trunc(errorCount,0,True)+"</td><td"+ralign+">"+trunc(fileWrites,0,True)+"</td>"
				WriteLine outfile,text
				WriteLine(outfile,"</tr>")
				count:+1
				'Print reformatdate(vals[0], "YYYYMMDD", "DD-MM-YYYY")
			Next
		End If
	End If
	WriteLine(outfile, "<tr>")
	Local pub_total:Int = Int((db_count(db_pntr,"published where status='1' and date >= '"+fromdate+"'")))
	Local error_total:Int = Int((db_count(db_pntr,"published where status='0' and date >= '"+fromdate+"'")))
	Local file_write_total:Int = Int((db_count(db_pntr,"file_writes where date >= '"+fromdate+"'")))
	WriteLine(outfile,"<th>Total files:</th><th"+ralign+">"+trunc(pub_total,0,True)+"</th><th"+ralign+">"+trunc(error_total,0,True)+"</th><th"+ralign+">"+trunc(file_write_total,0,True)+"</th>")
	WriteLine(outfile,"</tr>")
	WriteLine(outfile,"</table>")
	WriteLine(outfile," ")

'	Return
	message(1,"Reporting counts by hour")
	WriteLine(outfile, "<h3>Published counts by the hour (Last 48 hours)</h3>")
	WriteLine(outfile, "<table style=" + Chr(34) + "width:25%" + Chr(34) + ">")
	WriteLine(outfile,"<tr><th>Date</th><th>Time Range</th><th>Processed</th><th>Errors</th><th>Created</th></tr>")
	searchdate = getcurrentdate("YYYYMMDDhh",-2)
	rows = t_database.QueryDB("select max(starttime) from published where starttime like '" + todate + "%';", rc)
	If rc = False
		t_logfile.writelog(db_last_command)
		t_logfile.writelog(db_last_message)
		Return
	End If
	Local lasttime:String = reformatDate(rows[0], "YYYYDDMMhhmmss", "hh:mm")
'	rows = t_database.QueryDB("select date,substr(starttime,1,2),count(*) from published where date||substr(starttime,1,2) > '"+searchdate+"' group by date||substr(starttime,1,2) order by date||substr(starttime,1,2) asc;",rc) 
	command = "select substr(starttime,1,8),substr(starttime,9,2),count(*) from published where status = '1' and starttime > '" + searchdate + "' group by substr(endtime,1,10) order by substr(endtime,1,10) desc;"
	'Print command
	rows = t_database.QueryDB(command,rc) 
	count=0


	If rc = False
		t_logfile.writelog(db_last_command)
		t_logfile.writelog(db_last_message)
		Return
	Else
		If Len(rows) > 0
			For Local text:String = EachIn rows
				WriteLine(outfile,"<tr>")
				vals = parseVar(text,"|")
	'			Local errorCount:Int = Int(db_count(db_pntr,"published where status = '0' and date||substr(endtime,1,2) = '"+vals[0]+Left(vals[1],2)+"'"))
	'			Local fileWriteCount:Int = Int(db_count(db_pntr,"file_writes where date||substr(endtime,1,2) = '"+vals[0]+Left(vals[1],2)+"'"))
				Local errorCount:Int = Int(db_count(db_pntr, "published where status = '0' and endtime like '" + vals[0] + Left(vals[1], 2) + "%'"))
				Local fileWriteCount:Int = Int(db_count(db_pntr, "file_writes where endtime like '" + vals[0] + Left(vals[1], 2) + "%'"))
				If count = 0
					'text = reformatDate(vals[0],"YYYYMMDD","DD-MM-YYYY")+" "+LSet(vals[1]+":00-"+lasttime,12)+RSet(vals[2],10)
					col1=reformatDate(vals[0],"YYYYMMDD","DD-MM-YYYY")
					col2=LSet(vals[1]+":00-"+lasttime,12)
				Else
					'text = reformatDate(vals[0],"YYYYMMDD","DD-MM-YYYY")+" "+LSet(vals[1]+":00-"+vals[1]+":59",12)+RSet(vals[2],10)
					col1=reformatDate(vals[0],"YYYYMMDD","DD-MM-YYYY")
					col2=LSet(vals[1]+":00-"+vals[1]+":59",12)
				End If
				text = "<td>"+col1+"</td><td>"+col2+"</td><td"+ralign+">"+trunc(vals[2],0,True)+"</td><td"+ralign+">"+trunc(errorCount,0,True)+"</td><td"+ralign+">"+trunc(fileWriteCount,0,True)+"</td>"
				WriteLine(outfile,text)
				WriteLine(outfile,"</tr>")
				count:+1
			Next
		End If
	End If
	WriteLine(outfile, "</table>")
	WriteLine(outfile,"<p><h5>"+progname+"   ver:"+version+"</h5></p>")
	WriteLine(outfile,"</body>")
	WriteLine(outfile,"</html>")
	CloseFile outfile
End Function

Function Gen_data_files()
	' ************************************************************************************************************
	' Generate delimited files to be used by other programs
	' ************************************************************************************************************
	Local work_path:String = "..\work"
	If FileType(work_path) <> 2
		CreateDir(work_path)
	End If
	Local rows:String[] 
	Local outfile:TStream
	Local vals:String[]
	Local command:String
	Local searchdate:String
	Local fromdate:String
	Local todate:String
	'db_column_seperator = ","
	searchdate = getcurrentdate("YYYYMMDDhhmm")
	'searchdate = "201701161430"
	Local hour:Int = Int(Mid(searchdate,9,2))
	Local minute:Int = Int(Mid(searchdate,11,2))
	Local found:Int
	Local stephour:Int
	Local stepminute:Int
	Local outtext:String
	Local checktime:String
	Local checkdate:String
	Local outpath:String
	Local range:Int = 2
	
	'********************************************************************************************************
	'*        Generate data for Publication Graph Combined
	'********************************************************************************************************
	' initialize counts
	'Local pubCount_total:Int = 0
	'Local errorCount_total:Int = 0
	'Local fileCount_total:Int = 0
	Local rowTotal:Int
	Local rowCount:Int
	range = 24
	message(1,"Generating files for graphs")

	'********************************************************************************************************
	'*        Pubication Counts per minute since Midnight
	'********************************************************************************************************
	message(1, "  Publication all servers")
	outpath = work_path + "\prodcounts_Running_ALL.csv"
	DebugLog("outpath=" + outpath)
	t_output.ClearOutput() 'Clear output buffer
	stephour = hour-range
	If stephour < 0 Then stephour = 0 ' Set hour to midnight if value if below zero
	stepminute = minute
	rowTotal = 0
	'command = "select endtime,sum(count) as count from published_counts_by_server where endtime >= '" + Left(searchdate, 8) + "' group by endtime;"
	command = "select substr(endtime,1,12),count(*) from published where status = '1' and endtime >= '" + Left(searchdate, 8) + "' group by substr(endtime,1,12) order by endtime asc";
	'Print command
	rows = t_database.QueryDB(command, rc)
	If rc = False
		t_logfile.writeLog(db_last_command)
		t_logfile.writeLog(db_last_message)
	Else
		Repeat
			checktime = Right("00"+stephour,2)+Right("00"+stepminute,2)
			checkdate = Left(searchdate,8)
			
			rowCount = 0
			If Len(rows) > 0
				For Local rowdata:String = EachIn rows
					vals = rowdata.Split("|")
					If vals[0] = checkdate + checktime
						rowCount = Int(vals[1])
						Exit
					End If
				Next
			End If
			rowTotal:+rowCount
			outtext = Right("00" + stephour, 2) + ":" + Right("00" + stepminute, 2) + "," + rowTotal
			t_output.Add2Output(outtext)
			stepminute:+1
			If stepminute >= 60
				stepminute = 60-stepminute
				stephour:+1
			End If
		Until (stephour = hour) And (stepminute = minute)
	End If
	t_output.WriteOutput(outpath)


	'********************************************************************************************************
	'*        Pubication Error Counts per minute since Midnight
	'********************************************************************************************************
	message(1, "  Publication errors all servers")
	outpath = work_path + "\errorcounts_Running_ALL.csv"
	DebugLog("outpath=" + outpath)
	t_output.ClearOutput() 'Clear output buffer
	stephour = hour-range
	If stephour < 0 Then stephour = 0 ' Set hour to midnight if value if below zero
	stepminute = minute
	rowTotal = 0
	'command = "select endtime,sum(count) as count from published_counts_by_server where endtime >= '" + Left(searchdate, 8) + "' group by endtime;"
	command = "select substr(endtime,1,12),count(*) from published where status = '0' and endtime >= '" + Left(searchdate, 8) + "' group by substr(endtime,1,12) order by endtime asc";
	'Print command
	rows = t_database.QueryDB(command, rc)
	If rc = False
		t_logfile.writeLog(db_last_command)
		t_logfile.writeLog(db_last_message)
	Else
		Repeat
			checktime = Right("00"+stephour,2)+Right("00"+stepminute,2)
			checkdate = Left(searchdate,8)
			
			rowCount = 0
			If Len(rows) > 0
				For Local rowdata:String = EachIn rows
					vals = rowdata.Split("|")
					If vals[0] = checkdate + checktime
						rowCount = Int(vals[1])
						Exit
					End If
				Next
			End If
			rowTotal:+rowCount
			outtext = Right("00" + stephour, 2) + ":" + Right("00" + stepminute, 2) + "," + rowTotal
			t_output.Add2Output(outtext)
			stepminute:+1
			If stepminute >= 60
				stepminute = 60-stepminute
				stephour:+1
			End If
		Until (stephour = hour) And (stepminute = minute)
	End If
	t_output.WriteOutput(outpath)

	'********************************************************************************************************
	'*        Files read Count per minute since Midnight
	'********************************************************************************************************
	message(1, "  Publication errors all servers")
	outpath = work_path + "\filecounts_Running_ALL.csv"
	DebugLog("outpath=" + outpath)
	t_output.ClearOutput() 'Clear output buffer
	stephour = hour-range
	If stephour < 0 Then stephour = 0 ' Set hour to midnight if value if below zero
	stepminute = minute
	rowTotal = 0
	'command = "select endtime,sum(count) as count from published_counts_by_server where endtime >= '" + Left(searchdate, 8) + "' group by endtime;"
	command = "select substr(endtime,1,12),count(*) from file_writes where endtime >= '" + Left(searchdate, 8) + "' group by substr(endtime,1,12) order by endtime asc";
	'Print command
	rows = t_database.QueryDB(command, rc)
	If rc = False
		t_logfile.writeLog(db_last_command)
		t_logfile.writeLog(db_last_message)
	Else
		Repeat
			checktime = Right("00"+stephour,2)+Right("00"+stepminute,2)
			checkdate = Left(searchdate,8)
			
			rowCount = 0
			If Len(rows) > 0
			
				For Local rowdata:String = EachIn rows
					vals = rowdata.Split("|")
					If vals[0] = checkdate + checktime
						rowCount = Int(vals[1])
						Exit
					End If
				Next
			End If
			rowTotal:+rowCount
			outtext = Right("00" + stephour, 2) + ":" + Right("00" + stepminute, 2) + "," + rowTotal
			t_output.Add2Output(outtext)
			stepminute:+1
			If stepminute >= 60
				stepminute = 60-stepminute
				stephour:+1
			End If
		Until (stephour = hour) And (stepminute = minute)
	End If
	t_output.WriteOutput(outpath)
				
	'Return

	'********************************************************************************************************
	'*        Generate data for Publication Graphs each server
	'********************************************************************************************************
	range = 2	' set to last 2 hours
	t_output.ClearOutput()
	message(1,"  Publication by server")
	command = "select * from serverlist_view;"
	Local servers:String[] = t_database.QueryDB(command,rc)
	If rc=False
		message(1,db_last_message)
		Return
	End If
	
	For Local server:String = EachIn servers
		server = getString(server, "", "|")
		outpath = work_path + "\prodcounts_" + server + ".csv"
		DebugLog("outpath=" + outpath)
		t_output.ClearOutput() 'Clear output buffer
		stephour = hour-range
		If stephour < 0 Then stephour = 0 ' Set hour to midnight if value if below zero
		stepminute = minute
		rowTotal = 0
		command = "select substr(endtime,1,12),count(*) from published where status = '1' and server = '" + server + "' and endtime >= '" + Left(searchdate, 8) + "' group by substr(endtime,1,12) order by endtime asc";
		DebugLog("command=" + command)
		rows = t_database.QueryDB(command, rc)
		If rc = False
			t_logfile.writeLog(db_last_command)
			t_logfile.writeLog(db_last_message)
		Else
			Repeat
				checktime = Right("00"+stephour,2)+Right("00"+stepminute,2)
				checkdate = Left(searchdate,8)
				
				rowCount = 0
				If Len(rows) > 0
					For Local rowdata:String = EachIn rows
						vals = rowdata.Split("|")
						If vals[0] = checkdate + checktime
							rowCount = Int(vals[1])
							Exit
						End If
					Next
				End If
				'rowTotal:+rowCount
				outtext = Right("00" + stephour, 2) + ":" + Right("00" + stepminute, 2) + "," + rowCount
				t_output.Add2Output(outtext)
				stepminute:+1
				If stepminute >= 60
					stepminute = 60-stepminute
					stephour:+1
				End If
			Until (stephour = hour) And (stepminute = minute)
		End If
		t_output.WriteOutput(outpath)
	Next
	'********************************************************************************************************
	'*        Pubication Performance
	'********************************************************************************************************
	message(1, "  Publication performance")
	range = 2
	outpath = work_path + "\pub_perf_all.csv"
	DebugLog("outpath=" + outpath)
	t_output.ClearOutput() 'Clear output buffer
	stephour = hour-range
	If stephour < 0 Then stephour = 0 ' Set hour to midnight if value if below zero
	stepminute = Int(minute / 5) * 5
	rowTotal = 0
	'command = "select endtime,sum(count) as count from published_counts_by_server where endtime >= '" + Left(searchdate, 8) + "' group by endtime;"
	'command = "select substr(endtime,9,4),duration from duration_view where endtime >= '" + Left(searchdate, 8) + RSet("00" + stephour, 2) + RSet("00" + stepminute, 2) + "' and duration>=0 order by endtime;";
	command = "select substr(endtime,9,2)||':'||substr(endtime,11,2) as time,count(*) as entries, min(duration) as low,max(duration) as high, sum(duration)/count(*) as average" + ..
			" from duration_view" + ..
			" where starttime >= '" + Left(searchdate, 8) + RSet("00" + stephour, 2) + RSet("00" + stepminute, 2) + "' and duration > 0" + ..
			" group by substr(endtime,9,4)" + ..
			" order by endtime;"
	Print command
	rows = t_database.QueryDB(command, rc)
'	If rc = False
'		t_logfile.writeLog(db_last_command)
'		t_logfile.writeLog(db_last_message)
'	Else
'		For Local rowdata:String = EachIn rows
'			vals = rowdata.Split("|")
'			'Print rowdata
'			outtext = Left(vals[0], 2) + ":" + Right(vals[0], 2) + "," + Int(vals[4])
'			t_output.Add2Output(outtext)
'		Next
'	End If
'	t_output.WriteOutput(outpath)

	If rc = False
		t_logfile.writeLog(db_last_command)
		t_logfile.writeLog(db_last_message)
	Else
		Repeat
			checktime = Right("00"+stephour,2)+":"+Right("00"+stepminute,2)
			checkdate = Left(searchdate, 8)
			rowCount = 0
			If Len(rows) > 0
				For Local rowdata:String = EachIn rows
					vals = rowdata.Split("|")
					If vals[0] = checktime
						rowCount = Int(vals[4])
						Exit
					End If
				Next
			End If
			'rowTotal:+rowCount
			outtext = Right("00" + stephour, 2) + ":" + Right("00" + stepminute, 2) + "," + rowCount
			t_output.Add2Output(outtext)
			stepminute:+1
			If stepminute >= 60
				stepminute = 60 - stepminute
				stephour:+1
			End If
		Until (stephour = hour) And (stepminute = minute)
	End If
	t_output.WriteOutput(outpath)


		


'	For Local server:String = EachIn servers
'		server = getString(server,"","|")
'		outpath = work_path+"\prodcounts_"+server+".csv"
'		DebugLog("outpath=" + outpath)
'		stephour = hour - range
'		stepminute = minute
'		If stephour<0 Then stephour = 0
'		Repeat
'			checktime = Right("00"+stephour,2)+Right("00"+stepminute,2)
'			checkdate = Left(searchdate,8)
'			Local pubCount:Int = Int(db_count(db_pntr, "published where status = '1' and server = '" + server + "' and endtime like '" + checkdate + checktime + "%'"))
'			'Print db_last_command
'			outtext = Right("00"+stephour,2)+":"+Right("00"+stepminute,2)+","+pubcount '+","+errorcount+","+fileCount	
'			t_output.Add2Output(outtext)
'			stepminute:+1
'			If stepminute>=60
'				stepminute = 60-stepminute
'				stephour:+1
'			End If
'		Until (stephour = hour) And (stepminute = minute)
'		t_output.WriteOutput(outpath)
'	Next
	
	'********************************************************************************************************
	'*       Move files to data_path
	'********************************************************************************************************
	Local files:String[] = LoadDir(work_path)
	For Local filename:String = EachIn files
		Local source:String = work_path+"\"+filename
		Local dest:String = data_path+"\"+filename
		'rint source+"   "+dest
		If Delete_File(dest)
			rc=RenameFile(source,dest)
		EndIf
	Next
End Function

Function delete_file:Int(path:String)
	Local result:Int = False
	Local attempts:Int = 0
	Repeat
		If DeleteFile(path)
			result=True
			Exit		
		End If
		Delay 500
		attempts:+1
	Until attempts = 5
	Return result
End Function
Function write2file(file:TStream, text:String)
	'rint text
	WriteLine file,text
End Function
Function db_getCount:Int(command:String)
	'DebugLog(command)
	Local rc:Int
	Local rows:String[] = t_database.QueryDB(command, rc)
	If rc = False
		t_logfile.writelog(db_last_message)
		t_logfile.writelog(db_last_command)
		Return 0
	Else
		If Len(rows) = 0
			Return 0
		Else
			Return Int(getString(rows[0], "", "|"))
		EndIf
	End If
	
End Function

Function message(num:Int, text:String, clear:Int = False)
	If num=0
		If clear = True
			SetGadgetText teaTextArea0,text+"~n"
		Else
			SetGadgetText teaTextArea0,text+"~n"+GadgetText(teaTextArea0)
		EndIf
		RedrawGadget teaTextArea0
	Else
		If clear = True
			SetGadgetText teaTextArea1,text+"~n"
		Else
			SetGadgetText teaTextArea1,text+"~n"+GadgetText(teaTextArea1)
		EndIf
		RedrawGadget teaTextArea1
	End If
	'''rint text
	t_logfile.writelog(text)
End Function



Function checktime:String(time:String)
	Local vals:String[] = parseVar(Left(time,8),":")
	Local hours:Int = Int(vals[0])
	Local minutes:Int = Int(vals[1])
	Local seconds:Int = Int(vals[2])
		If Right(time,2) = "AM"
		If hours = 12 Then hours = 0
	Else
		If hours<>12 Then hours:+12
	End If
	Local result:String = Right("00"+hours,2)+Right("00"+minutes,2)+Right("00"+seconds,2)
	Return result
End Function

Type t_database
	Function BuildDB()
		message(1,"Build Database")
		'rint db_path
		DebugLog(db_openDB(db_path,db_pntr,True))
		Local columns:String
		Local command:String
		'********************************************************************************
		' Tables
		db_createTable(db_pntr,"system","PARM TEXT, VALUE TEXT")
		t_database.QueryDB("INSERT INTO SYSTEM VALUES('last_cleanup','');",rc)
		t_database.QueryDB("INSERT INTO SYSTEM VALUES('last_reindex','');",rc)
		DB_CreateTable(db_pntr, "published", "server TEXT, starttime TEXT, endtime TEXT, doc_name TEXT, status TEXT")
		DB_CreateTable(db_pntr, "file_writes", "server TEXT, endtime TEXT, doc_name TEXT, file_name TEXT")
		db_createTable(db_pntr, "log_size", "server TEXT,date TEXT,size_bytes TEXT")
		'********************************************************************************
		' Views
		db_command(db_pntr, "CREATE VIEW serverlist_view AS select distinct(server) from published order by server;")
		command = " drop view if exists duration_view;" + ..
				" create view duration_view as "+..
				" select "+.. 
				" 	server as server, "+..
				" 	endtime as timestamp, "+..
				" 	starttime, "+..
				" 	endtime, "+..
				" 	((substr(endtime,9,2)*360)+(substr(endtime,11,2)*60)+substr(endtime,13,2)) - ((substr(starttime,9,2)*360)+(substr(starttime,11,2)*60)+substr(starttime,13,2)) as duration , "+..
				" 	doc_name "+.. 
				" from "+.. 
				"  published " + ..
				" where "+.. 
				" 	(starttime <> '' and endtime <> '') "+.. 
				" order by "+..
				"  starttime, " + ..
				"  server; "
		db_command(db_pntr, command)
		'db_command(db_pntr, "CREATE VIEW published_counts_by_server as select server,substr(endtime,1,12) as endtime,count(*) as count from published where status = '1' group by server,substr(endtime,1,12) order by endtime asc;")
		'********************************************************************************
		' Indexes
		db_command(db_pntr, "CREATE INDEX published_idx1 ON published (server);")
		db_command(db_pntr, "CREATE INDEX published_endtime_idx1 ON published (server,endtime);")
		'db_command(db_pntr,"CREATE INDEX file_writes_idx1 ON file_writes (date, endtime);")
		'db_command(db_pntr,"CREATE INDEX file_writes_idx2 ON file_writes (server, date, endtime);")
		'db_command(db_pntr,"CREATE INDEX pub_date_idx ON published (date);")
		'db_command(db_pntr, "CREATE INDEX pub_doc_name_idx ON published (doc_name,date);")
		'db_command(db_pntr,"CREATE INDEX pub_server_idx ON published (server);")
		'db_command(db_pntr,"CREATE INDEX published_idx1 ON published (status, date, endtime);")
		'db_command(db_pntr,"CREATE INDEX published_idx2 ON published (status, server, date, endtime);")
		db_closeDb(db_pntr)
	End Function
	Function CleanupDB(days:Int)
		message(1,"Cleanup database  delete entries older than "+days+" days.")
		Local date:String = getCurrentDate("YYYYMMDD",-days)
		If Int((db_count(db_pntr, "PUBLISHED WHERE ENDTIME < '" + date + "'"))) > 0
			command = "DELETE FROM PUBLISHED WHERE ENDTIME < '" + date + "';"
			If trace = True Then message(1,command)
			t_database.QueryDB(command,RC)
	
			command = "DELETE FROM FILE_WRITES WHERE ENDTIME < '" + date + "';"
			If trace = True Then message(1,command)
			t_database.QueryDB(command,RC)
	
			command = "DELETE FROM LOG_SIZE WHERE DATE < '"+date+"';"
			If trace = True Then message(1,command)
			t_database.QueryDB(command,RC)
	
			t_database.QueryDB("VACUUM;",RC)
		EndIf 
	End Function
	Function QueryDB:String[] (command:String, rc:Int Var)
		If db_pntr = Null
			db_last_message = "Database not open"
			rc = False
			'Return Null
		Else
			If db_command(db_pntr,command)=SQLITE_OK
				rc = True
				'Return db_callbacklist
			Else
				rc = False
				'Return Null
			End If
		EndIf
		'SetPointer(POINTER_DEFAULT)
		If rc = True
			Return db_callbacklist
		Else	
			Return Null
		End If
	End Function		
End Type
Type t_output
	Global outputList:TList = CreateList()
	Function ClearOutput()
		ClearList(outputList)
	End Function
	Function Add2Output(text:String)
		ListAddLast OutputList,text
	End Function
	Function WriteOutput(path:String)
		message(1,path)
		Local outfile:TStream = WriteFile(path) 
		If outfile = Null
			'rint "Error opening "+path+" for write"
		Else
			If Not ListIsEmpty(OutputList)
				For Local linedata:String = EachIn OutputList
					WriteLine outfile,linedata
				Next
			End If
			CloseFile outfile
			ClearOutput()
		EndIf
	End Function
End Type


