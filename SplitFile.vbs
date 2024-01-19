Option Explicit
	Dim Conn ,rs, fs , Conn1, rs1 , Fieldseperator ,origonalxlsfilename, doublerows , RemoveXcharsfromDate ,FindStartingCharsandRemovePartID,FindStartingCharandRemoveAllAfter,objDictDoubleRowsRemoval ,daysSincelastMovementObsolete, discount, objDict, dbf , MaxScanRows, xlsSheetName ,FileName  ,multistock ,splitdateby , DiscountSubbrandMore , multisheets,FindsubsequentcharsBeginingatPostionX, LookingFor , StartPostionAT ,   FindFirstInstanceofAndRemoveit, filelist ,ConvertSheetstocsv,daysSincelastMovementScrap ,strLastMovement, lngLastMovement, today, i , j , e , t ,numberInStock, dealer , brand,  getbrands ,mincharactorallowedpartid , qty , maxcharactorallowedpartid , replacestartcharcount , replacestartchar , FindStartingCharandRemoveSpecifiedChars, disposition , Filelocation, filesavelocation, fileextention, files, file ,isObsolete, isScrap , isReject, obsFilePath, curFilePath, scrapFilePath , dateFormat , checkpartidlenght , getsqlParttIDmaxminmaxlengh , getnumberinstock, getsqlPartFilterBy ,regEx ,SettingsForEmptyDate ,SettingsForInvalidDate ,fileToBeSplit  , extinterfacesql, SetextInterfaceused, renamestockfiletoold   , matchandreplace  , MinDaysInStock  , PickDateToUse  , ObsCategories, ScrapCategories , ExcludeCategories , UseCategory , CatLengthcombo  ,counterCurrent , counterScrap , counterObs , xls  ,RemoveLeadingZerosWhenPartIDisAllNumbers , FindStartingCharandRemoveAllBefore,DateisText, xpartsID ,xnumberInStock, minprice, matcharray, xBrand,xOutlet ,xLastPurchase ,xLastSale ,xPartsCategory ,xFilterby ,xAndFilterBy ,xAndNotFilterBy ,xsubbrands ,xmanufacture ,xcondition  ,xretailprice ,xpurchaseprice, xdescription  ,objFSO , DealerDic  ,DaysSincelastMovementObsolete_origonal , DaysSincelastMovementScrap_origonal , usedatelastPurchase_Origonal, ObsCategories_Origonal , ScrapCategories_Origonal , ExcludeCategories_Origonal , UseDate__Origonal , minprice_Origonal ,RemoveXfromstart ,multifilekillrootfolder
	Dim removeColumnXfromTextFile, usedatelastPurchase,BrandWildcard,NumberinStockisInteger,RemoveEmptyScrapObsCurrent,WriteDisposition,ConvertXLSToTXTTAB,ConvertXMLtoCSV,ImportDealerInfo ,UsePartCategoryAsDefault ,UseDate ,showerror ,AllowStockFikeToBeAutoRenamed
	Const current = "current" ,scrap = "scrap" , obs = "obs" 			
	SetLocale(1045)  '' sets to dutch language settings '' 1033 for german descptions
	
'########################  v 117 #############################################

'CONFIGURATION SECTION
'v116 remove xchars from start of date column
'v115 is a todo to fix, nothing has changed colin still needs to work on this. 
'v114 Divide By QTY bug, too easy for raw qty to be larger than maxqyt if / 100, so a lot of 9999 qty can happen
'v113 Only ConverttoCSV for bad xls files
'v112 can use a Tab (vbTab) as a Fieldseperator 
'v111 remove chars after selected char FindStartingCharandRemoveAllAfter
'v110 Fix the worksheetno and worksheet name, merge into one function. Clean up the text inside of xls do not use regex. 
'v109 fix empty carts caterogy
'V108 FIX BUG OF 106, EMPTY VALUE REMOVES EVERYTHING
'v107 Remove leadingZeros if partid contains Leters
'v106 Remove Column from Text.cvs files. tested on 1017287
'105 Double Rows , handles in the RegeXRun function automatically
'v104 UNL files , handle automatically in the RegeXRun function for everyone
'#####################################################################

FileName = "brandA.txt,brandB.txt"			            'File name, multiple files seperate filenames with comma , USE false for autoreadfiles for when we dont know the filenameS.
AllowStockFikeToBeAutoRenamed= "false" 	    'This will check if there is any file with the same extension as the stock file and rename it to the name specified above for the stock file

Const FormatDelimited = ","					'F = Fixedlength , T = tabdelimited (must use T for xls/xml)  ; = semi colon Delimiter
Const ReadCharacterSet = 2		        '1=OEM  2=ANSI 3=Unicode  4=65001		
Const DecimalSymbol = "."					
Const DidgetGroupingSymbol =""				'1.000,00  is 1000 units. We need to always remove the "."

''Schema Columns  
Const S_partsID = "2"						'Fixed width file use StartPos and EndPosition seperated with a comma. EG "3,9" .XLS count starts with 0 | Converting xls > csv count starts at 1  | 				
Const S_numberInStock = "5"  			    'For MultiQTY( where one file has mulit dealers split by numberin stock) use a ",".  This will turn numberinstock into an array of dealers.  Will not work on width delimiter!
Const S_Brand = "4"					        
Const S_Outlet= "1"                          
Const S_LastPurchase = "9"                  
Const S_LastSale = "6"                       
Const S_PartsCategory =""                   
Const S_Filterby =""						'Pick Column here, and below in the  "FilterBy" you MUST insert what you want to filter by.   Will find matching values.   Use , seperateor for 
Const S_AndFilterBy=""						'Pick Column here, and below in the  "AndFilterBy" you MUST  insert what you want to filter by.   Will find matching values.   
Const S_AndNotFilterBy=""                   'Pick Column here, and below in the  "AndNotFilterBy"  you MUST insert what you want to filter by.  Will EXCLUDE matching values.   
 
' Column number PartFarm & ThePart
Const S_subbrands=""     					'Column number, XLS starts count with 0 txt, csv counts from 1.	Fixed width file use StartPos and EndPosition seperated with a comma.                                                           
Const S_retailprice=""                                     
Const S_description=""                     
Const S_manufacture=""
Const S_condition=""  
Const S_purchaseprice="" 

'Brands/Oulets data.
Const outlets="automatic"        	       	' "automatic" for all distinct outlets. Or specify seperate outlets seperated by commas. To speed up large file, best to insert values.
Const brands="automatic"					' "automatic" for all distinct brands. Or specify seperate brands seperated by commas
Const BrandsNotinColumns=""			    	' S_Brand must be "" ,  will find the BRAND somwwhere in your txt file. Will find brand match, write files for that brand, untill next brand is found.
const BrandsInPartID=""				    	' [leftover] FOR everything else. Enter brands with comma to seperate.  This will search for Brands inside the partid, only at the starting position
BrandWildcard = "false"				' makes the brands a wildcard value, so N will search for all values in brand table with nissan, nis, nestel, nss ect... default must be false otherwise true


'DATE format
DateFormat = "mm/dd/yyyy"		 			'Format fro dates. The seperator style is not important. use d/mm/yyyy , dd/mm/yyy ,  Cant have single day sinlge m like m/d/yyy then u have to use the SplitDateBy
SplitDateBy=""							    'only use this if dateformat does not work. This is plan B.	
DateisText="" 					    		' [JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC] leave empty to ignore
RemoveXcharsfromDate = 0					 'Default to 0; removes X chars from the start of date. so ="1" will change  120151201 to 12151201
DaysSincelastMovementObsolete = 730		'number of days after which part is considered obsolete, f.i. 730 = 2 years
DaysSincelastMovementScrap= 1460	'number of days after which part is considered scrap, f.i.1460 4 years
MinDaysInStock = 0						   	'number of days a part needs to be in stock before it is included in current. Use 0 to ignore. Can use this to ignore current stock, by setting it to same value as DaysSincelastMovementObsolete + 1 day.
	
'Date values
usedatelastPurchase = "true" 				'choose datelastpurchase or datelastsale, set to false if you want datelastsale or keep as true. If empty it will default to the other, if other is empty it will set to the 2 below settings.
SettingsForEmptyDate = current  		 	'use current, scrap or obs , so if date is empty we put this date into current , scrap or obs
SettingsForInvalidDate = current 		 	'use current, scrap or obs, hopefull we never get this far. if so, contact IT to fix this script

'Date categories
ObsCategories = ""         				    'seperate outlets with comma  When Using PartsCategory Category, all other parts categories  are current unless there is a match below 
ScrapCategories  = ""	
ExcludeCategories = ""	

'' Edit Partid's
RemoveXfromstart = 0							'1st, removes X chars from start of partsID. Default to 0. So ="3" in "12346789" will be 456789
FindStartingCharandRemoveAllBefore =""			'2st, Removes all chars before the specified character. So ="P" in "1234P56789" will be "56799"
FindStartingCharandRemoveAllAfter =""			'2.1st, Removes all chars after the specified character. So ="P" in "1234P56789" will be "1234"
FindStartingCharsandRemovePartID=""			   '3st, Removes whole partid if starts with x. So ="P" in "Poil2332" will be removed, USE a comma , to add multi options.
FindStartingCharandRemoveSpecifiedChars=""      '4th, Remove all specified chars from beg of string, left to right if it matches.
FindFirstInstanceofAndRemoveit=""				'5th  Remove first instance of this from the partid. So  ="x" in "123x456x78" will be 12345x78
FindsubsequentcharsBeginingatPostionX=""	    '6th.Removes all subsequent	chars after postion x. So ="p,3" 	in	"XR8ppp58044" will be XR858044. You need a dimlimiter here. 1ST IS CHAR LOOKING FOR, 2ND IS STARTATPOSITON
MatchAndReplace =""							    '7th, removes this chararcter from anywhere in the partsID. Use , demlimiter to add extra items to remove. Items can be more than one char
ReplaceStartChar = "" 							'8th, when partsID starts with this character then a number of characters (specified below) will be removed
ReplaceStartCharCount = 0						'how many chars to remove form the start of the partsID when it finds a match with the above (including the matched character) .
RemoveLeadingZerosWhenPartIDisAllNumbers=""  '9 Default is false or "", removes all leading zeros only if the partid Contains all numbers, it will leave partid the same if it contains any letters. So 000atu22 will stay same. while 0001234 will become 1234

Const InsertintoStartofPartID =""
Const InsertintoEndofPartID =""

''Filters
minprice = 0								    'will set minprice to be included in selection, you need to select column no. for S_retailprice to have to to work. Default is 0
Const FilterByMatching =""					    'Leave empty to ignore. To use this, we need a "Col2=FilterBy char width 15" ect.. to work ,We can use a , to have multiple values. A value here means , it will only inlcude results matching these values
Const AndFilterByMatching=""					'We can have where FilterBy='OES' and AndFilterBy='OEM' "
Const AndNotFilterBy=""							'We can have where Filterby='OES' and AndFilterby='OEM' and Not AndNotFilterBy='Recondition'
Const PartidMatchStartingWith =""               'will read the partid column, and only include these in the stockfile, sinlge dealer only! This is almost like the 'BrandsInPartID' varible, but this one will include spaces.
Const subbrandfilter=""							' extended interface, only include these subbrands in output
MinCharactorAllowedPartid = 2					'only includes partID's with a min of x chars
MaxCharactorAllowedPartid = 20  				'only includes partID's with a max of x chars\

'QTY
const divideqtyby = 1							' will divide numberinstock by x					
NumberinStockisInteger ="true"			'If culture issue arrise, or $ symbol. set the numberinstock to a string. Some cultures use comma's in integer, which can cause issues. Here we can treat the numberinstock as a string. Or $ symbols.
Const MaxStock = 0								' 0 is ignored, if 0+n then anything > 0+n will be ignored -  So sets limit on max qty

'BATCH FILE WORK
RemoveEmptyScrapObsCurrent = "true"   			'delete the scrap obs and cur files if they are empty 

'SAVE FILES
Const FileSourceLocation =""	' END WITH \ PLEASE! . Change only on dealers PC, leave empty for FTP server. Will copy the stockfile to same location as splitfile
Const FileLocationSave = ""						'Change only on dealers PC, leave empty for FTP server. examples:  "" or "c:/batman/robin/"
WriteDisposition= "false"                  'if "true" the selection criteria are included in the output files
const ObsoleteFile = "obs.csv"                  
const CurrentFile = "cur.csv"                   
const ScrapFile = "scrap.csv"                        
Fieldseperator = ";" 		  '' vbTab for tab no quotes, or else "," for others

'XLS						
Const xlsSheetNumber = ""      			    'use numbers to identify your sheet, 1,2,3 ect..
ConvertXLSToTXTTAB =""	    	 		'Will convert xls to csv. Set to true or leave empty. If your xls is giving problems, then it is beter to set this to true to convert xls to csv
Const ClientSideXLSTOCSVconvertion = ""  ' We only use this option if we cant open xls files on parex server. This means client is using pivot table or linking, so we just need the raw csv conversion
Const XLSBadDataIgnore=""						'XLS null values occur when Integers are mixed with strings.Tthis results in bad null values. Set to false to ingnore these values, and exclude them from stockfile. Beter option is to convert to csv file
Const ColumnFormatForQTY = "0.00"  				'CAN USE "Text" or "0.0"
Const deletecolumns=""                        ' work backwards AF,AE,AD,AC,AB,AA,Z,Y,W,V,U,T,P,N,M,L,K,I,H,G,F,E,D,C
xlsSheetName = ""      	  'Can have multiple sheets,seperate sheet names with commas. Use this for real XLS and not covererted to csv. We like to use sheetno above!

'XML
ConvertXMLtoCSV = "false"					'Default ="false" , Set to "true". Can now convert XML to csv. false is default. Only does one xml file for now. need to test this because of version 112 change, not tested on xml

Const MergeMultiLine= ""  '' works only on xls files for v95. upgrades needed for dif files. Insert the comma delimeter of qty on firt line. then second lines will be merged if first line is > 0. Check the 1.csv file for the delimeter reading, do not check the xls file

''''''''''  PART FARM 
Const discountcurrent = 0					'' use .1 for 10%    discount Settings on retail price. You need to insert a value into S_retailprice
Const discountobs = 0							'' use .3 for 30%
Const discountscrap = 0	

Const DiscountSubbrand = 0      			      '' used for special brands that want a different price. This will be default discount for all
Const DiscountPriceXDays = 0  			     	  '' We can add in an even greater discount on x days for special brands
Const DiscountSubbrandOlderThanX= 0 		      '' used for special brands that want a different price. Price;Days
Const DiscountSubbrandsList=""  '' list of sub-brands to increase the price on use comma , to separate
'''''''''' END PART FARM


'#####################################################################
' Certificate script masters only! 
'#####################################################################


doublerows = "false"   '' For really bad stock files with double rows. Set to true if there are double rows, they will be removed in the Regex Function.
removeColumnXfromTextFile = "" '' Default is "0" or "" to ignore. The values for the ''Schema Columns  will change after you have deleted a column just like xls.,
' The delete columns are tricky, after every col del the page gets reset with missing column. So partid,qty,rubbish,date. must be  ""1,2" 


' REGEX	- will make a new FILE! DOES A CLEAN! This works on the File not the partID only
const RegPattern = ""						'this is to remove all character (sequences) . It will write a new file for you with the extrated match 
const Replacewith = ""                      'matches are replced with these characters
'' always explain your regex here:  ^  then look for this char 0 then use * to repeat  ''


showerror = "false"''  this will show the partid that gives the error. please set to false by default, if false all date errors will just be invalid, for testing this should be true

'' WSL(WebSite Single Login) imports schema data for wsl
ImportDealerInfo = "false" 	'' This script will import some settings from other dealers in a file called DealerInfo.schema								''	DaysSincelastMovementObsolete =730 DaysSincelastMovementScrap= 1095 usedatelastPurchase= true UseDate =  false ObsCategories =   "NORMAL,AUTOMATICPHASEOUT"     				 ScrapCategories  = 	"MANUALORDER" ExcludeCategories = ""
UsePartCategoryAsDefault = "false"   '' if no dealerinfo is avalible on FTP dealer folder, we will use the default settings from this splitfile.
UseDate =  "true"					'' true is you want lastpurchasedate and lastsaledate to calculate your obs,cur and scrap. Fasle is you want to use the partscategory

Const skipxrows = 0						    'skips first x rows in stock file
'#####################################################################
' Below  is forbidden! 
'#####################################################################
	
CallSettings()

	for  t = 0 to ubound(filelist) '' Loop the files to process


				filetobesplit = filelist(t)'' moved here because differant stock file will have differant brands

				if  AllowStockFikeToBeAutoRenamed ="false"  then
						if not fs.FileExists( filetobesplit) then  
							''msgbox("no file")
							exit for
						end if
				end if
				
				
				brand =split(GetOutLetsOrBrands("brand"), ",")
				dealer=split(GetOutLetsOrBrands("outlet"),",")  '' dealers = outlets!
				
				if ubound(brand) + ubound(dealer) > 1 and FileLocationSave ="" then
					if multistock ="true" then
						dim tsave:  
						tsave =  fs.getfolder(".") & "\"
						if tsave =  FileSaveLocation  then
							FileSaveLocation = fs.getfolder(".") & "\results\"
						end if
					end if
				end if

				if InStr(filelist(t) , "." ) > 0 then  ' Xls sheet names do not have  extentions. so we add it here
						filetobesplit =  Cstr(filelist(t))
				else
						if xls ="true" then			
							filetobesplit =origonalxlsfilename  			''xls that was not converted to csv 
						else
							filetobesplit =  Cstr(filelist(t) + ".csv") 	''xls that was converted to csv
						end if
				end if
	
			Set conn = createobject("ADODB.Connection")
			Set rs = CreateObject("ADODB.recordset")
		

			if fs.FileExists( filelocation & filetobesplit) then  

						if xls = "true" then			
					dim odbcPath	
							odbcPath = "DRIVER={Microsoft Excel Driver (*.xls)}; IMEX=1; HDR=YES; Excel 8.0; DBQ=" & filelocation & filetobesplit & ";"    
						'xlsx driver 2007+
						if InStr( lcase(filetobesplit),".xlsx" ) > 0  then
							odbcPath= "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; DBQ=" & filelocation & filetobesplit & ";"  
				''			odbcPath "Driver={Microsoft Excel Driver (*.xls)}; Dbq=" &  filelocation & filetobesplit   & "; DefaultDir=" &  filelocation & filetobesplit 
						end if 				
					Conn.Open (odbcPath )
				elseif  dbf ="true"  then
							Conn.Open ("DBQ="  & filelocation  & ";DefaultDir=" & filelocation & ";Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277") 						
				else
					Conn.Open ("DBQ=" & filelocation & ";DefaultDir=" & filelocation & ";Driver={Microsoft Text Driver (*.txt; *.csv)};DriverId=27;Extensions=None,asc,csv,tab,txt;FIL=text;")
				end if
				
				rs.ActiveConnection = Conn
				
				if  ubound(dealer) = -1 then '' msgbox possible errors
						msgbox("nothing! please look at your Delimiter. Please alert colin whenever you see this and it is something other than a wrong  delimeter")
				end if
				
				for i = 0 to ubound(dealer) 

					if  dealer(i) = FileSaveLocation then
							if dealer(i) ="results" and len(FileSaveLocation) > 2 then
								dealer(i) = ""
							end if
					end if
					


					if ImportDealerInfo ="true" then ' WSL
						CallDealerSchemaInfo("20"  & DealerDic.item(Dealer(i)))
					end if	

					For j= 0 to Ubound(brand) '' must at least have one value

							if not fs.FolderExists (FileSaveLocation & Multifiles(filelist(t)) ) then
								fs.CreateFolder FileSaveLocation & Multifiles(filelist(t))	
							end if

							if not fs.FolderExists (FileSaveLocation & Multifiles(filelist(t)) & dealer(i)) then
										fs.CreateFolder FileSaveLocation & Multifiles(filelist(t)) & dealer(i)
							end if

							if Ubound(brand) => 0 then
								if not fs.FolderExists (FileSaveLocation & Multifiles(filelist(t)) & dealer(i) & "\" & brandloop(brand(j))) then
										 if not dealer(i) = "files" and  len(brandloop(brand(j))) > 1 then
											fs.CreateFolder FileSaveLocation & Multifiles(filelist(t)) & dealer(i) & brandloop(brand(j))	
										  end if
								 End if
							end if

								''  if save location has Files and there is also a dealer/outlet folder, then we remove the files folder from save location.
									if dealer(i) ="files" and len(brandloop(brand(j))) > 1 then
											if not fs.FolderExists (FileSaveLocation & Multifiles(filelist(t))  & brandloop(brand(j))) then
													fs.CreateFolder FileSaveLocation & Multifiles(filelist(t)) & brandloop(brand(j))		
											End if
										obsFilePath = FileSaveLocation & Multifiles(filelist(t))  & brandloop(brand(j))  & "\" & ObsoleteFile
										curFilePath = FileSaveLocation & Multifiles(filelist(t))   & brandloop(brand(j))  & "\" & CurrentFile
										scrapFilePath = FileSaveLocation & Multifiles(filelist(t))  & brandloop(brand(j)) & "\"  & ScrapFile
									else
										obsFilePath = FileSaveLocation & Multifiles(filelist(t)) & dealer(i) & brandloop(brand(j))  & "\" & ObsoleteFile
										curFilePath = FileSaveLocation & Multifiles(filelist(t))  & dealer(i) & brandloop(brand(j))  & "\" & CurrentFile
										scrapFilePath = FileSaveLocation & Multifiles(filelist(t)) & dealer(i)  & brandloop(brand(j)) & "\"  & ScrapFile
									end if
											
							'create files
							fs.createTextFile obsFilePath, "true"
							fs.createTextFile curFilePath, "true"
							fs.createTextFile scrapFilePath, "true"
							
					
							'open file connctions
						Dim tsObs, tsCur, tsScrap 
							Set tsCur = fs.openTextfile(curFilePath, 8 ,true , -1)
							Set tsScrap = fs.openTextfile(scrapFilePath, 8,true , -1)
							Set tsObs = fs.openTextfile(obsFilePath, 8 ,true , -1)
						
						
						if Writedisposition then
								disposition = "Obsolete after " & daysSincelastMovementObsolete & " days" &", scrap after " & daysSincelastMovementScrap & " days"
								tsScrap.writeline disposition & Formatdatetime(now(),1)
								tsObs.writeline disposition & Formatdatetime(now(),1)
								tsCur.writeline disposition & Formatdatetime(now(),1)
						end if


						''  SINGLE DEALER
						if ubound(dealer) = 0 then 
						
								if xls = "true" then
									rs.Open "SELECT * FROM [" & filelist(t) & "$]", , 3 	
								else							
										if dealer(0) ="files" or dealer(0) ="results" then
											
											if dbf ="true" then
											'' dbf files are now handled, i did not include all the vars, too much work can add in as we need. this is a good start
													'rs.Open "SELECT * from  "  & """" & fileToBesplit & """" , , 3
													'rs.Open "SELECT * from [" & fileToBesplit & "]"  , , 3
													'rs.Open "SELECT * from [" & fileToBesplit & "] where "& S_numberInStock & " > 0  ", , 3
												rs.Open "SELECT * from  "  & """" & fileToBesplit & """"  & getnumberinstock & CheckMinPrice() & getsqlParttIDmaxminmaxlengh &  CheckBrands(brand(j)) , , 3
											else
											
												rs.Open "SELECT * FROM "  & """"& fileToBesplit &""""& getnumberinstock  & CheckMinPrice()  & getsqlParttIDmaxminmaxlengh & getsqlPartFilterBy   & FileterSubBrands() & CheckBrands(brand(j)) &    ";", , 3
											end if
									
										else
												rs.Open "SELECT * FROM "  &""""& fileToBesplit &""""&  getnumberinstock  & CheckMinPrice()  & getsqlParttIDmaxminmaxlengh & getsqlPartFilterBy &  CheckMinPrice() &  FileterSubBrands() & CheckOutlets(dealer(i)) &  CheckBrands(brand(j)) & ";", , 3
										end if
								end if	
				
						'' MULTI DEALERS
						else 
						
						
							if xls = "true" then
								rs.Open "SELECT * FROM [" & filelist(t) & "$]", , 3	
							else
									if multistock ="true" then '' double qty column used.
										rs.Open "SELECT * FROM "  & """"& fileToBesplit & """" & " where 1=1 " &    CheckMinPrice() &  getsqlParttIDmaxminmaxlengh &  getsqlPartFilterBy &  FileterSubBrands() &  CheckBrands(brand(j)) &    ";", , 3			
									else
										rs.Open "SELECT * FROM "  & """"& fileToBesplit & """" &  getnumberinstock &  CheckMinPrice()  &   getsqlParttIDmaxminmaxlengh &  getsqlPartFilterBy &    FileterSubBrands() & CheckBrands(brand(j)) &   CheckOutlets(dealer(i)) &  ";", , 3			
									end if
							end if
						end if
						
						if rs.recordcount > skipxrows then
							if skipxrows > 0 then
								Rs.move skipxrows  
							 end if
						 end  if

							Do While Not rs.EOF 
						
								if  RegexReturnInt(rs.Fields(csvorxls("numberinstock") +  Cstr(QTYasDealers(i)))) >= MaxStock  then	''csv
									isObsolete = "false"
									isScrap = "false"
									isReject = "false"
									strLastmovement = null
									lngLastmovement = 0
									
									numberinstock= RegexReturnInt(rs.Fields(csvorxls("numberinstock") +  Cstr(QTYasDealers(i))))	'csv
									
						
									
							
									'' 1. check to user lastpurchase or last sale  . even if using category for date, we still want to convert dates
										Dim lastP ,lastS
										lastS = 0
										lastP = 0
									
										if UseCategory   then '' usecat has been maliputed by the  WSLUsePartCategory
											PickDateToUse =  rs.Fields("partscategory")
											
									
										else
										
											if Not S_LastPurchase ="" then
												If not IsNull(rs.Fields(csvorxls("lastpurchase")))  Then
													lastP  =  RegExReturnOnlyIntegers(rs.Fields(csvorxls("lastpurchase"))) '' cleans date string 
												
												end if 
											end if

											if Not S_LastSale ="" then
												If  Not IsNull(rs.Fields(csvorxls("lastsale"))) Then
													lastS = RegExReturnOnlyIntegers(rs.Fields(csvorxls("lastsale"))) '' cleans date string 
												end if
											end if
										
										    PickDateToUse =	UseLastPurchaseORLastSale(lastP , lastS)
											lngLastMovement= convertdate(PickDateToUse, rs.Fields(csvorxls("partsID")))	
										end if
								
										if UseCategory then '' we want to use category for obs,cur,scrap
											dim partsCategory 
												if not isnull(rs.Fields(csvorxls("partsCategory")))  then
													
												partsCategory=LCase(rs.Fields(csvorxls("partsCategory")))				
													''Do while Ubound(ObsCategories) + Ubound(ObsCategories) + Ubound(ObsCategories)
													dim cats
													For cats = 0 to CatLengthcombo '' we have allready added the combo length of all arrays
															if FindMatch(ObsCategories,partsCategory)= "0" then 
																isObsolete="true" 
																exit for
															end if
															if FindMatch(ScrapCategories,partsCategory) ="0" then 
																isScrap="true"
																exit for
															end if
															if FindMatch(ExcludeCategories,partsCategory) ="0"  then isReject = "true"
															'' else if no matched, then we use "SettingsForEmptyDate" var check that this works
													next
												end if
										end if
								
			
									if isReject = "false" then '' default is false this is for partscategory
										
											if not UseCategory then '' use normal lastpurchase and lastsale
												
												if lngLastMovement < (today - daysSincelastMovementScrap) then isScrap= "true"
												if lngLastMovement < (today - daysSincelastMovementObsolete) then isObsolete="true"
										
															
											'' v99 Special Discount, declare here, maybe use it later if subbrands match your special list
												DiscountSubbrandMore = PartFarmSubbrandDiscounts(lngLastMovement)	
											

												if  lngLastMovement < 2 then '' 0 is empty  1 is invalid data
													isObsolete = CheckisObsolete(lngLastMovement)
													isScrap = CheckisScrap(lngLastMovement)
												end if
											end if
											
											'' OVERRIDE SCRAP OBS AND CURRENT FOR PARTFAMR SUBBRANDSS
											'******************************************************************************************************************
						
											If lnglastmovement <= today - MinDaysInStock or  UseCategory then 'to filter out possibly reserved parts

											dim ext
												dim record
												
													if multistock ="true" then '' no dealers here, we use numberinstock as dealers
											

													record = PartsID(rs.Fields(csvorxls("partsID"))) & Fieldseperator &  divide(clng(rs.Fields(csvorxls("NumberInStock" &  Cstr(i + 1))))) & Fieldseperator  & PickDateToUse  & ExtendedInterface(rs)
													else
										
												    record = PartsID(rs.Fields(csvorxls("partsID"))) & Fieldseperator  &  divide(clng(rs.Fields(csvorxls("NumberInStock"))))  & ExtendedInterface(rs) & Fieldseperator 	 & PickDateToUse

							
												end if
							
												if Instr(record , Fieldseperator) > 1 then '' Sometimes a partid is null, we skip
														
												
														If isScrap = "true" Then
																tsScrap.writeline record
																counterScrap = 1
														ElseIf isObsolete = "true" Then
																tsObs.writeline record
																counterObs = 1
														Else
																tsCur.writeline record
																counterCurrent = 1
														End If
													end if
														
											end if
							
								end if
										
				end if
				
							if dbf ="true" then
								on error resume next
							end if
				
							rs.MoveNext	
							
							
							Loop

							tsObs.Close
							tsCur.Close
							tsScrap.close
							rs.Close
							
						RemoveEmptyScrapObsCurrentFiles()
				
						clearcounter()
				
					next 
				next

				Conn.Close
				RenameStockFileEnd()
			
			end if
	next
	
			Set tsObs = Nothing
			Set tsScrap =Nothing
			Set tsCur = nothing
			Set rs = Nothing
			Set Conn = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCTIONS ''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCTIONS ''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCTIONS ''''''''''''''''''''''''''''''''''''''''''''''''''

''functions for schema
Function CallSettings()
Set fs = CreateObject("Scripting.FileSystemObject")

MaxScanRows = 35						' max number of invalid rows to skip. Think we can phase this out

if S_Brand=""  and  len(BrandsNotinColumns) > 0 then
	CreateCompleteNewBrandStockFiles()
end if

	MoveFileSourceLocation()
	CheckForErros()
	ImportDealerInfoSchema()'' WSL IMPORTS DEALER SCHEMA SETTING FROM OTHER DEALERS
	
	if  ImportDealerInfo  = "false" then '' for importing we need both partscat and lastp date 
		if (len(S_LastPurchase) + len(S_LastSale)) > 0 and  len(S_PartsCategory) > 0 then
			msgbox("cant use both lastpurchase and partscats")
		end if
	end if


	if  Instr(FileName, ".dbf") > 1  then
		dbf = "true"
	end if
	
	
	ConvertPDFToText()
	
	CheckDoubleRows()
	
	RegeXRun()  

	fileToBeSplit = stockfilename() 
	
	

	RenameStockFile() '' we need to rename the stockfile from dealer to match that of our schema file
	
	ConvertXLStoCSVFiles() 	'' need to run the convert tool first, before schema	


	ConvertXMLtoCSVFiles()	
	
	'' new v 95 merger 2 lines
	''MultiMergLine()
	
	
	CreateSchema()
	

	fixdateformat()
		
	'brand =split(GetOutLetsOrBrands("brand"), ",")
	'dealer=split(GetOutLetsOrBrands("outlet"),",") 

		if not len(MatchAndReplace) = 1 then
			matcharray ="true"
			MatchAndReplace = split(MatchAndReplace, ",")
		else
			matcharray ="false"
		end if

		if len(FindsubsequentcharsBeginingatPostionX) > 0 then
			Dim validstrting
			validstrting = split(FindsubsequentcharsBeginingatPostionX, ",")
				if ubound(validstrting) = 0 then
						msgbox("FindsubsequentcharsBeginingatPostionX IS NOT VALID, u need a Char and a counter")
					else
					Dim  subsquentarray
						subsquentarray = split(FindsubsequentcharsBeginingatPostionX, ",")
		
							LookingFor = subsquentarray(0) '' remember to auto enlarge this.
							StartPostionAT = subsquentarray(1)
				end if
		end if 

	ObsCategories= split(Lcase(ObsCategories),",")
	ScrapCategories= split(LCase(ScrapCategories),",")
	ExcludeCategories= split(LCase(ExcludeCategories),",")
	
	
	CatLengthcombo = CatLength()
	UseCategoryforDates() '' set usecategory to true or false
	today = year(now())*365 + month(now()) * 30 + day(now())
	Set conn = createobject("ADODB.Connection")
	Set rs = CreateObject("ADODB.recordset")
	Set fs = CreateObject("Scripting.FileSystemObject")
	Filelocation =  GetFolderLocation()    '' Gets Folderlocation, default is empty which means splitfile is in same folder as stock files
	
	FileSaveLocation = GetFileSaveLocation()
	checkpartidlenght = CheckPartIDSize()   '' function will check to see if we need to validate the size of parttid stinng size,
	getnumberinstock =  GetNumberinStockSQL() '' is numberinstock a string or integer ? choose sql accordinaly
	

	getsqlPartFilterBy = CheckFilterByClauses() '' INsert EXTRA sql select by clause

	counterCurrent = 0 
	counterScrap = 0 
	counterObs = 0

	dateformat =  DateFormatClean(dateformat)    '' cleans the dateformat,removes bad chars
	SetextInterfaceused = ExtInterfaceused()
	

		if  ubound(filelist) > 0 then '' we have mutlifiles/sheets to split
			multisheets ="true"
		end if

	if multistock ="true" and xls="true" then '' DOES NOT ALLOW multi qty columns
		msgbox("please convert to csv to enable a double numberinparts column")
	end if
	
	
	if discountscrap +  discountobs + discountcurrent > 0 then
	discount = "true"
else
	discount = "false"
end if
end function



function CreateCompleteNewBrandStockFiles()

		Dim createdfile , cf , lineread ,listFile ,currentbrand  , CurrentFile  ,objExtension , createnewfilename
		currentbrand = ""

'' CREATE THE NEW STOCK FILES


objExtension = Right(filename,4) 


			createdfile	= split(BrandsNotinColumns,",")
				for cf = 0  to Ubound(createdfile)
					fs.createTextFile FileSaveLocation & LCase(createdfile(cf)) & objExtension, "true"
					Set CurrentFile = fs.openTextfile(FileSaveLocation & LCase(createdfile(cf)) & objExtension, 8 ,true , -1)
					createnewfilename =  createnewfilename + createdfile(cf) &  objExtension & ","
				next 
				
				createnewfilename = Lcase(Left(createnewfilename, len(createnewfilename) -1))
				Set listFile = fs.OpenTextFile(FileName)
				
					Do until listFile.AtEndofStream
					
						lineread =  LCase(listFile.ReadLine())
						
								
								for cf = 0  to Ubound(createdfile)'' loop brand array
									 If  InStr(lineread ,LCase(createdfile(cf))) > 1 Then '' change brand , CHANGE CURRENTFILE
									 
									       
											currentbrand = LCase(createdfile(cf))
											
											Set CurrentFile = nothing
											Set CurrentFile = fs.openTextfile(FileSaveLocation & LCase(createdfile(cf)) & objExtension, 8 ,true , -1)
											
												CurrentFile.WriteLine  lineread
												
			
									else
									
										if LCase(createdfile(cf)) = currentbrand then
											CurrentFile.WriteLine  lineread
											end if
												
									End if	
									
								next 
												
					loop
					listFile.Close
					Set listFile = Nothing
					Set CurrentFile = nothing

FileName =createnewfilename

end function

Function CheckForErros()

usedatelastPurchase =  Lcase(usedatelastPurchase)
BrandWildcard =  Lcase(BrandWildcard)
usedatelastPurchase =  Lcase(usedatelastPurchase)
NumberinStockisInteger =  Lcase(NumberinStockisInteger)
RemoveEmptyScrapObsCurrent =  Lcase(RemoveEmptyScrapObsCurrent)
WriteDisposition =  Lcase(WriteDisposition)
ConvertXLSToTXTTAB =  Lcase(ConvertXLSToTXTTAB)
ConvertXMLtoCSV =  Lcase(ConvertXMLtoCSV)
ImportDealerInfo =  Lcase(ImportDealerInfo)
UsePartCategoryAsDefault =  Lcase(UsePartCategoryAsDefault)
UseDate =  Lcase(UseDate)
showerror =  Lcase(showerror)
AllowStockFikeToBeAutoRenamed =  Lcase(AllowStockFikeToBeAutoRenamed)
doublerows = Lcase(doublerows)
	
	

		if AllowStockFikeToBeAutoRenamed="true" and  len(BrandsNotinColumns) > 0 then
			Wscript.Echo "allowstockfiletorenamedtomatchschema must be false,beause you will be using dynamiclly be using multiple stockfile"
			WScript.Quit	
		end if
		
		if AllowStockFikeToBeAutoRenamed="true" and FileName ="false" then
			msgbox("cant have filename false and automatic file rename set to true. ")		
		end if
		
	
		if  AllowStockFikeToBeAutoRenamed ="true" and Instr(FileName,",") > 0 then
		Wscript.Echo "cant rename multiple stock files, set AllowStockFikeToBeAutoRenamed to false and ensure filenames are correct, use prebatch file if needed"
		WScript.Quit
		end if
	
		
		
		if  ConvertXLSToTXTTAB ="true" and Instr(FileName,".xl") = 0 then
		Wscript.Echo "cant mix xls and text files please check filename and ConvertXLSToTXTTAB"
		WScript.Quit
		end if
		
		
		'' check date format
		if Not Instr(LCase(DateFormat),"mm") > 0  and Not Instr(LCase(DateFormat),"dd") > 0  then
			if len(SplitDateBy) = 0 then
				msgbox("cant have d/m/yyyy format withouth splitbydate value.")		
			end if
		end if
		
			
		if len(S_Brand) > 0 and  len(BrandsNotinColumns) > 0 then
			Wscript.Echo "you cant have s_brand value and brandsinotincolum, s_brand is for columns brandsnotincolum is for brand in random lines"
			WScript.Quit
		end if
		
		

		IF len(brands) > 0 and len(BrandsNotinColumns) > 0 then
		if brands <> "automatic" then
		
					Wscript.Echo "if you are using brandsnotincolumn, then please leaave brands empty"
					WScript.Quit
					end if
		end if
				
	''check the schema settings
	if  (FormatDelimited = DecimalSymbol) or (FormatDelimited =  DidgetGroupingSymbol)   then
		Wscript.Echo "You cant split columns with FormatDelimted ad Decimal or DigetGroupingSymbol being the same"
		WScript.Quit
	end if
	
	


	if 	minprice > 0 and len(S_retailprice) = 0  then
		Wscript.Echo "cant have minprice and empty s_retailprice"
		WScript.Quit
	end if
	
	if len(xlsSheetName) = 0 and len(xlsSheetNumber) = 0 and  ConvertXLSToTXTTAB="true" then
			Wscript.Echo "cant have no sheet name and also ConvertXLSToTXTTAB to true"
			WScript.Quit
	end if
	
	'xls						
	if FormatDelimited = DecimalSymbol then
		Wscript.Echo "cant have FormatDelimited and DecimalSymbol the same"
		WScript.Quit
	end if	
	
	'' for wsl we can have both partscat and lastsales
	if ImportDealerInfo="false" then
	
	
			if len(S_partsID ) = 0 or ( len(S_numberInStock  ) = 0 ) then
			Wscript.Echo "you need s_partdid or s_numberinstock must have values"
			WScript.Quit
		end if
	
		if len(S_PartsCategory) > 0 and ( len(S_LastSale ) > 0 or len(S_LastPurchase ) > 0) then
			Wscript.Echo "can not use both partscategory and lastsales dates"
			WScript.Quit
		end if
		
		if  (len(S_LastSale ) > 0 or len(S_LastPurchase ) > 0) and  (len(ObsCategories) > 0 or len(ScrapCategories )>0) then
		
		Wscript.Echo "can use both lastsales and obscategories together"
			WScript.Quit
		end if 
		
		''Filters
		
		if len(FilterByMatching ) > 0 and  S_Filterby ="" then
			Wscript.Echo "you need to choose the column S_FilterBy to match the data with FilterbYmATCHING"
			WScript.Quit
		end if
		
		if len(S_Filterby ) > 0 and   FilterByMatching ="" then
			Wscript.Echo "you need to choose the column S_FilterBy to match the data with FilterbYmATCHING"
			WScript.Quit
		end if
		
		if len(AndFilterByMatching ) > 0 and  S_AndFilterBy ="" then
			Wscript.Echo "you need to choose the column S_AndFilterBy to match the data with AndFilterByMatching"
			WScript.Quit
		end if
		
		if len(S_AndFilterBy ) > 0 and   AndFilterByMatching ="" then
			Wscript.Echo "you need to choose the column S_AndFilterBy to match the data with AndFilterByMatching"
			WScript.Quit
		end if
		
	end if
	
	


End Function


Function MultiMergLine()

'' need to do this after xls conversion. then run the regex conversion afterwards

		if dbf = "true" then
			exit function
		end if


			dim fss, ts1, ts2, str, stringWriter, i, count
			files = split (FileName, ",")
			Set fss = CreateObject("Scripting.FileSystemObject")
			Filelocation =  fs.getfolder(".") & "\"
			count=0	
					for each file in files 
						if fss.FileExists(filelocation & file ) and instr(file,".xl") = 0 then
							fss.createTextFile filelocation & "temp.tmp", True
							Set ts1 = fss.openTextfile(filelocation & file, 1) '' source
							Set ts2 = fss.OpenTextFile(filelocation & "temp.tmp", 8) '' temp
							str = ""
							do while not ts1.atEndofStream
								str = ts1.readLine
								

								
								  If IsNumeric(mid(str,160,1)) Then
										
										  str=str & ts1.readLine & ts1.readline
										  Ts2.writeLine str
										  str=""
									   else
											str=""
										end if
								
								
								
							Loop
							ts1.close
							ts2.close
					 
						fss.CopyFile filelocation & "temp.tmp",filelocation & file,true
						fss.DeleteFile filelocation & "temp.tmp"
							count = count + 1
					
						end if
					next
	end function



Function MoveFileSourceLocation()

if len(FileSourceLocation) > 1 then

		if instr(FileSourceLocation,":") > 0 then
		
			If fs.FileExists(FileSourceLocation & FileName) Then
				fs.CopyFile   FileSourceLocation & FileName, fs.getfolder(".")  &"\"   
			End If
		else
			If fs.FileExists(fs.getfolder(".") & "\" &  FileSourceLocation & FileName) Then
			fs.CopyFile  fs.getfolder(".") &  "\" & FileSourceLocation & FileName, fs.getfolder(".")  &"\"   
		End If
end if



end if

End Function



Function  ImportDealerInfoSchema()

	if ImportDealerInfo = "true" then
		Dim arrLines   ,strData , splitdic , strLine
			if fs.FileExists(GetFileSaveLocation() &"\help\dealers.txt") then  

				Set DealerDic = CreateObject("Scripting.Dictionary")
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				strData   = objFSO.OpenTextFile(GetFileSaveLocation() & "\help\" & "dealers.txt", 1).ReadAll
				arrLines = Split(strData,vbCrLf)
				
				
				For Each strLine in arrLines
					if  len(strLine) > 4 then
						splitdic = split(strline, ",")
						DealerDic.add splitdic(0) , splitdic(1)
					end if
				Next
				Set objFSO = Nothing
			end if
		
			DaysSincelastMovementObsolete_origonal  = DaysSincelastMovementObsolete
			DaysSincelastMovementScrap_origonal = DaysSincelastMovementScrap
			ObsCategories_Origonal = ObsCategories
			ScrapCategories_Origonal  = ScrapCategories
			ExcludeCategories_Origonal	= ExcludeCategories
			UseDate__Origonal =  UseDate
			usedatelastPurchase_Origonal = usedatelastPurchase
			minprice_Origonal= minprice
	end if	
	
end function


'' IMPORTS SOME SCHEMAA SETTINGS FROM OTHER DEALERS
Function CallDealerSchemaInfo(ByVal DealerID)

	Dim arrLines1   ,strData1 , strLine1 ,  splitschema


	if fs.FileExists("..\"  & DealerID & "\DealerInfo.schema") then  

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strData1   = objFSO.OpenTextFile("..\"  & DealerID & "\DealerInfo.schema", 1).ReadAll
  
	   arrLines1 = Split(strData1,vbCrLf)
	   
					
					For Each strLine1 in arrLines1
						splitschema = split(strline1, "=")
							Dim schemakey, schemavalue
								schemakey   =  Lcase(Trim(splitschema(0)))
								schemavalue =  Lcase(Trim(splitschema(1)))
			

								Select  Case schemakey
										Case "dayssincelastmovementobsolete"
											DaysSincelastMovementObsolete = schemavalue
										Case "dayssincelastmovementscrap"
											DaysSincelastMovementScrap = schemavalue
										Case "usedate"
													if schemavalue = "true" then
														UseCategory = false	
													else
														UseCategory = true
													end if
													
										Case "usedatelastpurchase"
													if schemavalue = "true" then
														usedatelastPurchase = "true"	
													else
														usedatelastPurchase = "false"
													end if
										Case "minprice"
												minprice =  schemavalue
										Case "obscategories"
										Dim first, last 
											first = Instr(1, schemavalue ,"""")
												if first  > 0 then
													schemavalue = Right(schemavalue, len(schemavalue) - first)
													last = InstrRev(schemavalue ,"""")
													schemavalue = Left(schemavalue, last - 1)
												ObsCategories= split(Lcase(schemavalue),",")
												end if
												
											
										Case "scrapcategories"
											Dim firsts, lasts 
												firsts = Instr(1, schemavalue ,"""")
													if firsts  > 0 then
														schemavalue = Right(schemavalue, len(schemavalue) - firsts)
														lasts = InstrRev(schemavalue ,"""")
														schemavalue = Left(schemavalue, lasts - 1)
													ScrapCategories= split(Lcase(schemavalue),",")
													end if
										
										Case "excludecategories"
											Dim firste, laste 
												firste = Instr(1, schemavalue ,"""")
													if firste  > 0 then
														schemavalue = Right(schemavalue, len(schemavalue) - firste)
														laste = InstrRev(schemavalue ,"""")
														schemavalue = Left(schemavalue, laste - 1)
													ExcludeCategories= split(Lcase(schemavalue),",")
													end if
											
										End Select

					Next
					
		Set objFSO = Nothing
	else
	

			DaysSincelastMovementObsolete =	DaysSincelastMovementObsolete_origonal 
			DaysSincelastMovementScrap= DaysSincelastMovementScrap_origonal
			ObsCategories= split(Lcase(ObsCategories_Origonal),",")
			ScrapCategories =   split(Lcase(ScrapCategories_Origonal),",")  
			ExcludeCategories = split(Lcase(ExcludeCategories_Origonal),",") 
		    usedatelastPurchase = 	usedatelastPurchase_Origonal
			minprice = minprice_Origonal
			
					if UseDate__Origonal = "true" then
						UseCategory = false	
					else
						UseCategory = true
					end if
	end if
end function





	Function Multifiles(ByVal filen)
			
			if InStr(filen , "." ) > 0 then
				filen = Left(filen, Len(filen) - 4)
			end if

			
			
			if  multisheets ="true"  then 
					if  multifilekillrootfolder ="true" then '' this is for filename=false
						Multifiles = ""		
					else
						Multifiles =   "/" & filen & "/"  
					end if
			else
				Multifiles = ""		
			end if					
	End Function





Function ConvertXLStoCSVFiles() 
	if ConvertXLSToTXTTAB ="true" or  ClientSideXLSTOCSVconvertion ="true" then
		Dim splitdelcolumn , columD,useSheetName ,useSheetNo ,concatfilenames , Local
	'' FILE FORMATS http://msdn.microsoft.com/en-us/library/bb241279%28office.12%29.aspx
		if Instr(FileName, ",") > 0 then
			msgbox("Cant have multi XLS file, we can split sheets")
		else
			Dim	 objExcel , objExcelBook ,objWB , objWS , sh 
			if  (len(xlsSheetName) > 0) and (len(xlsSheetNumber ) > 0) then
				msgbox("cant have both xlsheetname and xlssheetnumber")
			end if
				useSheetName ="false"
			    useSheetNo = "false"
		' one function for both worksheet name and worksheet number
				if  len(xlsSheetName) > 0 then
					filelist= split(xlsSheetName, ",")
					useSheetName ="true"
				end if
				if len(xlsSheetNumber) > 0 then
				    filelist= split(xlsSheetNumber, ",")
					useSheetNo ="true"
				end if
				  	For sh = 0 to ubound(filelist)
						Set objExcel = CreateObject("Excel.Application")
						Set objWB = objExcel.Workbooks.Open(GetFolderLocation & "\" & FileName , Local= True)
						
						IF useSheetNo ="true" then
							Set objWS = objWB.worksheets(sh +1)
						else
							Set objWS = objWB.worksheets(filelist(sh))
						end if
						
						objExcel.application.visible=false
						objExcel.application.displayalerts=False
						splitdelcolumn	= Split(deletecolumns,",")	
							For columD = 0 to ubound(splitdelcolumn)
								objWS.Range("" + splitdelcolumn(columD) + ":" + splitdelcolumn(columD) +"").Delete
							next
						if len(S_retailprice) > 0 then objWS.columns(Cint(S_retailprice)).numberformat= ColumnFormatForQTY
						
				
						if len(S_description) > 0 then 
						
							objWS.columns(Cint(S_description)).replace "," ,""
							objWS.columns(Cint(S_description)).replace """" ,""
							objWS.columns(Cint(S_description)).replace """""" ,""
						''	objWS.columns(Cint(S_description)).replace chr(34) ,""
				
						''	objWS.columns(Cint(S_partsid)).replace chr(34) ,""
							
							
							
						end if
						objWS.columns(Cint(S_numberInStock)).numberformat= ColumnFormatForQTY
						
						objWS.saveas GetFolderLocation & filelist(sh) & ".txt" , 42'' colin change from 6 to 42 unicode
						set objWS = Nothing
						objExcel.application.quit
						set objExcel=nothing
						Set objWB = Nothing
						concatfilenames = concatfilenames + filelist(sh) + ".txt" + "," 
						''RegeXRunFile(filelist(sh) + ".txt" )
					next
				
					if  Instr(concatfilenames, ",") <>0 then concatfilenames = 	Left(concatfilenames,Len(concatfilenames)-1)
					filelist= split(concatfilenames, ",")	
		end If


	Set objWS = Nothing
	Set	objExcel = Nothing
	
	
	if ClientSideXLSTOCSVconvertion="true" then
				 wscript.quit
	end if
	
	
	
	
	end if
end function



Function ConvertXMLtoCSVFiles	() 

	if ConvertXMLtoCSV ="true" then
	
	
		if Instr(FileName, ",") > 0 then
			msgbox("Cant have multi xml files yet, contact colin if we need this")
		else
	
			Dim xlApp, xlWkb, SourceFolder,TargetFolder,filexml ,FullTargetPath
			Set xlApp = CreateObject("excel.application")
			Const xlNormal=1
			Const xlCSV=6

					filexml = GetFolderLocation & FileName
					TargetFolder= GetFolderLocation

					xlApp.Visible = false

						Set xlWkb = xlApp.Workbooks.Open(filexml)
						
						FullTargetPath=TargetFolder  & "parex.csv"
						
						xlWkb.SaveAs FullTargetPath, xlCSV, , , , , , 2
						xlWkb.Saved = True
						xlWkb.close
						Set xlWkb = Nothing
						Set xlApp = Nothing
						
						dim namelist
						namelist ="parex.csv"
					filelist= split(namelist, ",")
				
							
	end if
	end if
end function



Function QTYasDealers(ByVal i)
	if multistock ="true" then

			QTYasDealers = Cstr(i + 1)
			else
		if xls = "true" then
		QTYasDealers = 0
		else
		QTYasDealers = ""
			
			end if
	end if
end function


Function csvorxls(byval columnname)
	if xls ="true" then
			csvorxls = eval("x" + columnname)
		elseif dbf ="true" then
			csvorxls = eval("S_" + columnname )
		else
		csvorxls = columnname
	end if
end function

				''if  rs.Fields(csvorxls("numberinstock") &  Cstr(QTYasDealers(i))) > 0  then    ' works with csv txt
				''if  rs.Fields(csvorxls("numberinstock")) &  Cstr(QTYasDealers(i)) > 0  then     ' works with excell only
Function csvorxlsgroupqty(byval columnname ,byval count)
			if xls ="true" then
				csvorxlsgroupqty = eval("x" + columnname + QTYasDealers(i))

			''if  rs.Fields(csvorxls("numberinstock")) &  Cstr(QTYasDealers(i)) > 0  then     ' works with excell only

		else
		
		csvorxlsgroupqty = columnname &  Cstr(QTYasDealers(i))
	end if
end function



Function fixdateformat()
'' we  d/m/yyyy  this is for excell where d/m/yyyy is used, normally we dont use a split
	if not  splitdateby = "" then
		if  not InStr(dateformat,"dd") = 1 then
			dateformat = replace(dateformat , "d" , "dd" )
		end if
		
		if  not InStr(dateformat,"mm") = 1 then
			dateformat = replace(dateformat , "m" , "mm" )
		end if
	end if
	

end function

'' we use this for MaxStock aswell. Divide then set maxstock
Function divide(ByVal NumberInStock)
	
	
	
	if divideqtyby > 1 then
	

		divide = abs(NumberInStock/ divideqtyby)
	else
	

	
		divide = abs(NumberInStock)
		
		
		
	end if
	
	if MaxStock > 0 and NumberInStock >  MaxStock then
			divide =  MaxStock
	end if
	

	if divideqtyby < 1 then
		if  NumberInStock  > 9999     then
				if MaxStock > 0 then
					divide =  MaxStock
				else
					divide = 9999
				end if
		end if
	end if

	
end function


Function ExtInterfaceused ()
			If  LEN(S_subbrands) > 0 THEN
				ExtInterfaceused = "true"
				exit function
			end if
			If  LEN(S_description) > 0 THEN
				ExtInterfaceused = "true"
				exit function
			end if
			If  LEN(S_purchaseprice) > 0 THEN
				ExtInterfaceused = "true"
				exit function
			end if
			If  LEN(S_retailprice) > 0 THEN
					ExtInterfaceused = "true"
					exit function
			end if
			If  LEN(S_condition) > 0 THEN
				ExtInterfaceused = "true"
				exit function
			end if
			If  LEN(S_manufacture) > 0 THEN
				ExtInterfaceused = "true"
				exit function
			end if
	ExtInterfaceused = "false"
End Function

Function ExtendedInterface(ByVal rs)
		if SetextInterfaceused = "true" then
	
			Dim sqlstringbuilder
				If  LEN(S_subbrands) > 0 THEN
					
					If  IsNull(rs.Fields(csvorxls("subbrands"))) Then
						sqlstringbuilder = sqlstringbuilder + Fieldseperator
					else
						sqlstringbuilder = sqlstringbuilder + Fieldseperator + rs.Fields(csvorxls("subbrands"))
					end if

				end if
				
				If  LEN(S_description) > 0 THEN
				
	
			
					If  IsNull(rs.Fields(csvorxls("description"))) Then
						sqlstringbuilder = sqlstringbuilder + Fieldseperator
					else
					
					dim tempdesc 
				     tempdesc =  rs.Fields(csvorxls("description"))
					
					tempdesc =  replace(tempdesc , ">" ,"")
					tempdesc =	replace(tempdesc , ">" ,"")
					tempdesc =	replace(tempdesc , "<" ,"")
					tempdesc =	replace(tempdesc , "*" ,"")
					''tempdesc =  replace(tempdesc , """" ,"""""") do this in the regex file
					
					sqlstringbuilder = sqlstringbuilder + Fieldseperator +  tempdesc

					end if
					

				end if

				If  LEN(S_purchaseprice) > 0 THEN
					if len( rs.Fields(csvorxls("purchaseprice"))) >  6 then
							sqlstringbuilder = sqlstringbuilder + Fieldseperator +  Replace(rs.Fields(csvorxls("purchaseprice")), DidgetGroupingSymbol ,"")
				    else
						If  IsNull(rs.Fields(csvorxls("purchaseprice"))) Then
							sqlstringbuilder = sqlstringbuilder + Fieldseperator + "0" 
						ELSE
							sqlstringbuilder = sqlstringbuilder + Fieldseperator + rs.Fields(csvorxls("purchaseprice"))	
						end if
 
					end if
				end if
				
			
				If  LEN(S_retailprice) > 0 THEN
				
	
					if	IsNull(rs.Fields(csvorxls("retailprice")))  or ( len( rs.Fields(csvorxls("retailprice"))) >  10)  Then
								sqlstringbuilder = sqlstringbuilder + Fieldseperator + "0" 
					else
							if discount = "true" then
								Dim retailPrice 
							
									
								retailPrice = Replace(rs.Fields(csvorxls("retailprice")), DidgetGroupingSymbol ,"")  '' remove the special digngroup 
									
								


									Dim getprice , korting , orgprice ,priceincrease
										
										orgprice = CDBL(retailPrice)
										
											priceincrease  ="false"
											
											If isScrap = "true" Then 
													korting = Round(orgprice * Cdbl(discountscrap ),2)
														if Cdbl(discountscrap ) > 1 then
															priceincrease ="true"
														end if
											ElseIf isObsolete = "true" Then
													korting = Round(orgprice * Cdbl(discountobs),2)
														if Cdbl(discountobs ) > 1 then
															priceincrease ="true"
													    end if
											Else
													korting =  Round(orgprice * Cdbl(discountcurrent), 2)	
														if Cdbl(discountcurrent ) > 1 then
															priceincrease ="true"
														end if													
											End If
											
											'' below will override the above discounts based on price.
											if Len(DiscountSubbrandsList) > 0 and DiscountSubbrandMore > 0  then
													Dim branddiscount , e 
													branddiscount =split(DiscountSubbrandsList, ",")
													
													For e = 0 to Ubound(branddiscount) 
															if Lcase(branddiscount(e)) =  Lcase(rs.Fields(csvorxls("subbrands"))) then
															korting = Round(orgprice * Cdbl(DiscountSubbrandMore ),2)
																if Cdbl(DiscountSubbrandMore ) > 1 then
																	priceincrease ="true"
																end if
															end if
													next
																						
											 end if

	
											if priceincrease ="true" then
												sqlstringbuilder = sqlstringbuilder + Fieldseperator +  Cstr(korting )
											else
												sqlstringbuilder = sqlstringbuilder + Fieldseperator +  Cstr(orgprice - korting )	
											end if
					

							else
							
						
								sqlstringbuilder = sqlstringbuilder + Fieldseperator +  Cstr(rs.Fields(csvorxls("retailprice")) )
							end if

					
					
					end if
					
				end if
				
				
				
				If  LEN(S_condition) > 0 THEN
			
					If  IsNull(rs.Fields(csvorxls("condition"))) Then
						sqlstringbuilder = sqlstringbuilder +  Fieldseperator + "0"
					else
						sqlstringbuilder = sqlstringbuilder + Fieldseperator + rs.Fields(csvorxls("condition"))
					end if
					
					
				end if
				If  LEN(S_manufacture) > 0 THEN
			
					If  IsNull(rs.Fields(csvorxls("manufacture"))) Then
						sqlstringbuilder = sqlstringbuilder + Fieldseperator + "0"
					else
						sqlstringbuilder = sqlstringbuilder + Fieldseperator + rs.Fields(csvorxls("manufacture"))
					end if
					
				end if
			ExtendedInterface =  sqlstringbuilder	
			

		
		end if

End Function



	function clearcounter()
	
	if not  ubound(brand) > 0 or not ubound(dealer) > 0 then

		if (counterCurrent +  counterObs + counterScrap) = 0 then
			msgbox("all split stock files are empty, please check your ReadCharacter Set value, or format delimiter")
			WScript.Quit	
		end if
		
	
		counterCurrent = 0
		counterObs= 0
		counterScrap = 0
	end if	
	end function

	function brandloop(Byval brandchar)
			if brand(0) ="NOBRANDS"  then
				brandloop =""
			else
				brandloop = "\" & Cstr(brandchar)
			end if
	end function

Function numnerinstocksplit()
	   multistock = "false"
   			if not S_numberInStock + S_Outlet = "" then '' will only work if not a width delimtier and there are no outlets set
				if  InStr(S_numberInStock , ",") > 0 then '' this is a width or a mulit qty split ?
						if Not InStr(S_partsID , ",") > 0 then '' if partsid does not have a "," then this qty is a split. So we have two qty columns in this stock file. Problem TODO, what if this is a width? screwed
						multistock = "true"
						end if
				end if
			end if
end function

 Function  CreateSchema()
 
	 if dbf ="true" then
	 
	 filelist= split(FileName , ",")
		exit function
	 end if
		numnerinstocksplit()
		if  ConvertXLSToTXTTAB="" then
			if InStr(FileName,".xls") > 0 then
				xls = "true"
	
			end if
		end if
		
		Dim widthorcolumndelimeter  ' do we split via a width setting, or a delimeter
				if InStr(S_partsID,",") > 0 then
					widthorcolumndelimeter = "width"   ' width seperator
				else
					widthorcolumndelimeter ="delimeter" '' a easy delimiter seperator
				end if
	
			Dim aFile
				if fs.FileExists("schema.ini") then  
					Set aFile = fs.GetFile("schema.ini")
						aFile.Delete
				end iF
				
			'' convert htm to txt
			if InStr(FileName,".htm") >0 then
				origonalxlsfilename =  FileName
			
				Dim tfile
				tfile = left(FileName , len(Filename) - Instr(Filename, ".")  +1) & "txt"
				fs.copyfile FileName, Filelocation & tfile,3
				FileName =   tfile
				
		
			end if		
						
				
				
			'' are we splitting xsl or csv ?	
			if InStr(FileName , "xls" ) > 0 or InStr(FileName , "xml" ) > 0  then
		
						origonalxlsfilename = FileName 
						if ConvertXLSToTXTTAB ="" and  len(xlsSheetName) > 0 then
						'' this is for old way of using xls , we dont convert to csv here
							filelist= split(xlsSheetName , ",") 
						end if

			else
			
				filelist= split(FileName , ",")
			end if
			
			
				fs.createTextFile "schema.ini", True
				Dim os
				Set os = fs.openTextfile("schema.ini", 8)
				Dim strDict()
				Dim objKey
				Dim strKey,strItem
				Dim X,Y,Z

		
		for e = 0 to ubound(filelist)
		

				if InStr(filelist(e) , "." ) > 0 then
					FileName =  Cstr(filelist(e))
				else
						FileName =  Cstr(filelist(e) + ".csv") '' this is csv files created for excel
				end if
				
	
		
				if not xls ="true" then
					
					xls ="false"
						os.writeline "[" + FileName + "]"
						os.writeline "ColNameHeader= false"
						os.writeline "Format=" + getFormatdelimited()
						os.writeline "MaxScanRows=" + Cstr(MaxScanRows)
						os.writeline "CharacterSet" + Cstr(getCharacterSet())
							Dim Schema  ,dictKey	 ,    intSort ,dictItem
							Set Schema = CreateObject("Scripting.Dictionary")
							AddSchemaDic(Schema) '' Schema.Add "S_partsid", S_partsID
								dictKey  = 1
								dictItem = 2
								intSort =2
									Z = Schema.Count
									If Z > 1 Then
									  ReDim strDict(Z,2)
									  X = 0
										  For Each objKey In Schema
											  strDict(X,dictKey)  = CStr(objKey)
											  strDict(X,dictItem) = CStr(Schema(objKey))
											  X = X + 1
										  Next
										  For X = 0 to (Z - 2)
											For Y = X to (Z - 1)
												Dim inta, intb
													inta =  Cint(strDict(X,intSort))
													intb = Cint(strDict(Y,intSort))
														If inta  > intb Then
															 strKey  = strDict(X,dictKey)
															 strItem =  strDict(X,dictItem)
															 strDict(X,dictKey)  = strDict(Y,dictKey)
															 strDict(X,dictItem) = strDict(Y,dictItem)
															 strDict(Y,dictKey)  = strKey
															 strDict(Y,dictItem) = strItem
														 End If
											Next
										  Next
										Schema.RemoveAll
									
										For X = 0 to (Z - 1)
												Schema.Add strDict(X,dictKey), strDict(X,dictItem)
										Next
										erase	strDict
									End If
									
									
								Dim columnCounter , chartype , width
									columnCounter = 1
								Dim laststart ,lastend
									laststart = 1
									lastend = 1
					 
									For Each i In Schema
										
										if  Instr(Cstr(i),"instock") > 0 then
											
											if LCase(NumberinStockisInteger) ="true" then
													chartype = " Integer"
											else
													chartype = " Char"
											end if
										else
											chartype = " Char"
										end if
										''todo this is new keep eye on this
										if Instr(Cstr(i), "price") > 0 then
											chartype = " float"
										end if
										
										

										if widthorcolumndelimeter ="width" then 
												width = Split(Schema.item(i), ",") '' date comes as  Col1,7=outlet
													Dim startpos
													Dim Endpos
													startpos =  width(0)
													endpos = width(1)
														if ( Cint(lastend ) <> Cint(startpos) ) then '' if laststart'' check if laststartpos + 1 = current startpositon, if not, then we need a skip column
															os.writeline "Col" + Cstr(columnCounter)  + "=skip Char width " + Cstr(startpos - Cint(lastend ) )
															columnCounter = columnCounter + 1	
														end if
													os.writeline "Col" + Cstr(columnCounter)  + "="  + Replace(Cstr(i) ,"S_" ,"") + chartype + " width " + Cstr(endpos - startpos )
													columnCounter =  columnCounter + 1	
													laststart = startpos
													lastend = endpos 

										else '' delimeter seperator
												if Cstr(columnCounter) <> Cstr(Schema.item(i)) then  
														Dim skipcount  , c
														skipcount = Schema.item(i) - columnCounter
															 For c = 0 to skipcount -1
																os.writeline "Col" + Cstr(columnCounter) + "=skip Char"
																columnCounter =  columnCounter + 1
															next
												end if

												os.writeline "Col" + Schema.item(i) + "="  + Replace(Cstr(i) ,"S_" ,"") + chartype
												 columnCounter =  columnCounter + 1
										end if
							Next
							
					os.writeline "DecimalSymbol=" + DecimalSymbol
	
			else '' XLS 
				xls = "true"
				
			
					if  IsNumeric(S_partsID) then
						xpartsID = Cint(S_partsID)
					end if
					if  IsNumeric(S_numberInStock) then
						xnumberInStock  = Cint(S_numberInStock)
					end if	
					
					if  IsNumeric(S_Brand) then
						xBrand = Cint(S_Brand)
					end if
					if  IsNumeric(S_Outlet) then
						xOutlet = Cint(S_Outlet)
					end if
					if  IsNumeric(S_LastPurchase) then
						xLastPurchase = Cint(S_LastPurchase)
					end if
					if  IsNumeric(S_LastSale) then
						xLastSale = Cint(S_LastSale)
					end if
					if  IsNumeric(S_PartsCategory) then
						xPartsCategory = Cint(S_PartsCategory)
					end if
					if  IsNumeric(S_Filterby) then
						xFilterby = Cint(S_Filterby)
					end if
					if  IsNumeric(S_AndFilterBy) then
						xAndFilterBy = Cint(S_AndFilterBy)
					end if
					if  IsNumeric(S_AndNotFilterBy) then
						xAndNotFilterBy = Cint(S_AndNotFilterBy)
					end if
					if  IsNumeric(S_subbrands) then
						xsubbrands = Cint(S_subbrands)
					end if
					if  IsNumeric(S_manufacture) then
						xmanufacture = Cint(S_manufacture)
					end if
					if  IsNumeric(S_condition) then
						xcondition  = Cint(S_condition)
					end if
					if  IsNumeric(S_retailprice) then
						xretailprice = Cint(S_retailprice)
					end if
					if  IsNumeric(S_purchaseprice) then
						xpurchaseprice = Cint(S_purchaseprice)
					end if
					if  IsNumeric(S_description) then
						xdescription = Cint(S_description)
					end if
					
			end if
			
		next	 '' LOop files


	End Function


	Function  AddSchemaDic(ByRef Schema )
			if not S_partsID ="" then
				Schema.Add "S_partsid", S_partsID
			end if
	if not S_numberInStock ="" then
	
					if multistock = "true" then '' todo bug with width 
							Dim splitnoinstock , no
							splitnoinstock=split(S_numberInStock, ",") 
									for no = 0 to ubound(splitnoinstock) '' we want to skip the first one
										dim cur
										cur = Cstr(no + 1)
										Schema.Add "S_numberinstock" + cur , splitnoinstock(no)
									next	
					else
						Schema.Add "S_numberinstock", S_numberInStock
					end if

			end if
			if not S_Brand ="" then
				Schema.Add "S_brand", S_Brand
			end if
			if not S_Outlet ="" then
				Schema.Add "S_outlet", S_Outlet
			end if
			if not S_LastPurchase ="" then
				Schema.Add "S_lastpurchase", S_LastPurchase
			end if
			if not S_LastSale ="" then
				Schema.Add "S_lastsale", S_LastSale
			end if
			if not S_PartsCategory ="" then
				Schema.Add "S_partscategory", S_PartsCategory
			end if
			if not S_Filterby ="" then
				Schema.Add "S_Filterby" ,    S_Filterby 	
			end if
			if not S_AndFilterBy ="" then			       
				Schema.Add "S_AndFilterBy"  ,S_AndFilterBy
			end if
			if not S_AndNotFilterBy ="" then							
				Schema.Add "S_AndNotFilterBy" , S_AndNotFilterBy
			end if
	
			if not S_subbrands ="" then
				Schema.Add "S_subbrands", S_subbrands
			end if
			if not S_manufacture ="" then
				Schema.Add "S_manufacture", S_manufacture
			end if
			if not S_condition ="" then
				Schema.Add "S_condition", S_condition
			end if
			if not S_retailprice ="" then
				Schema.Add "S_retailprice" ,    S_retailprice 	
			end if
			if not S_purchaseprice ="" then			       
				Schema.Add "S_purchaseprice"  ,S_purchaseprice
			end if
			if not S_description ="" then							
				Schema.Add "S_description" , S_description
			end if
			
			
			
	End Function
 
 
	Function getCharacterSet()
		Dim myc
		Select Case ReadCharacterSet
			Case 1
			getCharacterSet = "=OEM "
			Case  2
			getCharacterSet = "=ANSI"
			Case 3
			getCharacterSet = "=Unicode"
			Case 4
			getCharacterSet = "=65001"
			Case 5
			getCharacterSet ="UTF-32"
			Case else
			msgbox ("bad  CharacterSet")
		end select
	End Function

	function getFormatdelimited()
		Dim myd
		myD = Lcase(Left(FormatDelimited,1))
		Select Case myd
			Case "f"
			getFormatdelimited = "Fixedlength"
			Case  "t"
		getFormatdelimited = "tabdelimited"
			Case else
			getFormatdelimited = "Delimited(" + myD + ")"
		end select
	end function

	Function CatLength()
		Dim x
			If  Ubound(ObsCategories) > -1 THen 
				x = Ubound(ObsCategories) + 1
			end if
			If  Ubound(ScrapCategories) > -1 THen 
				x = x + Ubound(ScrapCategories) + 1
			end if
			If  Ubound(ExcludeCategories) > -1 THen 
				x = x +  Ubound(ExcludeCategories) + 1
			end if
		CatLengthCombo = x
	end Function

	Function FindMatch(ByRef asource, ByVal strValue)
	   Dim intLB, intUB
	   Dim i
	   Dim intMatch
	   intMatch = -1
	   if not isnull(strValue) and not isnull(asource) then
			intLB = LBound(aSource)
			intUB = UBound(aSource)
			   
			For i = Max(intLB, 0) To intUB
				If CStr(aSource(i)) = CStr(strValue) Then
					intMatch = "0"
					Exit For
				End If
			Next
			FindMatch = Cstr(intMatch)
	   end if
	End Function

	Function Max(ByVal a, ByVal b) 
	  If a > b Then 
		Max = a 
	  Else 
		Max = b 
	  End If 
	End Function

	Sub UseCategoryforDates()

		if  ubound(ObsCategories) > -1 or ubound(ScrapCategories) > -1  or ubound(ExcludeCategories) > -1   then
			UseCategory = true
		else
		UseCategory = false
		end if
		
		
	end sub

	

	Function GetOutLetsOrBrands(Byval outletorbrand)
	

		if outletorbrand="brand"   then '' we need a schema setting for brand for this to work
				
			
			
				if len(brands) > 0 and len(BrandsInPartID) > 0 then
					msgbox ("cant have both brand and brandsinpartsid")
				end if
				
				if len(BrandsInPartID) > 0 then
						GetOutLetsOrBrands = BrandsInPartID
					exit function
				end if
				
			
				if not S_Brand= "" then '' we need to have a value in schema to work
					if brands = "automatic" then
						GetOutLetsOrBrands = Distinct("brand")
					elseif len(brands) > 0 then
						GetOutLetsOrBrands = brands
					else
						GetOutLetsOrBrands ="NOBRANDS"
					end if
			  else
					GetOutLetsOrBrands ="NOBRANDS"  '' we always need at least one value to continue the for loop
			  end if
				
					
		end if
		

		if outletorbrand  ="outlet" Then
			if not S_Outlet = "" then
			
						if outlets = "automatic" then
								GetOutLetsOrBrands = Distinct("outlet")
						elseif outlets = ""  and FileLocationSave ="" then
								GetOutLetsOrBrands = "files"  ' only one dealer for FTP server, we want files in the files folder
						else		
								GetOutLetsOrBrands = outlets  '' they manually type in the outlets
						end if						
		else
				if multistock ="true" then '' we turn numberinstock into Dealers!  TODO, does not work on a width delimiter. only column del
						GetOutLetsOrBrands =  S_numberInStock  '' each line has 2 x numberinstock, which are differant dealers.
				else
						if  FileLocationSave = "" then
							GetOutLetsOrBrands = "files"
						else
							GetOutLetsOrBrands ="results"
						end if	
				end if	
			end if	

			
		end if

	End Function
	
	
	Function Distinct( Byval columnname)

Dim concat


				Set conn1 = createobject("ADODB.Connection")
							Set rs1 = CreateObject("ADODB.recordset")

							if xls ="true" then
								Conn1.Open ("Driver={Microsoft Excel Driver (*.xls)}; Dbq=" & filelocation  &  "; DefaultDir=" & filelocation )
							else
								Conn1.Open ("DBQ=" & filelocation & ";DefaultDir=" & filelocation & ";Driver={Microsoft Text Driver (*.txt; *.csv)};DriverId=27;Extensions=None,asc,csv,tab,txt;FIL=text;")
							end if
									rs1.ActiveConnection = Conn1
									'' cant have numberinstock > 0 because of culture issues
									if multistock ="true" then
										rs1.Open "SELECT DISTINCT " & Cstr(columnname) & " from " &  fileToBesplit  &" where len(numberinstock1) > 1  ;", , 3
									else
											if LCase(NumberinStockisInteger) ="true" then
												    rs1.Open "SELECT DISTINCT " & Cstr(columnname) & " from """ &  fileToBesplit  &""" where numberinstock > 0  ;", , 3
											else
													rs1.Open "SELECT DISTINCT " & Cstr(columnname) & " from """ & fileToBesplit & """ where len(numberinstock) > 1  ;", , 3
											end if
									end if
							
									Do While Not rs1.EOF
									 If IsNull(rs1.Fields(columnname).Value)  Then
									 
									 concat = concat +  "blank,"
									Else
											if not RegexReturn(rs1.Fields(columnname)) ="" then
											concat = concat +  RegexReturn(rs1.Fields(columnname))  + ","	
											end if
									End If
												rs1.MoveNext	
										Loop	
							rs1.Close
							Conn1.Close
							Set rs1 = Nothing
							Set Conn1 = Nothing
							
								if len(concat) > 0 then
										Distinct =  Left(concat, Len(concat) - 1) 
								else
										Distinct = ""
								end if	


							
	end function
	
	
	
Function RegexReturn(ByVal RemoveSpecalChars) '' return only A-Z and numbers

	Dim regEx   ' Create variable.
				Set regEx = New RegExp   
				regEx.Pattern = "[^A-Za-z0-9 ]"   
				regEx.IgnoreCase = True   
				regEx.Global = True   
				
		
				RegexReturn =  RegEx.Replace(RemoveSpecalChars,"") 
	
	End Function
Function RegexReturnInt(ByVal RemoveSpecalChars) '' return only A-Z and numbers



	
	 If not IsNull(RemoveSpecalChars) Then


			Dim regEx   ' Create variable.
						Set regEx = New RegExp   
						regEx.Pattern = "[^0-9 ]"   
						regEx.IgnoreCase = True   
						regEx.Global = True   
						Dim returniNT
						returniNT =  RegEx.Replace(RemoveSpecalChars,"") 
						IF IsNumeric(returniNT) then
							RegexReturnInt= returniNT
						else
							RegexReturnInt = 0
						end if	
	else
	
	RegexReturnInt = 0
	end if
	
End Function	
	
	
	
	

	Function UseLastPurchaseORLastSale(ByVal lastP , ByVal LastS )
				

				
		if usedatelastPurchase = "true"  then 
				 if lastP <> "" and (lastP * 2 > 2) then  '' IF NOT empty and number is not fulled with zeros's
						UseLastPurchaseORLastSale =  lastP
				 elseIf LastS <>"" and (LastS * 2 > 2) then
				   UseLastPurchaseORLastSale =  LastS
				 else
				   UseLastPurchaseORLastSale = "00000000"
				 end if
 
			else
			
				 if lastS <> "" and (lastS * 2 > 2) then 
					UseLastPurchaseORLastSale =  lastS
				 elseIf lastP <>"" and (lastP * 2 > 2) then
				   UseLastPurchaseORLastSale =  lastP
				 else
				   UseLastPurchaseORLastSale = "00000000"
				 end if
			end if
			
			
		UseLastPurchaseORLastSale= 	RemoveXfromStartofDate(UseLastPurchaseORLastSale)
			
		
	end function
	
	
	Function RemoveXfromStartofDate(UseLastPurchaseORLastSale)

			if RemoveXcharsfromDate > 0  and  len(UseLastPurchaseORLastSale) > RemoveXcharsfromDate  then
				   RemoveXfromStartofDate =		RIGHT(UseLastPurchaseORLastSale,LEN(UseLastPurchaseORLastSale)-RemoveXcharsfromDate)
				else
				RemoveXfromStartofDate =  UseLastPurchaseORLastSale
			end if	
	
	End function
	

	Function GetFolderLocation()
			GetFolderLocation =   fs.getfolder(".") & "\"
	End function

	
	
	
	
	
	
	
	
	
	
	
	
	Function GetFileSaveLocation()

		if len(FileLocationSave  ) > 2 then
			if instr(FileLocationSave,"\") > 0 then

				GetFileSaveLocation =  FileLocationSave  
			else
				GetFileSaveLocation =  FileLocationSave & "\"  	
	
			end if
		
				
		else
			GetFileSaveLocation =   fs.getfolder(".") & "\"
		end if
	End function


	Function RemoveEmptyScrapObsCurrentFiles()
	


		if  RemoveEmptyScrapObsCurrent = "true" then
					'' dealer would most times be the File folder for a single dealer
					if  counterCurrent = 0 then
				
						DeleteFileFunction( curFilePath)
					end if
					if  counterScrap = 0 then
		
						DeleteFileFunction(scrapFilePath)
					End if
					if counterObs = 0 then
		
						DeleteFileFunction(obsFilePath)
					end if
		end if
		
		RemoveEmptyScrapObsCurrentFiles=""
		
	End Function

	Function DeleteFileFunction(ByVal Filename)
							Dim aFile
							Set fs = CreateObject("Scripting.FileSystemObject")
							if fs.FileExists(Filename) then  
								Set aFile = fs.GetFile(Filename)
									aFile.Delete
									Set aFile = Nothing
							end if	
	end function

	Function RenameStockFile()
	
	
		if AllowStockFikeToBeAutoRenamed = "true" then
	
			dim  filelistt
			filelistt= split(FileName , ",")
			
				if Ubound(filelistt) < 0 then
	
				DIM fso, objFile , colFiles ,objFolder , filename ,obext 
				Set fso = CreateObject("Scripting.FileSystemObject")
					

					if fso.FileExists(fileToBeSplit) then
						renamestockfile = fileToBeSplit
					else  '' if the schema stockfile name is not on server, then we need to rename it to schema
					 fileextention = fso.GetExtensionName(fileToBeSplit)
					
							Set objFolder = fso.GetFolder(".") 
							Set colFiles = objFolder.Files 
							
							For Each objFile in colFiles 
								filename  = objFile.Name
								Set obext = CreateObject("Scripting.FileSystemObject")
								if LCase(obext.GetExtensionName(objFile.Name))  = LCase(fileextention) then '' IF there is a match btw root stockfile and schema stockfile, then we like
												
										filename = objFile.Name
									   obext.MoveFile  filename , fileToBeSplit
									   
										exit for
									end if
								Set obext = Nothing
							Next 
					end if
				set obext = Nothing
			end if
		end if
	end function

	Function RenameStockFileEnd()
	
	
				if Lcase(renamestockfiletoold) = "true" or Instr(GetFolderLocation , "E:\FTP-DealerUpload\Dealers") > 0  then  '' if live we always want to renane stock files to old
					Dim aFile
					Set fs = CreateObject("Scripting.FileSystemObject")
				
				
					if fs.FileExists(fileToBeSplit & ".old") then  
						Set aFile = fs.GetFile(fileToBeSplit & ".old")
							aFile.Delete
							Set aFile = Nothing
					end if
			
				
					if fs.FileExists(fileToBeSplit) then  
						Set aFile = CreateObject("Scripting.FileSystemObject")
							aFile.MoveFile  fileToBeSplit , fileToBeSplit & ".old"
							Set aFile = Nothing
					end if
					
				
					if fs.FileExists(origonalxlsfilename) then  
						Set aFile = CreateObject("Scripting.FileSystemObject")
						
							if fs.FileExists(origonalxlsfilename & ".old") then
								Dim dFile
								Set dFile = fs.GetFile(origonalxlsfilename & ".old")
									dFile.Delete
									Set dFile = Nothing
							 end if
							 
						aFile.MoveFile  origonalxlsfilename , origonalxlsfilename & ".old"
						Set aFile = Nothing
					end if
					

			end if
	End Function

	Function stockfilename()
	
				
	if FileName = "false" or Filename = "autoread" then  '' we want to autoread, so we dont know 
	'' we need to do this to make a dynamic list of file names, we rename to short 1,2,3,4 ect...
	
			  Dim fso ,folder, files, NewsFile,sFolder , i  , concatfilename , fileext
			  Set fso = CreateObject("Scripting.FileSystemObject")
			  Set folder = fso.GetFolder(GetFileSaveLocation)
			  Set files = folder.Files
	
			  
			  i = 1
			  
			  concatfilename = ""
	
			  For each folder In files
					if fso.GetExtensionName(folder.Name) = "csv" or fso.GetExtensionName(folder.Name) = "txt" then
						fileext=  "." + fso.GetExtensionName(folder.Name)
						fso.MoveFile folder.Name, Cstr(i) + fileext
						concatfilename = concatfilename +  Cstr (i) + fileext +  ","
						i = i + 1
					end if
			  Next
			  
			  if  Len(concatfilename) > 1 then
			  stockfilename = Left(concatfilename,Len(concatfilename) -1)
			  else
			  stockfilename = ""
			  end if
			  
			 
			 Filename =  stockfilename

			   multifilekillrootfolder ="true"
			 WScript.Sleep(1000)



			Set fso = Nothing
			Set folder = Nothing
			Set files = Nothing
	else
		stockfilename = LCase(FileName)
		
	end if
	
	end function
	
	

	
	
		Function PartFarmSubbrandDiscounts(ByVal lngLastMovement)
		'' 0 is empty  so the stock file is giving no date ''1 is invalid data, stock file gives as date, but it is a bad format
		'SettingsForEmptyDate = current     '' use current, scrap or obs   
			'SettingsForInvalidDate =  current '' use current, scrap or obs		
			
		
		
			If  lngLastMovement = 0 then   ''0 is empty
	
				Select Case SettingsForEmptyDate
					Case "obs"
						PartFarmSubbrandDiscounts = DiscountSubbrandOlderThanX
					Case "scrap"
						PartFarmSubbrandDiscounts = DiscountSubbrandOlderThanX
					Case Else
						PartFarmSubbrandDiscounts	= DiscountSubbrand 
				end select
			elseif lngLastMovement = 1 then    ''1 is invalid
			

				Select Case SettingsForInvalidDate
					Case "obs"
						PartFarmSubbrandDiscounts	= DiscountSubbrandOlderThanX
							Case "scrap"
						PartFarmSubbrandDiscounts = DiscountSubbrandOlderThanX
					Case Else
						PartFarmSubbrandDiscounts	= DiscountSubbrand 
				end select
				
			else
					if ((today- lngLastMovement) > DiscountPriceXDays) then 
							PartFarmSubbrandDiscounts	= DiscountSubbrandOlderThanX 
					else 
							PartFarmSubbrandDiscounts	= DiscountSubbrand   
					end if						
			end if
	End Function
	
	

	Function CheckisObsolete(ByVal lngLastMovement)
		'' 0 is empty  so the stock file is giving no date ''1 is invalid data, stock file gives as date, but it is a bad format
		'SettingsForEmptyDate = current     '' use current, scrap or obs   
			'SettingsForInvalidDate =  current '' use current, scrap or obs		
			
		
			If  lngLastMovement = 0 then   ''0 is empty
				Select Case SettingsForEmptyDate
					Case "obs"
						CheckisObsolete = "true"
					Case Else
						CheckisObsolete = "false"
				end select
			elseif lngLastMovement = 1 then    ''1 is invalid
				Select Case SettingsForInvalidDate
					Case "obs"
						CheckisObsolete = "true"
					Case Else
						CheckisObsolete = "false"
				end select
			end if
	End Function

	Function CheckisScrap(ByVal lngLastMovement)
		'' 0 is empty  so the stock file is giving no date ''1 is invalid data, stock file gives as date, but it is a bad format
		'SettingsForEmptyDate = current     '' use current, scrap or obs   
		'SettingsForInvalidDate =  current '' use current, scrap or obs		
	

		If  lngLastMovement = 0 then   ''0 is empty
				Select Case SettingsForEmptyDate
					Case "scrap"
						CheckisScrap = "true"
					Case Else
						CheckisScrap = "false"
				end select
			elseif lngLastMovement = 1 then    ''1 is invalid
				Select Case SettingsForInvalidDate
					Case "scrap"
						CheckisScrap = "true"
					Case Else
						CheckisScrap = "false"
				end select
			end if
	End Function

	function CheckDoubleRows()
	
	if  LCase(doublerows) = "true" then
			Set objDictDoubleRowsRemoval = CreateObject("Scripting.Dictionary") 
	end if

	end function
	

	Function RegeXRun()
'' not used in xsl untill after file conversion

		if dbf = "true" then
			exit function
		end if
	
			Set regEx = New RegExp  
			regEx.Pattern = RegPattern    '''regex.Pattern = Chr(34)  remove quotes
		
			dim fss, ts1, ts2, str, stringWriter, i, count  , objDict , GetFileExtention
			files = split (FileName, ",")
			Set fss = CreateObject("Scripting.FileSystemObject")
			Filelocation =  fs.getfolder(".") & "\"
			count=0		

					for each file in files 
				
						if fss.FileExists(filelocation & file ) and instr(file,".xl") = 0  then
						
					
							fss.createTextFile filelocation & "temp.tmp", True
								Set ts1 = fss.openTextfile(filelocation & file, 1)
								Set ts2 = fss.OpenTextFile(filelocation & "temp.tmp", 8)
								str = ""
								do while not ts1.atEndofStream
									str = ts1.readLine
									if	not removeColumnXfromTextFile =  "0" and not removeColumnXfromTextFile= "" then
								
										
										dim arr	,  newcleanerarray  , arr_ofcolumns , ColumnToRemove
											arr = Split(str,FormatDelimited) ' this is the whole line in txt file, slit by delimiter
											arr_ofcolumns = Split(removeColumnXfromTextFile,",")'' remove all the colums we desire!
											
											 for ColumnToRemove = 0 to  ubound(arr_ofcolumns)
												
													if ubound(arr) > ColumnToRemove  then
															newcleanerarray =	ArrayRemoveAt(arr,Cint(arr_ofcolumns(ColumnToRemove)) -1 )
															Dim newstr 
																newstr =join(newcleanerarray,FormatDelimited)		
													end if
													
											 next
												
	
											str = newstr
											
										
									end if
										
								
				
									if doublerows ="true" then  ' ideally we should never get doublerows, sometimes it just cant be avoided.
											 if  not objDictDoubleRowsRemoval.Exists(str) then 
												objDictDoubleRowsRemoval.add str , str
												str = regex.replace(str,Replacewith)
												str =  Replace(str,"\""","")
												ts2.writeline str
											end if
									else
											str = regex.replace(str,Replacewith)
											str =  Replace(str,"\""","")
											ts2.writeline str	
									end if
							
								Loop
								ts1.close
								ts2.close
								
							fss.CopyFile filelocation & "temp.tmp",filelocation & RenameGarminUNLifneeded(file),true
							fss.DeleteFile filelocation & "temp.tmp"
								count = count + 1
					
						end if
					next
	end function
	

	
Function  ArrayRemoveAt(arr, pos)
  Dim i
  If IsArray(arr) Then
    If pos >= 0 And pos <= UBound(arr) Then
      For i = pos To UBound(arr) - 1
        arr(i) = arr(i + 1)
      Next
      ReDim Preserve arr(UBound(arr) - 1)
	  
	  ArrayRemoveAt = arr
    End If
  End If
End Function 


		Function RenameGarminUNLifneeded(file)

		dim GetFileExtention , fsss , inputfile
		inputfile = file
		
		Set fsss = CreateObject("Scripting.FileSystemObject")
		
				GetFileExtention =  fsss.GetExtensionName(file)
					if  GetFileExtention ="unl" then
							file =  Replace(file,"unl" , "txt")
							FileName=  file   ' sets new Filename at the top
							RenameGarminUNLifneeded = file '' returns new filenmae
							
							fsss.DeleteFile filelocation & inputfile
							else
							RenameGarminUNLifneeded = file
					end if
						

		end function
		
		
		Function RegeXRunFile(file)
		
		
	  ' This function is used on xml/xls conversion by default every time!
		if dbf = "true" then
			exit function
		end if

			Set regEx = New RegExp  
			regex.Pattern = RegPattern    '''regex.Pattern = Chr(34)  remove quotes
			dim fss, ts1, ts2, str, stringWriter, i, count
			Set fss = CreateObject("Scripting.FileSystemObject")
			Filelocation =  fs.getfolder(".") & "\"
			count=0

						if fss.FileExists(filelocation & file ) and instr(file,".xl") = 0 then
							fss.createTextFile filelocation & "temp.txt", True
				
							Set ts1 = fss.openTextfile(filelocation & file, 1 ,1)
							Set ts2 = fss.OpenTextFile(filelocation & "temp.txt", 8 ,1)
					
							str = ""
							do while not ts1.atEndofStream
								str = ts1.readLine
								str = regex.replace(str,Replacewith)
								str =  Replace(str,"\""","")
							
								ts2.writeline str
							Loop
							ts1.close
							ts2.close	 
						fss.CopyFile filelocation & "temp.txt",filelocation & file,true
						fss.DeleteFile filelocation & "temp.txt"
							count = count + 1					
						end if

			if   len(MergeMultiLine) > 0 then
				Mergelines(file)
			end if

				
								
	end function
	
	
	
		
	Function ConvertPDFToText()

			dim fss ,pdfsoftwarelocation
			Set fss = CreateObject("Scripting.FileSystemObject")
				Filelocation =  fs.getfolder(".") & "\"
					
				if fss.FileExists(filelocation & FileName ) and instr(FileName,".pdf") > 1 then  '' If the pdf exsists then convert it
						
						pdfsoftwarelocation =""
						if fss.FileExists("E:\UploadTool\pdftotext.exe")  then
							pdfsoftwarelocation =  "E:\UploadTool\pdftotext.exe"
						end if
						
						if fss.FileExists("D:\DEV\DLL\PDF\pdftotext.exe")  then
							pdfsoftwarelocation = "D:\DEV\DLL\PDF\pdftotext.exe"
						end if
		
						if fss.FileExists(pdfsoftwarelocation)  then  '' if pdftotext.ext does not exsist then alert us

									dim WshShell
									Set WshShell = WScript.CreateObject("WScript.Shell") 
									WshShell.Run pdfsoftwarelocation + " -layout stock.pdf stock.txt" 
									WScript.Sleep(500)
									
									fss.CopyFile filelocation & FileName, filelocation & FileName  & ".old", true
									fss.DeleteFile filelocation & FileName
									
									FileName = "stock.txt"
								
						else
								msgbox("need the pdftotext software in E:\UploadTool\ or  D:\DEV\DLL\PDF\pdftotext.exe")
						end if
				
			
				end if

				

				
			

	End Function
	
	
	

function Mergelines(file) 
'' if file has value then this is from teh xls conversion.

'' todo still need to work on a txt and real csv file and array of files
		dim fs, ts1, ts2, str, strsplit, getline1, getline2 ,getValue
		
			Set fs = CreateObject("Scripting.FileSystemObject")
			Filelocation =  fs.getfolder(".") & "\"

			fs.createTextFile filelocation & "temp.tmp", True
			Set ts1 = fs.openTextfile(filelocation & file,  1)
			Set ts2 = fs.OpenTextFile(filelocation & "temp.tmp", 8)

		
			do while not ts1.atEndofStream
					  getline1 = ts1.readLine
						getValue = split(getline1, ",")
						
						If IsNumeric(getValue(MergeMultiLine)) Then

										str= getline1 
										
										
											if  Not ts1.AtEndOfStream then
												getline2 =  ts1.readLine
												
												str = str &  getline2
											end if
											
											Ts2.writeLine str
											str=""
						   else
								str=""
							end if

			Loop
		ts1.close
		ts2.close
		fs.CopyFile filelocation & "temp.tmp",filelocation & file,true
		fs.DeleteFile filelocation & "temp.tmp"
	end function
	
	
	function CheckFilterByClauses()

	
			'' to use filterby, we need this
			''Const FilterBy =""			''Leave empty to ignore. To use this, we need a "Col2=FilterBy char width 15" ect.. to work
			''Const AndFilterBy=""			''We can have  where FilterBy='OES' and AndFilterBy='OEM' "
			''Const AndNotFilterBy=""	
		Dim buildFilter
			
			if len(FilterByMatching) > 0 then
				
			
					if inStr(FilterByMatching ,",") then
						buildFilter = BuildSQLSplit(FilterByMatching)
					
					else
						buildFilter = "and FilterBy='" & FilterByMatching & "'"		
					end if
					

			 end if
				
			if len(AndFilterByMatching) > 0 then
					if inStr(AndFilterByMatching ,",") then
						buildFilter =  buildFilter & BuildSQLSplit(AndFilterByMatching)
					else
						buildFilter =  buildFilter & "and FilterBy='" & AndFilterByMatching & "'"	
					end if
			 end if

			if len(AndNotFilterBy) > 0 then
				buildFilter =  buildFilter &  "and Not AndNotFilterBy='" & AndNotFilterBy & "'"							
			 end if

			 
			CheckFilterByClauses =  buildFilter
			
	end function
	
	
	function BuildSQLSplit(FilterBy)

	Dim  splitFilter, spf , concatstring
		 splitFilter = split(FilterBy, ",")
		 for spf = 1 to  ubound(splitFilter)
		 concatstring =  concatstring + "or FilterBy='" & splitFilter(spf) & "'" 
		 next
	
		BuildSQLSplit=  "and FilterBy='" + splitFilter(0) & "'"  & concatstring
	end function
	

	
	function CheckBrands(j)

			'' CHANGED  14th jan 2013 added nobrands
			if Ubound(brand) = 0 and j ="NOBRANDS"  then
				CheckBrands =""			
			else
		
					'' this is for brandsinsidethepartid NOT FOR DBF FILES YET
					if len(BrandsInPartID) > 0 then
							if j = "leftover" then	
					Dim lo, leftover , splitbrandsin

					
						splitbrandsin =split(BrandsInPartID, ",")
								for  lo = 0 to ubound(splitbrandsin) 
									leftover =  leftover &  " and InStr(1,partsid,'" & splitbrandsin(lo)  &  "')= 0 "
								next
					
									CheckBrands = leftover

							else
							CheckBrands =  " and InStr(1,partsid,'" & j & "')= 1"
							end if
						
					else
							if BrandWildcard = "true" then  '' this will catch  j* 
								if dbf ="true" then
									CheckBrands =  " and brand like'" & j  &   "%'"   
								else
									CheckBrands =  " and brand like'" & j  &   "%'"   
								end if
						else
								if dbf ="true" then
									CheckBrands =  " and brand='" & j  &   "'"   
								else
									CheckBrands =  " and brand='" & j  &   "'"   
								end if
				
						end if

					end if
			
		
			end if
			
			
	end function
	
	
		function CheckOutLets(i)
		
			if Ubound(dealer) = 0 then
				if i ="files" or i= FileLocationSave then
					CheckOutLets =""
				else
					CheckOutLets = " and outlet='" & i  &   "'"
				end if
			elseIF i  = "blank"  then '' we concider empty outlets as a dealer too.
	        		CheckOutLets= " AND outlet IS null "
			else
				CheckOutLets= " and outlet='" & i  &   "'"
			end if 
			
			
	end function


	function GetNumberinStockSQL()
		if LCase(NumberinStockisInteger) ="true" then
		
			if dbf ="true" then
				GetNumberinStockSQL= " where " & S_numberInStock & " > 0 "
			else
			
	
				GetNumberinStockSQL = " where numberinstock >=" & MaxStock 
				
			
			end if
			
		else
			GetNumberinStockSQL =" where 1= 1 "
		end if
	end function
	
	function CheckPartIDSize()
		if mincharactorallowedpartid > 0 AND maxcharactorallowedpartid > 0  AND ( maxcharactorallowedpartid >  mincharactorallowedpartid )then
				if dbf ="true" then
					getsqlParttIDmaxminmaxlengh = " and len(Trim("& S_partsID &")) >= "  & mincharactorallowedpartid  & "  and len(Trim("& S_partsID &")) <= " & maxcharactorallowedpartid         
				else
					getsqlParttIDmaxminmaxlengh = " and len(Trim(PartsID)) >= "  & mincharactorallowedpartid  & "  and len(Trim(PartsID)) <= " & maxcharactorallowedpartid         
				end if
			
			CheckPartIDSize = "true"
		else
		
		if dbf ="true" then
				getsqlParttIDmaxminmaxlengh = "and len("& S_partsID &") > 1"
				else
				getsqlParttIDmaxminmaxlengh = "and len(PartsID) > 1"
				end if
			
			CheckPartIDSize = "false"
		end if
    end function
	
	
	Function CheckMinPrice()
	

		if minprice  > 0 and len(S_retailprice) > 0 then
			if dbf = "true" then
				CheckMinPrice= " and " & S_retailprice & "  >  " & minprice 
			else
				CheckMinPrice= " and retailprice  >  " & minprice 
			end if

			else
			     CheckMinPrice =""
			end if 
	End Function
	
	Function FileterSubBrands()
		if len(subbrandfilter) > 0 then
			Dim loo, buildsubbrands ,splitsubbrandsin
			splitsubbrandsin =split(subbrandfilter, ",")
			buildsubbrands = " and ( "
									for  loo = 0 to ubound(splitsubbrandsin) 
										if loo > 0 then
											buildsubbrands =  buildsubbrands &  " or subbrands='"&splitsubbrandsin(loo)& "'"
										else
											buildsubbrands =  buildsubbrands &  " subbrands='"&splitsubbrandsin(loo)& "'"
										end if
									next
			buildsubbrands = buildsubbrands & " ) "
			FileterSubBrands = buildsubbrands
		end if
	end function
	
	

function PartsID(Byval parts)

		'' NULL VALUES we ignore or give error, settings in config
		 If IsNull(parts) Then
				if xls ="true" then
					if LCase(XLSBadDataIgnore) = "true" then
						PartsID =""
					else
						msgbox("xls format for parts is bad, maybe set ConvertXLSToTXTTAB to true")
					end if
					exit function
				else
					PartsID =""
					exit function
				end if
		 end if
	
	
	Dim temppart
	temppart =  Trim(parts)
		'' first one is part of the filter
		if len(PartidMatchStartingWith) > 0 then
				if not left(temppart,3) = PartidMatchStartingWith then
					temppart = ""
				end if
		end if
			
		'' 1st  RemoveXfromstart  PartID format	
			
		if RemoveXfromstart > 0  and  len(temppart) > RemoveXfromstart  then
				temppart =		RIGHT(temppart,LEN(temppart)-RemoveXfromstart)
				
		end if		
				
			
			
		'' 2nd PartID format. remove chars from beg of string, right to start
		if len(FindStartingCharandRemoveAllBefore) > 0 Then
			Dim positon
				positon = Instr(temppart , FindStartingCharandRemoveAllBefore )
				if positon > 1 then
					temppart =	right(temppart , len(temppart) - positon)
				end if
		end if
		
		
		'' 2.1nd PartID format. remove chars from beg of string, right to start
		if len(FindStartingCharandRemoveAllAfter) > 0 Then
			Dim positon21
				positon21 = Instr(temppart , FindStartingCharandRemoveAllAfter )
				if positon21 > 1 then
			
					temppart =	left(temppart ,  positon21 )
				end if
		end if
		
		
	'' 3rd PartID format.  remove WHOLE partid if founds match at start, we use a comma delimiter to add more than one item to remove
		if len(FindStartingCharsandRemovePartID) > 0 Then
			
			Dim positon1 , splitstartingchars, removethis
			splitstartingchars = split(FindStartingCharsandRemovePartID, ",")
			
			for each  removethis in  splitstartingchars
				
				positon1 = Instr(temppart , removethis )
				if positon1 = 1 then
					temppart =	""
				end if
			next
			
		
		end if
		
	''''  Raleigh Code 23/12/2013
        '''' 4th rd PartID format. remove all specified chars from beg of string, left to right if it matches FindStartingCharandRemoveSpecifiedChars
		
		if len(FindStartingCharandRemoveSpecifiedChars) > 0 Then
			If left(temppart,len(FindStartingCharandRemoveSpecifiedChars) )= FindStartingCharandRemoveSpecifiedChars then
				temppart = right(temppart, len(temppart) - len(FindStartingCharandRemoveSpecifiedChars))
			end if
		end if
		
		
		
	  ''5th PartID format.  remove WHOLE partid if founds match at start
		if len(FindFirstInstanceofAndRemoveit) > 0 Then
			Dim positon2
				positon2 = Instr(temppart , FindFirstInstanceofAndRemoveit )
		if  positon2 > 0 then
				Dim sub_str, sub_str2
						sub_str = Mid(temppart, 1, positon2 - 1) '' get first half before the Replace value
				        sub_str2 = Mid (temppart,positon2 + 1, len(temppart) - positon2  ) '' get 2nd half after the Replace value
						temppart = sub_str + sub_str2 '' join first half and second half. there is no real substring option in vbs.
					end if			
		end if


		 ''6th
		if len(FindsubsequentcharsBeginingatPostionX) > 0 Then
				Dim EndInt, StartEndInt, cc
		
				if  Instr(temppart , LookingFor ) =   StartPostionAT + 1  then '' find out if starting postion matches our Char
							for cc = 1 to 10
								if Instr( StartPostionAT  + cc , temppart , LookingFor ) =   StartPostionAT  + cc then
									EndInt =  Instr( StartPostionAT  + cc , temppart , LookingFor )
								else
									exit for
								end if	
							next
							
				Dim sub_mat, sub_mat2
								sub_mat = Mid(temppart, 1, Instr(temppart , LookingFor ) - 1) '' get first half before the Replace value
								sub_mat2 = Mid (temppart,EndInt + 1, len(temppart) - Instr(temppart , LookingFor )  ) '' get 2nd half after the Replace value
								temppart = sub_mat + sub_mat2 '' join first half and second half. there is no real substring option in vbs.


				else
						EndInt = 0
				end if
		
		
		
		end if 
	
	
				
		''7th and last. match and replace
	if matcharray ="true" then	
		if UBound(MatchAndReplace) => 0 then
			dim m,n
			for m = 0 to UBound(MatchAndReplace)
				temppart =Replace(temppart,MatchAndReplace(m),"")
		    next
		End if
	else
	
		'' match and replace
		if len(matchandreplace) > 0 then
			temppart =Replace(temppart,matchandreplace,"")
		End if
	
	end if
	
		
		
		'''' 8th rd PartID format. remove x chars from beg of string, left to right
		if len(replacestartchar) > 0 Then
			If left(temppart,replacestartcharcount)= replacestartchar then
				temppart = right(temppart, len(temppart) - replacestartcharcount )
			end if
		end if



''9th RemoveLeadingZerosWhenPartIDisAllNumbers
	if Lcase(RemoveLeadingZerosWhenPartIDisAllNumbers) = "true" then
			If IsNumeric(temppart) Then
			
					
					
				Dim regEx   ' Create variable.
				Set regEx = New RegExp   ' Create a regular expression.
				regEx.Pattern = "^0*"   ' \D = not a numeric digit.
				regEx.IgnoreCase = True   ' Set case insensitivity.
				regEx.Global = True   ' Search entire string
			
			
				temppart =  RegEx.Replace(temppart,"") 
					
				
			end if
	
	end if
	


		if len(InsertintoStartofPartID) > 0 then
			temppart =	 InsertintoStartofPartID +	temppart
		end if

		if len(InsertintoEndofPartID) > 0 then
				temppart =	 temppart + InsertintoEndofPartID
		end if


		

		PartsID= temppart

	end function


function DateMonthSwopTextforInteger(Byval ChangeMonthtoInt)


Dim months , m , mo

months = Split(DateisText, ",")

			for  m = 0 to ubound(months)
				
				if 	Instr(ChangeMonthtoInt ,months(m))then
			
			if m +1  <10 then
				mo = "0" & ( m + 1)
				else
				mo =  m +1 
			end if
			
			
					
					DateMonthSwopTextforInteger =	Replace(ChangeMonthtoInt,months(m),mo)
					
					
					exit for
				end if
				
				
				
			next


end function

	
	

	
	
	function RegExReturnOnlyIntegers(ByVal stringtoclean)
	

	
		if len(DateisText) > 10 then
			stringtoclean = DateMonthSwopTextforInteger(stringtoclean)
		end if
		
		'' first we remove the ss:mm:hh
		if len(stringtoclean)>10 then '' this is for hour, min and seconds 00:00:00  '' or we could do it, get first pos of : , remove two to the left, and two to the right, if another : exsists, the remove that aswell. 
			Dim postofblackspace
			postofblackspace =  len(stringtoclean) - InStr(stringtoclean, " ") 
			stringtoclean=left(stringtoclean, len(stringtoclean) - postofblackspace )
		end if
	
	
	
	
		if  not  splitdateby ="" then '' we can split by char or normal ddmmyyy option
				Dim splitd , oArrayList, e , d,m,y , dd,mm,yy , newdate
					splitd = split(stringtoclean, splitdateby)
						d=  InStr(dateformat,"d") 
						m = InStr(dateformat,"m") 
						y = InStr(dateformat,"y") 
						Dim concatdate
						for  e = 0 to ubound(splitd)
								if len(splitd(e)) = 1 then
									concatdate = concatdate +  "0" + splitd(e)
									else
									concatdate = concatdate +  splitd(e)
							end if
						next		
						stringtoclean = concatdate				
		end if
	
	
				RegExReturnOnlyIntegers=trim(stringtoclean) 
				Dim regEx   ' Create variable.
				Set regEx = New RegExp   ' Create a regular expression.
				regEx.Pattern = "\D"   ' \D = not a numeric digit.
				regEx.IgnoreCase = True   ' Set case insensitivity.
				regEx.Global = True   ' Search entire string
			
				Dim returniNT
				returniNT =  RegEx.Replace(RegExReturnOnlyIntegers,"") 
				
			
				IF returniNT = "" then
					RegExReturnOnlyIntegers = 00000000
				else
					RegExReturnOnlyIntegers= returniNT
				end if	
				
				
				

	end function
	function DateFormatClean(ByVal dateformat)
	

			DateFormatClean=trim(dateformat) 
			DateFormatClean =Replace(DateFormatClean,"-","")
			DateFormatClean =Replace(DateFormatClean,"/","")
			DateFormatClean =Replace(DateFormatClean,",","")
			DateFormatClean =Replace(DateFormatClean,"\","")
			DateFormatClean =Replace(DateFormatClean,".","")
			DateFormatClean =Replace(DateFormatClean,":","")
			DateFormatClean =Replace(DateFormatClean,"!","")
	end function

	function convertdate(byval datestring, Byval partid)
	
	


		if IsNumeric(datestring) then
		
		
		'' bug!  if m/yy then it will never be > 1000
			if datestring * 1  => 100 then
				
				dim year, month, day
				dim dformat ,mformat ,yformat  ''  format of date given one day or two d d/mm/yyyy" or dd/mm/yy"
				dim dpos ,mpos ,ypos  ''  format of date given one day or two d d/mm/yyyy" or dd/mm/yy"

				''''''''''''''''''''''''''''''''''''''''''''' READ DATE FORMATS 
					if InStr(dateformat,"dd") > 0 then
						dformat = 2
						dpos = InStr(dateformat,"dd")
					else
						dformat = 1
						dpos = InStr(dateformat,"d")
					end if

					
					
					if InStr(dateformat,"mm") > 0 then
						mformat = 2
						mpos = InStr(dateformat,"mm")
					else
						mformat = 1
						mpos = InStr(dateformat,"m")
					end if

					if InStr(dateformat,"yyyy") > 0 then
						yformat = 4
						ypos = InStr(dateformat,"yyyy")
					else
						yformat = 2
						ypos = InStr(dateformat,"yy")
					end if
					
				
				
			if dpos  + mformat + yformat  > 4  then '' looking to exclude small date format single m/yy
						if datestring * 1  < 1000 then
						'' bug 94. needs more checking	 . we learnt that mm/yy can be smaler than 1000. so we need to only run this check if using day and month and year
							convertdate = 1
							exit function
						end if
			end if
			
		''''''''''''''''''''''''' GET YEAR ''''''''''''''''''''''''
			
				if dformat = 1 and mformat = 1 then
						msgbox("BadDateForamt, day and month cant both be singles , one needs to be a dd or d.")
				end if

	
	'' get year first and crop if year is at end
				if ypos > dpos and ypos > mpos then '' the year is at the end
					if yformat = 4 then
						year  =  CStr(right(datestring,yformat))
					
					else
				
						if CStr(right(datestring,yformat)) > 50 then
							year  =  1900 +CStr(right(datestring,yformat))
						else
							year  =  2000 +CStr(right(datestring,yformat))
						end if
					
				
		
				end if

				datestring =CStr(left(datestring, len(datestring) - yformat)) '' we want to crop the date out of this
				
				
				else '' year in start
				
				if yformat = 4 then
					year  =  CStr(left(datestring,yformat))
				else
			
							if CStr(left(datestring,yformat)) > 50 then
								year  =  1900 +CStr(left(datestring,yformat))
							else
								year  =  2000 +CStr(left(datestring,yformat))
							end if
				end if


				datestring = CStr(Right(datestring, len(datestring)  - yformat)) 
				end if
			



			'''' NEXT WE GET MONTH OR DAY , WHAT IS LEFT OVER WILL BE THE OTHER
				if mformat = 2 then '  GET MONTH
				
		

						if mpos < dpos then '' MONTH IN BEGIN
							month = CStr (left(datestring, 2))
							datestring =Cstr(Right(datestring, len(month) )) '' changed from left to right
							day =CStr(left(datestring,dformat))	
							

						else '' MONTH IN END
					
							month = CStr (right(datestring, 2))
							datestring = Cstr(Left(datestring, len(month) )) '' changed from right to left
							day =CStr(right(datestring,dformat))	 '' last change this for 00000000 dates
			
						end if
						
				elseif dformat = 2 then   '' or GET DAY
			
	
						if dpos < mpos then ''DAY IN BEGININNG
							day = CStr (left(datestring, 2))
							datestring =Cstr(Right( datestring,len(datestring) - len(day) ))'DONT USE REPLACE!
							month =CStr(left(datestring,dformat))	

						else '' DAY AT END
							day = CStr ( right(datestring, 2))
						    datestring =Cstr(left(datestring, len(datestring) - len(day) ))
							month =CStr(right(datestring,dformat))
						end if
						
				end if
				
				
				
				on error resume next
					 convertdate =  year * 365 + month * 30 + day
						if err.number = 0 then
							convertdate =  year * 365 + month * 30 + day
						else
							convertdate =  1
								if showerror="true" then
									msgbox ("day-"& day & "month-"  &  month & "year-"  & year &  "&PARTID( " & partid & " )")
								end if
						end if

					
		
			else

			convertdate = 1
			end if
		else
			convertdate = 1 '' we set to 1 because there is a date value, it is invaild, but there is a value
		end if
	end function
