<%

' - Constants ------------------------------------------------
	' For disabling password security set wexPassword = "" 
	Const wexPassword = ""
	' Root folder, it can be a physical or virtual folder: "c:\test", "/test"
	' Beware that you will not have web access(no web browsing of files/folders) if you use a physical folder like "c:\test"
	' If you want to have web access with a physical folder, you should create a virtual folder(IIS alias like "/folder") for that physical folder and use it instead.
	Const wexRoot = "/UserFiles"
	' Show files and folders that have hidden attribute set?
	Const showHiddenItems = true
	' Calculate total size of the current folder? Disable if it takes long time with huge folders.
	Const calculateTotalSize = true
	' Calculate total sizes of the folders in the listing? Disable if it takes long time with huge folders.
	Const calculateFolderSize = true
	' List of file extensions that can be edited by clicking the icon.
	Const editableExtensions = "htm,html,asp,aspx,asa,asax,txt,inc,css,aspx,js,vbs,shtm,shtml,xml,xsl,log,bas,bat,c,cfg,cpp,css,csv,cxx,diz,doc,h,inf,ini,nfo,php,reg,rtf"
	' List of file extensions that can be viewed by clicking the icon.
	Const viewableExtensions = "gif,jpg,jpeg,png,bmp,jpe,avi,mov,mpeg,mpg,mpe,wmv"
	' List of file other extensions 
	Const otherExtensions = "386,ace,ade,adn,adp,aif,aifc,aiff,ani,arc,arj,asf,asx,au,audiocd,bin,bsc,c2d,cab,cat,ccd,cda,cdi,cdx,cer,chm,cif,cls,cmd,cnf,cpl,ctl,cue,cur,db,dbp,dib,dif,dll,dox,drv,dsk,dsr,dwg,dxg,emf,enc,exe,fcd,fdf,fla,flask,fon,frm,gcd,gz,hlp,ht,htt,htw,htx,hxx,icm,ico,img,imz,iso,jar,jfif,jse,key,lha,lzh,m1v,m3u,mdb,mdl,midi,mod,mp2,mp2v,mp3,mpa,mpe,msc,msi.msp.ncd,nrg,ocx,pag,pak,pcx,pcm,pdf,pdx,pfm,pif,pkp,pps,ppt,psd,ptl,qdat,qds,qpx,qt,qtl,qtp,qts,ram,rwr,raw,rc,rct,rdp,rmf,rmi,scf,scr,sct,sd2,sdb,shb,shs,sln,snd,swf,sys,tar,tgz,tif,tiff,unk,url,vaf,val,vap,vbe,vbg,vbp,vbz,vcd,vip,vup,vxd,wab,wav,wax,wm,wma,wmf,wmp,wmv,wmx,wri,wsc,wsf,wsh,wvx,xls,xnk,xsl,zip"	
	' Display full physical path of the current folder? Set it to false for hiding your server path structure.
	Const displayPath = true
	' List of file extensions to be monitored during Upload, Rename, Editor Save and Create New File actions.
	' Allowance or denial is due to the value of denyMonitored constant. Extensions should be separated by comma.
	Const monitoredExtensions = ""
	' True means monitored extensions are denied and false means only monitored extensions are allowed but not any other.
	Const denyMonitored = true
	
	Const Language="Greek.txt"
	Const showZipUnzip	 = true
' ------------------------------------------------------------

' - Variables ------------------------------------------------
	' Set script timeout value to higher values (in seconds) if the script fails when uploading large files
	Server.ScriptTimeout = 300
	' Preferred character set, default value is "ISO-8859-1" (Western European character set)
	' Don't bother to change it unless you are having problems with handling text in your language
	' Character sets and supporting code pages can be found at
	' http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/charsets/charset4.asp
	Response.CharSet = "windows-1253"
	' Preferred code page, default value is 1252 (Western European codepage)
	' Don't bother to change it unless you are having problems with handling text in your language
	Session.CodePage = 1253
	' Preferred locale identifier, default value is 1033 (English - United States)
	' Locale ID (LCID) Chart can be found at
	' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/script56/html/vsmscLCID.asp
	Session.LCID = 1033
' ------------------------------------------------------------

' - Plugins --------------------------------------------------
	' File transfer plugin
	%><!-- #include file="./plugins/file transfer/WexGeneric.asp" --><%
' ------------------------------------------------------------

%>