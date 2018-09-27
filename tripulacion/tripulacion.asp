<% @ LCID = 1034 %>
<!--#include file="../modulos/vali_sesion.asp" -->
<!--#include file="../modulos/funcion.asp" -->
<!--#include file="../modulos/control.asp" -->
<!--#include file="../modulos/loader.asp"-->
<!--#include file="../js/funcion.js" -->
<!--#include file="../js/autosuggest.js" -->
<!--#include file="../js/datetimepicker_css.js" -->
<%	


	
	
	
	Call lista_arch_js("siep\js\list_autogest_pais.js", "pais", "Nom_pais", "Siep_pais", "")
	Call lista_arch_js2("siep\js\list_autogest_pais_cod.js", "pais_cod", "Cod_pais", "Nom_pais", "Siep_pais", "")
	'On Error Resume Next
	'Se tomó la decisión de utilizar un script para tripulación y otro para pasajero'
	'debido a su complejidad'
	Response.Buffer = true
	Dim load
	Set load = nothing
	Set load = new Loader
	
	load.initialize

	cortesia_bd =  selectSimpleSQL2("Cor_susc", "Siep_susc", "Num_clie = " & Session("Num_clie"), 1)
	if cortesia_bd = true then
		cortesia = " (Suscripci&oacute;n de cortes&iacute;a)"
	else
		cortesia = ""
	End if 
	
	
	
	
	
	
	fuente = "<font color = '#00ade5'><b>Hola. Bienvenido " & Session("Nom_clie") & " " & cortesia & "</b></font>"
	html2 = img("../imagenes/botonsalir.gif", "125", "28", "", "", "")
	salir = enlace("", "../salir.asp", "border-style:none", "", html2, "")
	Nom_cli = tag_class("label", "", "", "", fuente, "")
	Dim ind
	z = load.getValue("aleat")
	
	'Insertar, borrar o editar'
	if val_inser = "" then
		val_inser = "0"
	End if
	ed_in_bor = load.getValue("edit_inser_borr")
'RESPONSE.WRITE "codigo-------" & ed_in_bor 	
	nume_trip = load.getValue("num_trip")
	
	
	'response.write "numero" & nume_trip
	'Usando Session ("rand")'
	'Es la mejor solución hasta el momento para evitar que se inserte'
	'al hacer un refresh'
	if z <> Session ("rand") then
		Session ("rand") = z
		refresco = false
	else
		refresco = true
	End if
	guard = load.getValue("env")
	Set cnnf = Server.CreateObject("ADODB.Connection")
	Set rstf = Server.CreateObject("ADODB.Recordset")
	Set rstr = Server.CreateObject("ADODB.Recordset")
	Set rste = Server.CreateObject("ADODB.Recordset")
	Set rstc = Server.CreateObject("ADODB.Recordset")
	
	cnnf.Open dsn()
	
	nom_aper = ""
	nom_ape2r = ""
	cedular = ""
	cedula2r = ""
	ape_tripr = ""
	ape_trip2r = ""
	
	fec_nac_trir = ""
	num_lic_trir = ""
	num_lic_tri2r = ""
	num_cer_trir= ""
	num_cer_tri2r= ""
	fec_venc_licr = ""
	fec_venc_cerr = ""
	img_lic_trir = ""
	
	'+++++++++++++
	num_lic_eua_trir = ""
	num_lic_eua_tri2r = ""
	'fec_venc_eua_licr = ""
	img_lic_eua_trir = ""
	'+++++++++++++

	'++++++Nuevo+++++++
	num_cer_eua_trir = ""
	num_cer_eua_tri2r = ""
	fec_exp_eua_cerr = ""
	img_cer_eua_trir = ""
	'+++++++++++++
	
	img_cer_trir = ""
	nac_trir = ""
	cod_nac_trir = ""
	tel_celr = ""
	e_mailr = ""
	e_mail2r = ""
	num_pasapr = ""
	num_pasap2r = ""
	fec_vencr = ""
	img_pasapr = ""
	num_visa_euar = ""
	num_visa_eua2r = ""
	fec_venc_visa_euar = ""
	img_visa_euar = ""
	'Se comenta tip_trip porque
	'se considera que se puede volver a usar
	'Se mantiene en B/D
	'tip_tripr = ""
	activor = ""
	
	filtrobusqueda = ""
	filtrobusqueda = limpia(load.getValue("busquedaPas"))	
	'response.write filtrobusqueda

	
	'Humberto Rojas 19032014
	sexor = ""
	sexorF = ""
	paisResr = ""
	direccionr = ""
	edo_vzla_tripr = ""
	ciudadtripr = ""
	tipovisar = ""
	fec_exp_visar = ""
	ciudadvisar = ""
	cod_paisResr = ""
	cod_ciudadtripr = "" 
	cod_edo_vzla_tripr = ""
	
	if guard = "1" then
		'Se verifica que se hizo un submit en vez de un'
		'Refresh'
		if refresco = false then
			
			Call crear_carpeta(Session("Nom_clie") & "_" & Session("Num_clie"))
			err_nums =  0
			nom_aper = limpia(load.getValue("nom_ape"))
			nom_ape2r = limpia(load.getValue("nom_ape2"))
			cedular = limpia(load.getValue("cedula"))
			cedula2r = limpia(load.getValue("cedula2"))
			ape_tripr = limpia(load.getValue("Ape_trip"))
			ape_trip2r = limpia(load.getValue("Ape_trip2"))

			fec_nac_trir = limpia(load.getValue("fec_nac_tri")) 
			num_lic_trir = limpia(load.getValue("num_lic_tri"))
			num_lic_tri2r = limpia(load.getValue("num_lic_tri2"))
			num_cer_trir = limpia(load.getValue("num_cer_tri"))
			num_cer_tri2r = limpia(load.getValue("num_cer_tri2"))
			'Nuevos datos
			num_cer_eua_trir = limpia(load.getValue("num_cer_eua_tri"))
			num_cer_eua_tri2r = limpia(load.getValue("num_cer_eua_tri2"))
			fec_exp_eua_cerr = limpia(load.getValue("fec_exp_eua_cer"))
			img_cer_eua_trir = limpia(load.getValue("img_cer_eua_tri"))
			
			num_lic_eua_trir = limpia(load.getValue("num_lic_eua_tri"))
			num_lic_eua_tri2r = limpia(load.getValue("num_lic_eua_tri2")) 
			'No hay fecha de vencimiento en la licencia de EUA
			'fec_venc_eua_licr = "01/01/3000" limpia(load.getValue("fec_venc_eua_lic"))
			img_lic_eua_trir = limpia(load.getValue("img_lic_eua_tri"))
			
			fec_venc_licr = limpia(load.getValue("fec_venc_lic"))
			fec_venc_cerr = limpia(load.getValue("fec_venc_cer"))
			img_lic_trir = limpia(load.getValue("img_lic_tri"))
			img_cer_trir = limpia(load.getValue("img_cer_tri"))
			img_cer_eua_trir = limpia(load.getValue("img_cer_eua_tri"))
			nac_trir = load.getValue("nac_tri")
			cod_nac_trir = load.getValue("cod_nac_tri")

			cod_pais = cod_nac_trir 'selectSimpleSQL2("Cod_pais", "Siep_pais", " Cod_pais='" & cod_nac_trir & "'", 1)'
			tel_celr = limpia(load.getValue("tel_cel")) 
			e_mailr = limpia(load.getValue(LCase(("e_mail"))))
			e_mail2r = limpia(load.getValue(Lcase(("e_mail2"))))
			num_pasapr = limpia(load.getValue("num_pasap"))
			num_pasap2r = limpia(load.getValue("num_pasap2"))
			fec_vencr = limpia(load.getValue("fec_venc"))
			img_pasapr = limpia(load.getValue("img_pasap"))
			num_visa_euar = limpia(load.getValue("num_visa_eua"))
			num_visa_eua2r = limpia(load.getValue("num_visa_eua2"))
			fec_venc_visa_euar = limpia(load.getValue("fec_venc_visa_eua"))
			
			
			img_visa_euar = limpia(load.getValue("img_visa_eua"))
			'tip_tripr = load.getValue("tip_trip")
			activor = load.getValue("activo")
			'Request.ServerVariables("ALL_RAW")'
			
			'Humberto Rojas 19-03-2014
			sexor = limpia(load.getValue("sexo"))
			sexorF =  limpia(load.getValue("sexo2"))

			
			
			paisResr = limpia(load.getValue("paisRes"))
			direccionr = limpia(load.getValue("direccion"))
			edo_vzla_tripr = limpia(load.getValue("edo_vzla_trip"))
			ciudadtripr = limpia(load.getValue("ciudadtrip"))
			tipovisar = limpia(load.getValue("tipovisa"))
			fec_exp_visar = limpia(load.getValue("fec_exp_visa"))
			ciudadvisar = limpia(load.getValue("ciudadvisa"))
			cod_paisResr = limpia(load.getValue("cod_paisRes"))
            

			
			if (CInt(ed_in_bor) <> 2) And (cod_paisResr <> "") And (paisResr <> "") then
				
				if (edo_vzla_tripr <> "") then
					rstf.Open  "Select Cod_esta From Siep_esta Where (Cod_pais = " & cod_paisResr & ") And  (Des_esta = '" & edo_vzla_tripr & "')", cnnf, 1, 2
					if rstf.EOF then	
						Cod_esta = ""
					else
						Cod_esta = rstf("Cod_esta")
					End if
					rstf.Close
					if Cod_esta = "" then
						rstf.Open  "Siep_esta", cnnf, 1, 2
						rstf.AddNew
						rstf("Cod_pais") = cod_paisResr
						rstf("Des_esta") = edo_vzla_tripr
						rstf.Update
						rstf.Close
						rstf.Open  "Select Last(Cod_esta) As codE From Siep_esta", cnnf, 1, 2
						Cod_esta = rstf("codE")
						rstf.Close
					End if
				End if
				
				if (ciudadtripr <> "") then
					rstf.Open  "Select Cod_ciuda From Ciudad Where (Cod_pais = " & cod_paisResr & ") And  (Nom_ciuda = '" & ciudadtripr & "')", cnnf, 1, 2
					if rstf.EOF then	
						Cod_ciu = ""
					else
						Cod_ciu = rstf("Cod_ciuda")
					End if
					rstf.Close
					if Cod_ciu = "" then
						rstf.Open  "Select Max(Cod_ciuda) As CodMax From Ciudad", cnnf, 1, 2
						CodMax = CInt(rstf("CodMax")) + 1
						rstf.Close
						rstf.Open  "Ciudad", cnnf, 1, 2
						rstf.AddNew
						rstf("Cod_ciuda") = CStr(CodMax)
						rstf("Cod_pais") = cod_paisResr
						rstf("Nom_ciuda") = ciudadtripr
						rstf.Update
						rstf.Close
						rstf.Open  "Select Last(Cod_ciuda) As codC From Ciudad", cnnf, 1, 2
						Cod_ciu = rstf("codC")
						rstf.Close
					End if
				End if
			End if		
				cod_ciudadtripr = Cod_ciu 'limpia(load.getValue("cod_ciudadtrip")) 
				cod_edo_vzla_tripr = Cod_esta 'limpia(load.getValue("cod_edo_vzla_trip")) 
				
				'//////////////////////////////////////'			
				'Solo nombre y tipo de tripulación no pueden ser null'
				'No se admiten pasaporte, cédula, visa y/o correo repetidos '
				'para un mismo cliente si no son null'
				'////////////////////////////////////'

				if Tabla_Vacia("Siep_trip", " Nom_trip = '" & nom_aper & "' And Ape_trip = '" & ape_tripr & "' And Num_clie = " & Session("Num_clie")) = false And (nom_aper <> nom_ape2r And ape_tripr <> ape_trip2r) then
					call alerta("El nombre " & nom_aper & " " & ape_tripr & " se encuentra registrado ")
					err_nums =  err_nums + 1
				End if 
				if tel_celr = "" then
					tel_celr = null
				End if
				if cedular <> "" And (cedular <> cedula2r) then
					if Tabla_Vacia("Siep_trip", " Num_cedu_tri = '" & cedular & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("La cédula " & cedular & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if
				End if
				
				if cedular = "" then
					cedular = null
				End if
				if fec_nac_trir = "" then
					fec_nac_trir = null
				End if
				if num_lic_trir <> "" And (num_lic_trir <> num_lic_tri2r) then
					if Tabla_Vacia("Siep_trip", " Num_lic_tri = '" & num_lic_trir & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("La licencia " & num_lic_trir & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if
				End if
				
				if num_lic_trir = "" then
					num_lic_trir = null
				End if
			
				if fec_venc_licr = "" then
					fec_venc_licr = null
				End if			
				if img_lic_trir = "" And CInt(ed_in_bor) = 0 then
					img_lic_trir = null
				End if
			
				if img_lic_trir <> "" then
					'Por el momento se monta en B/D'
					fileDataLic = load.getFileData("img_lic_tri")
					Dim fileNameLic
					fileNameLic = LCase(load.getFileName("img_lic_tri"))
					Dim filePathLic
					filePathLic = load.getFilePath("img_lic_tri")
					Dim filePathCompleteLic
					filePathCompleteLic = load.getFilePathComplete("img_lic_tri")
					Dim fileSizeLic
					fileSizeLic = load.getFileSize("img_lic_tri")
					if fileSizeLic > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						img_lic_trir = ""
						err_nums =  err_nums + 1
					else
						Dim fileSizeTranslatedLic
						fileSizeTranslatedLic = load.getFileSizeTranslated("img_lic_tri")
						Dim contentTypeLic
						contentTypeLic = load.getContentType("img_lic_tri")
						Dim nameInputLic
						nameInputLic = load.getValue("img_lic_tri")
						Dim pathToFileTri
						Dim carpetaTri
						carpetaTri = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFileTri = CStr(Server.mapPath(carpetaTri) & "\" & fileNameLic) '"
					End if
				End if
					
				'/////////////Datos Nuevos///////////////////////
			
				if num_cer_eua_trir = "" then
					num_cer_eua_trir = null
				End if
			
				if fec_exp_eua_cerr = "" then
					fec_exp_eua_cerr = null
				End if
			
				if img_cer_eua_trir = "" And CInt(ed_in_bor) = 0 then
					img_cer_eua_trir = null
				End if
				
				if img_cer_eua_trir <> "" then
					'Por el momento se monta en B/D'
					fileDataCerEUA = load.getFileData("img_cer_eua_tri")
					Dim fileNameCerEUA
					fileNameCerEUA = LCase(load.getFileName("img_cer_eua_tri"))
					Dim filePathCerEUA
					filePathCerEUA = load.getFilePath("img_cer_eua_tri")
					Dim filePathCompleteCerEUA
					filePathCompleteCerEUA = load.getFilePathComplete("img_cer_eua_tri")
					Dim fileSizeCerEUA
					fileSizeCerEUA = load.getFileSize("img_cer_eua_tri")
					
					if fileSizeCerEUA > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						img_cer_eua_trir = ""
						err_nums =  err_nums + 1
					else
						Dim fileSizeTranslatedCerEUA
						fileSizeTranslatedCerEUA = load.getFileSizeTranslated("img_cer_eua_tri")
						Dim contentTypeCerEUA
						contentTypeCerEUA = load.getContentType("img_cer_eua_tri")
						Dim nameInputCerEUA
						nameInputCerEUA = load.getValue("img_cer_eua_tri")
						Dim pathToFileCerEUA
						Dim carpetaCerEUA
						carpetaCerEUA = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFileCerEUA = CStr(Server.mapPath(carpetaCerEUA) & "\" & fileNameCerEUA) '"
					End if		
				End if
				
				if num_cer_eua_trir <> "" And (num_cer_eua_trir <> num_cer_eua_tri2r) then
					if Tabla_Vacia("Siep_trip", " Num_cer_eua_tri = '" & num_cer_eua_trir & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El certificado " & num_cer_eua_trir & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if
				End if
				
				if num_cer_trir = "" then
					num_cer_trir = null
				End if
			
				if fec_venc_cerr = "" then
					fec_venc_cerr = null
				End if
			
				if img_cer_trir = "" And CInt(ed_in_bor) = 0 then
					img_cer_trir = null
				End if
				
				'''''''''''''''''''''''''''''
				
				if img_cer_trir <> "" then
					'Por el momento se monta en B/D'
					fileDataCer = load.getFileData("img_cer_tri")
					Dim fileNameCer
					fileNameCer = LCase(load.getFileName("img_cer_tri"))
					Dim filePathCer
					filePathCer = load.getFilePath("img_cer_tri")
					Dim filePathCompleteCer
					filePathCompleteCer = load.getFilePathComplete("img_cer_tri")
					Dim fileSizeCer
					fileSizeCer = load.getFileSize("img_cer_tri")
					if fileSizeCer > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						img_cer_trir = ""
						err_nums =  err_nums + 1
					else
						Dim fileSizeTranslatedCer
						fileSizeTranslatedCer = load.getFileSizeTranslated("img_cer_tri")
						Dim contentTypeCer
						contentTypeCer = load.getContentType("img_cer_tri")
						Dim nameInputCer
						nameInputCer = load.getValue("img_cer_tri")
						Dim pathToFileCer
						Dim carpetaCer
						carpetaCer = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFileCer = CStr(Server.mapPath(carpetaCer) & "\" & fileNameCer) '"
					End if
				End if
				
				if num_cer_trir <> "" And (num_cer_trir <> num_cer_tri2r) then
					if Tabla_Vacia("Siep_trip", " Num_cer_tri = '" & num_cer_trir & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El certificado m\u00E9dico" & num_cer_trir & " se encuentra registrado ")
						err_nums =  err_nums + 1
					End if
				End if
			
				if num_lic_eua_trir <> "" And (num_lic_eua_trir <> num_lic_eua_tri2r) then
					if Tabla_Vacia("Siep_trip", " Num_lic_eua_tri = '" & num_lic_eua_trir & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("La licencia " & num_lic_eua_trir & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if
				End if
			
				if num_lic_eua_trir = "" then
					num_lic_eua_trir = null
				End if
			
				'No hay fecha de vencimiento de licencia de EUA (30/01/2014)
			
				if img_lic_eua_trir = "" And CInt(ed_in_bor) = 0 then
					img_lic_eua_trir = null
				End if
				
				if img_lic_eua_trir <> "" then
					'Por el momento se monta en B/D'
					fileDataLicEua = load.getFileData("img_lic_eua_tri")
					Dim fileNameLicEua
					fileNameLicEua = LCase(load.getFileName("img_lic_eua_tri"))
					Dim filePathLicEua
					filePathLicEua = load.getFilePath("img_lic_eua_tri")
					Dim filePathCompleteLicEua
					filePathCompleteLicEua = load.getFilePathComplete("img_lic_eua_tri")
					Dim fileSizeLicEua
					fileSizeLicEua = load.getFileSize("img_lic_eua_tri")
				
					if fileSizeLicEua > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						img_lic_eua_trir = ""
						err_nums =  err_nums + 1
					else
						Dim fileSizeTranslatedLicEua
						fileSizeTranslatedLicEua = load.getFileSizeTranslated("img_lic_eua_tri")
						Dim contentTypeLicEua
						contentTypeLicEua = load.getContentType("img_lic_eua_tri")
						Dim nameInputLicEua
						nameInputLicEua = load.getValue("img_lic_eua_tri")
						Dim pathToFileLicEua
						Dim carpetaLicEua
						carpetaLicEua = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFileLicEua = CStr(Server.mapPath(carpetaLicEua) & "\" & fileNameLicEua) '"
						'img_lic_eua_trir = pathToFileLicEua
					End if
				End if
				
				if e_mailr <> "" And (e_mailr <> e_mail2r) then
					if Tabla_Vacia("Siep_trip", " Ema_trip = '" & e_mailr & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El correo " & e_mailr & " se encuentra registrado ")
						err_nums =  err_nums + 1
					End if
				End if
				if e_mailr = "" then
					e_mailr = null
				End if
				if num_pasapr <> "" And (num_pasapr <> num_pasap2r) then
					if Tabla_Vacia("Siep_trip", " Num_pasa_tri = '" & num_pasapr & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El pasaporte " & num_pasapr & " se encuentra registrado ")
						err_nums =  err_nums + 1
					End if
				End if
				if e_mailr = "" And CInt(ed_in_bor) = 0 then
					num_pasapr = null
				End if
				if num_visa_euar <> "" And (num_visa_euar <> num_visa_eua2r) then
					if Tabla_Vacia("Siep_trip", " Num_visa_tri = '" & num_visa_euar & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("La visa " & num_visa_euar & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if 
				End if
				if num_visa_euar = "" then
					num_visa_euar = null
				End if
				
				if fec_vencr = "" then	
					fec_vencr = null
				End if
				if img_pasapr = "" And CInt(ed_in_bor) = 0 then
					img_pasapr = null
				End if
				
				if img_pasapr <> "" then
					'Por el momento se monta en B/D'
					fileData = load.getFileData("img_pasap")
					Dim fileName
					fileName = LCase(load.getFileName("img_pasap"))
					Dim filePath
					filePath = load.getFilePath("img_pasap")
					Dim filePathComplete
					filePathComplete = load.getFilePathComplete("img_pasap")
					Dim fileSize
					fileSize = load.getFileSize("img_pasap")
					if fileSize > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						img_pasapr = ""
						err_nums =  err_nums + 1
					else
						Dim fileSizeTranslated
						fileSizeTranslated = load.getFileSizeTranslated("img_pasap")
						Dim contentType
						contentType = load.getContentType("img_pasap")
						Dim nameInput
						nameInput = load.getValue("img_pasap")
						Dim pathToFile
                	
						Dim carpeta
						carpeta = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFile = CStr(Server.mapPath(carpeta) & "\" & fileName) '"
					End if
				End if
				
				if num_visa_euar = "" And CInt(ed_in_bor) = 0 then
					num_visa_euar = null
				End if
				if fec_venc_visa_euar = "" then
					fec_venc_visa_euar = null
				End if
				if img_visa_euar = ""  And  CInt(ed_in_bor) = 0 then
					img_visa_euar = null
				End if
				if tip_tripr = "" then
					tip_tripr = null
				End if 
				
				if img_visa_euar <> "" then
					fileData2 = load.getFileData("img_visa_eua")
					Dim fileName2
					fileName2 = LCase(load.getFileName("img_visa_eua"))
					Dim filePath2
					filePath2 = load.getFilePath("img_visa_eua")
					Dim filePathComplete2
					filePathComplete2 = load.getFilePathComplete("img_visa_eua")
					Dim fileSize2
					fileSize2 = load.getFileSize("img_visa_eua")
					if fileSize2 > 100000 then
						Call alerta("No se puede subir un archivo de mas de 100 KB")
						err_nums =  err_nums + 1
						img_visa_euar = ""
					else
						Dim fileSizeTranslated2
						fileSizeTranslated2 = load.getFileSizeTranslated("img_visa_eua")
						Dim contentType2
						contentType2 = load.getContentType("img_visa_eua")
						Dim nameInput2
						nameInput2 = load.getValue("img_visa_eua")
						Dim pathToFile2
						Dim ruta2
						Dim carpeta2
						carpeta2 = Session("Nom_clie") & "_" & Session("Num_clie") & "/"
						pathToFile2 = CStr(Server.mapPath(carpeta2) & "\" & fileName2) '"
					End if
				End if
				
				longi_tot = fileSize2 + fileSize + fileSizeLicEua + fileSizeCer + fileSizeLic
			
				if longi_tot > 650000 Then
					Call alerta("El peso total de las imagenes a montar supera los 650 KB")
					err_nums =  err_nums + 1
				end if
			
				
			if err_nums = 0 then
					Select case CInt(ed_in_bor)
						case 0, 1
							
							
							'/////////////-------------//////////////////////'
							if CInt(ed_in_bor) = 0 then
								rstf.Open  "Siep_trip", cnnf, 1, 2
								rstf.AddNew
								rstf("Fec_clie_reg") = CStr(Session("Fec_clie_reg"))
								rstf("Num_clie") = CStr(Session("Num_clie"))
							Else
								rstf.Open  "Select * From Siep_trip Where Num_trip = " & nume_trip, cnnf, 1, 2
								val_inser = "0"
							End if
							rstf("Nom_trip") = CStr(nom_aper)
							rstf("Ape_trip") = CStr(ape_tripr)
							rstf("Num_cedu_tri") = cedular
							rstf("Tlf_trip_mov") = tel_celr
							rstf("Ema_trip") = LCase(e_mailr)
							rstf("Num_pasa_tri") = num_pasapr
							rstf("Num_visa_tri") = num_visa_euar
							rstf("Fec_venc_pas") = fec_vencr
							rstf("Fec_venc_vis") = fec_venc_visa_euar
							'Humberto Rojas 19032014
							'response.Write sexor
							'response.End()
                        	
							if sexor <> "0" and sexor <>"1" then
								sexor ="0"
							end if
		                	
                        	
							rstf("Sex_trip") = sexor '*	
							
							
							'response.Write sexor
							'response.End()
												
							if cod_paisResr = "" then
							   cod_paisResr = null
							end if
							rstf("Cod_pais_resd") = cod_paisResr '*	
							rstf("Dir_trip") = direccionr
							if cod_edo_vzla_tripr = "" then
							   cod_edo_vzla_tripr =null
							end if						
							rstf("Cod_esta_trip") = cod_edo_vzla_tripr '*
							if cod_ciudadtripr = "" then
							   cod_ciudadtripr = null
							end if	
							rstf("Cod_ciu_trip") = cod_ciudadtripr '*
							
							rstf("Tip_visa") = tipovisar
							if fec_exp_visar = "" then
							   fec_exp_visar = null
							end if 
							rstf("Fec_exp_visa") = fec_exp_visar	
							rstf("Ciu_exp_visa") = ciudadvisar
							
							
							if Not IsNull(img_cer_trir) And img_cer_trir <> "" then
								rstf("Ima_cer_tri") = img_cer_trir
								'rstf("Bin_cer_tri").AppendChunk fileDataCer
								rstf("Tip_con_cer") = contentTypeCer
							End if 
							if Not IsNull(img_cer_eua_trir) And img_cer_eua_trir <> "" then
								rstf("Ima_cer_eua_tri") = img_cer_eua_trir
								'rstf("Bin_cer_eua_tri").AppendChunk fileDataCerEUA
								rstf("Tip_con_eua_cer") = contentTypeCerEUA
							End if
							if Not IsNull(img_lic_trir) And img_lic_trir <> "" then
								rstf("Ima_lic_tri") = img_lic_trir
								'rstf("Bin_lic_tri").AppendChunk fileDataLic
								rstf("Tip_con_lic") = contentTypeLic
							End if
                        	
							if Not IsNull(img_lic_eua_trir) And img_lic_eua_trir <> "" then
								rstf("Ima_lic_eua_tri") = img_lic_eua_trir
								'rstf("Bin_lic_eua_tri").AppendChunk fileDataLicEua
								rstf("Tip_cont_eua_lic") = contentTypeLicEua
							End if
																	
							if Not IsNull(img_pasapr) And img_pasapr <> "" then
								rstf("Ima_pasa") = img_pasapr
								
								'****************************************************************************response.write img_pasapr
								
								
								'rstf("Bin_pas").AppendChunk fileData
								rstf("Tip_cont_pas") = contentType
							End if
							
							if Not IsNull(img_visa_euar) And img_visa_euar <> "" then
								rstf("Ima_visa") = img_visa_euar
								
								'****************************************************************************response.write img_visa_euar
								
								'rstf("Bin_visa").AppendChunk fileData2
								rstf("Tip_cont_vis") = contentType2
							End if
							
							rstf("Fec_nac_tri") = fec_nac_trir
							rstf("Num_lic_tri") = num_lic_trir 
							rstf("Num_lic_eua_tri") = num_lic_eua_trir
							'30/01/2014 -> No hay fecha de vencimiento en la licencia de EUA
							'rstf("Fec_venc_eua_lic") = fec_venc_eua_licr
							rstf("Num_cer_tri") = num_cer_trir
							'Nuevo
							rstf("Num_cer_eua_tri") = num_cer_eua_trir
							rstf("Fec_venc_lic") = fec_venc_licr
							rstf("Fec_venc_cer") = fec_venc_cerr
							rstf("Fec_exp_eua_cer") = fec_exp_eua_cerr
							rstf("Fec_nac_tri") = fec_nac_trir
							if cod_pais = "" then
							   cod_pais = null
							end if
							rstf("Nac_tri") = cod_pais
							'Ernesto escribio la sentencia siguiente para que guardara directo Tip_trip como 1|PIC ya que esaba dancd error'
							rstf("Tip_trip") = "1"
                        	
							if activor = "on" then
								rstf("Sta_acti") = "1"
							else
								rstf("Sta_acti") = "0"
							End if
							'rstf("Tip_trip") = tip_tripr
							rstf.Update
							rstf.Close
							
						case 2
							nums = load.GetValue("lista")
							msj_err = 0
							'msj = "El o los integrantes(s) "
							'for fila = 1 to nums-1
							'	num_tripx = "selct_fila" & CStr(fila) & "_0"
							'	selec = "selct_fila" & CStr(fila)
							'	if load.GetValue(selec) = "on" then 
									'if Tabla_Vacia("Siep_matr",  "Num_trip  = " & load.GetValue(num_tripx)) = true then
									'if Tabla_Vacia("Siep_matr",  "Num_trip  = " & load.GetValue(nume_trip)) = true then
									if Tabla_Vacia("Siep_matr",  "Num_trip  = " & nume_trip) = true then	
										'rstf.Open "Delete From Siep_trip Where Num_trip  = " & load.GetValue(num_tripx), cnnf, 1, 2
										'rstf.Open "Delete From Siep_trip Where Num_trip  = " & load.GetValue(nume_trip), cnnf, 1, 2
										rstf.Open "Delete From Siep_trip Where Num_trip  = " & nume_trip, cnnf, 1, 2
									else
										msj_err = msj_err + 1
										'msj = msj & load.GetValue("selct_fila" & fila & "_1") & ", "
									end if
							'	End if
							'next
							if msj_err > 0 then
								'msj = left(msj, len(msj)-2)
								'msj = msj & " no puede(n) ser eliminado(s). Solo se puede(n) colocar en estatus inactivo"
								'Call alerta(msj)
								Call alerta("El usuario no puede ser eliminado solo se puede colocar en estatus 'inactivo'")
							End if
							val_inser = "0"	
					End Select
					
												'//////////////////-------//////////////////'
							
							nume_pasa = ""
							nom_aper = ""
							nom_ape2r = ""
							cedular = ""
							cedula2r = ""
							Ape_tripr = ""
							Ape_trip2r = ""
							
							fec_nac_trir = ""
							num_lic_trir = ""
							num_lic_tri2r = ""
							num_cer_trir= ""
							num_cer_tri2r= ""
							fec_venc_licr = ""
							fec_venc_cerr = ""
							img_lic_trir = ""
							img_cer_trir = ""
							nac_trir = ""
							cod_nac_trir = ""
							
							'+++++++++++++
							num_lic_eua_trir = ""
							num_lic_eua_tri2r = ""
							'fec_venc_eua_licr = "" -> No tiene (30/01/2014)
							img_lic_eua_trir = ""
							'+++++++++++++
						
							'+++++++++++++
							num_cer_eua_trir = ""
							num_cer_eua_tri2r = ""
							fec_exp_eua_cerr = ""
							img_cer_eua_trir = ""
							'+++++++++++++
							
							tel_celr = ""
							e_mailr = ""
							e_mail2r = ""
							num_pasapr = ""
							num_pasap2r = ""
							fec_vencr = ""
							img_pasapr = ""
							num_visa_euar = ""
							num_visa_eua2r = ""
							fec_venc_visa_euar = ""
							img_visa_euar = ""
							activor = ""
							
							sexor = ""
							sexorF = ""
							cod_paisResr = ""
							direccionr = ""
							cod_edo_vzla_tripr = ""
							cod_ciudadtripr = ""
							tipovisar = ""
							fec_exp_visar = ""	
							ciudadvisa = ""
							
							paisRes = ""
							direccion = ""
							edo_vzla_trip = ""
							ciudadtrip = ""
							tipovisa = ""
							fec_exp_visa = ""
						
							paisResr = ""
							edo_vzla_tripr = ""
							ciudadtripr = ""
							ciudadvisar = ""
					
				else
				
					'////////////////
					
					'if CInt(ed_in_bor) = 1 then
					'	m = load.getValue("lista")
					'	Redim lista_check(m-1)
					'	For j = 1 to m-1
					'		if load.getValue("selct_fila" & j) = "on" then
					'			lista_check(j) = "checked"
					'		Else
					'			lista_check(j) = ""
					'		End if
					'	Next
					'End if
					val_inser = ed_in_bor

					'////////////////
				End if
				
			'End if	-> 252	
		End if
	End if
	

	
	thidden = ""
	thidden = thidden & hidden("num_trip", nume_trip)
	
	imgcal = img("../images_cal/cal.gif", "16", "16", "Pick a date", "", "")
	
	'<input name="nom_ape" type="text" id="nom_ape" size="20" maxlength="20"/>'
	nom_ape = campo("nom_ape", "text", "nom_ape", "30", "50", "1", nom_aper, "OnBlur='caractNoPermit(this)'")
	nom_ape = nom_ape & hidden("nom_ape2", nom_ape2r)
	ast_nom_ape = tag("span", "ast_nom_ape", "visibility:hidden", marca_error(), "")
	
	
	Ape_trip = campo("Ape_trip", "text", "Ape_trip", "30", "50", "2", ape_tripr, "OnBlur='caractNoPermit(this)'")
	Ape_trip = Ape_trip & hidden("Ape_trip2", Ape_trip2r)
	ast_ape = tag("span", "ast_ape", "visibility:hidden", marca_error(), "")
	
	'<inputast_cedula name="cedula" type="text" id="cedula" size="20" maxlength="20"/>'
	cedula = campo("cedula", "text", "cedula", "30", "12", "3", cedular, "OnBlur='verCedula(this)'")
	cedula = cedula & hidden("cedula2", cedula2r)
	ast_cedula = tag("span", "ast_cedula", "visibility:hidden", marca_error(), "")
	
	'sexo = radio("sexo", "sexo", "0","")
	'sexoF = radio("sexo", "sexo", "1","")
	'sexo = sexo & hidden("sexo2", sexo2r)
	'ast_sexo = tag("span", "ast_sexo", "visibility:hidden", marca_error(), "")
	
	'sexo = radio("sexo", "sexo", "0", "") 
	'sexo2 = radio("sexo", "sexo", "1",  "")
	
	sexo = radio("sexo", "sexo", "0", "") 
	sexo2 = radio("sexo", "sexo", "1",  "")
	
	
	
	
	
	'sexo = radio("sexo", "sexo", "0","")
	'sexoF = radio2("sexoF", "sexoF", "1","")
	'sexoF = sexoF & hidden("sexoF2", sexoF2r)
	'ast_sexoF = tag("span", "ast_sexoF", "visibility:hidden", marca_error(), "")
	
	
	
	paisRes = campo("paisRes", "text", "paisRes", "30", "15", "4", paisResr, "OnBlur='caractNoPermit(this)'") &    hidden("cod_paisRes", "")
	ast_paisRes = tag("span", "ast_paisRes", "visibility:hidden", marca_error(), "")
	
	direccion = campo("direccion", "text", "direccion", "30", "100", "5", direccionr, "")
	direccion = direccion & hidden("direccion2", direccion2r)
	ast_direccion = tag("span", "ast_direccion", "visibility:hidden", marca_error(), "")
	
	
	edo_vzla_trip = campo("edo_vzla_trip", "text", "edo_vzla_trip", "30", "30", "6", edo_vzla_tripr, "OnBlur='caractNoPermit(this)'") & 	hidden("cod_edo_vzla_trip", "")
	ast_edo_vzla_trip = tag("span", "ast_edo_vzla_trip", "visibility:hidden", marca_error(), "")
	
	ciudadtrip = campo("ciudadtrip", "text", "ciudadtrip", "30", "15", "7", ciudadtripr, "OnBlur='caractNoPermit(this)'") & 	hidden("cod_ciudadtrip", "")
	ast_ciudadtrip = tag("span", "ast_ciudadtrip", "visibility:hidden", marca_error(), "")
		
	
	tipovisa = campo("tipovisa", "text", "tipovisa", "30", "12", "25", tipovisar, "")
	tipovisa = tipovisa & hidden("tipovisa2", tipovisa2r)
	ast_tipovisa = tag("span", "ast_tipovisa", "visibility:hidden", marca_error(), "")
	
	fec_exp_visa = campo_readonly("fec_exp_visa", "text", "fec_exp_visa", "30", "15", "26", fec_exp_visar, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_exp_visa"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_exp_visa = tag("span", "ast_fec_exp_visa", "visibility:hidden", marca_error(), "")
	
	ciudadvisa = campo("ciudadvisa", "text", "ciudadvisa", "30", "12", "27", ciudadvisar, "")
	ciudadvisa = ciudadvisa & hidden("ciudadvisa2", ciudadvisa2r)
	ast_ciudadvisa = tag("span", "ast_ciudadvisa", "visibility:hidden", marca_error(), "")
	
	
	'/////////////////////////////'
	'num_lic_tri2r = campo("tel_cel", "text", "tel_cel", "20", "15", "3", tel_celr, "OnBlur='esNum(this)'")'
	'num_cer_tri2r = campo("tel_cel", "text", "tel_cel", "20", "15", "3", tel_celr, "OnBlur='esNum(this)'")'
	'/////////////////////////////'
	
	fec_nac_tri = campo_readonly("fec_nac_tri", "text", "fec_nac_tri", "30", "15", "8", fec_nac_trir, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_nac_tri"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_nac_tri = tag("span", "ast_fec_nac_tri", "visibility:hidden", marca_error(), "")
	
	'Licencia de Venezuela
	'	
	num_lic_tri = campo("num_lic_tri", "text", "num_lic_tri", "30", "15", "9", num_lic_trir, "OnBlur='caractNoPermit(this)'") & hidden("num_lic_tri2", num_lic_tri2r)
	ast_num_lic_tri = tag("span", "ast_num_lic_tri", "visibility:hidden", marca_error(), "")
	fec_venc_lic = campo_readonly("fec_venc_lic", "text", "fec_venc_lic", "30", "15", "10", fec_venc_licr, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_venc_lic"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_lic = tag("span", "ast_fec_venc_lic", "visibility:hidden", marca_error(), "")
	'<input name="img_cer_tri" type="file" id="img_cer_tri" size="30" maxlength="50" />'
	
	'img_lic_tri = campo("img_lic_tri", "file", "img_lic_tri", "30", "50", "11", img_lic_trir, "")
	'ast_img_lic_tri = tag("span", "ast_img_lic_tri", "visibility:hidden", marca_error(), "")
	
	   
	   
	   
	   
		
		
		a = 1 'cambiar este valor por uno que me indique si se esta editando para mostrar el boton ver
		
		'if a > 0 then
		
		   img_lic_tri = campo_readonly("img_lic_tri", "text", "img_lic_tri", "30", "50", "20", img_lic_trir, "readonly", "") & hidden("img_lic_trir", img_lic_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=9"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")  
		   
		     'botonVer2 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			 botonVer2 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		   
		
		'else
		
		
		   'img_lic_tri = campo_readonly("img_lic_tri", "text", "img_lic_tri", "30", "50", "20", img_lic_trir, "readonly", "") & hidden("img_lic_trir", img_lic_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=9"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")
		   
		   
		 
		'end if
		
		'& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=3&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
	
	
	
	'Esta linea se deja en comentario permanente
	'boton_lic = button("boton_lic", "Ver", "OnClick='capa(""capa_lic|capa_lic2|capa_lic3"", ""boton_lic"", 3)'") 
	
	imagenOjo_lic = img2("../imagenes/ojorg abrir.gif", "img_lic", "img_lic", "25", "25", "", "")
	hidden_lic = hidden("hlic", "Ver")
	boton_lic = enlace("", "javascript:void(0)", "text-decoration:none", "", imagenOjo_lic,  "OnClick='capa(""capa_lic|capa_lic2|capa_lic3|capa_lic4"", ""img_lic"", 4, ""hlic"")'")
	
	'Licencia de EUA
	'
	num_lic_eua_tri = campo("num_lic_eua_tri", "text", "num_lic_eua_tri", "30", "15", "12", num_lic_eua_trir, "OnBlur='caractNoPermit(this)'") & hidden("num_lic_eua_tri2", num_lic_eua_tri2r)
	ast_num_lic_eua_tri = tag("span", "ast_num_lic_eua_tri", "visibility:hidden", marca_error(), "")
	fec_venc_eua_lic = "" 'campo_readonly("fec_venc_eua_lic", "text", "fec_venc_eua_lic", "50", "15", "8", fec_venc_eua_licr, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_venc_eua_lic"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_eua_lic = "" 'tag("span", "ast_fec_venc_eua_lic", "visibility:hidden", marca_error(), "")
	'<input name="img_cer_tri" type="file" id="img_cer_tri_eua" size="50" maxlength="50" />'
	
	'img_lic_eua_tri = campo("img_lic_eua_tri", "file", "img_lic_eua_tri", "30", "50", "13", img_lic_eua_trir, "")	
	'ast_img_lic_eua_tri = tag("span", "ast_img_lic_eua_tri", "visibility:hidden", marca_error(), "")
	
	
	
		
	
	

	
	
	
	'a = 1
	
	'if a > 0 then
		
		  		   
		   
		   img_lic_eua_tri = campo_readonly("img_lic_eua_tri", "text", "img_lic_eua_tri", "30", "50", "20", img_lic_eua_trir, "readonly", "") & hidden("img_lic_eua_trir", img_lic_eua_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=10"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=6&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")  
		   
		   'botonVer = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=6&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
		   
		    botonVer = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=6&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		    'mostrar_esconder(0, 'VerLicEua')
		    'botonVer = button("pasp" & CStr(ind), "Ver", "OnClick='mostrar_esconder(1, ""VerLicEua"")'")
		   
		
		'else
		
		
		  'img_lic_eua_tri2 = campo_readonly("img_lic_eua_tri", "text", "img_lic_eua_tri", "30", "50", "20", img_lic_eua_trir, "readonly", "") & hidden("img_lic_eua_trir", img_lic_eua_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=10"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")   
		   
		   
		 
		'end if
	
	
	
	
	
	
	
	
	
	'Esta linea se deja en comentario permanente
	'boton_eua_lic = button("boton_eua_lic", "Ver", "OnClick='capa(""capa_eua_lic|capa_eua_lic2|capa_eua_lic3"", ""boton_eua_lic"", 3)'")  
	imagenOjo_lic_eua = img2("../imagenes/ojorg abrir.gif", "img_lic_eua", "img_lic_eua", "25", "25", "", "")
	hidden_lic_eua = hidden("hlic_eua", "Ver")
	boton_eua_lic = enlace("", "javascript:void(0)", "text-decoration:none", "", imagenOjo_lic_eua,  "OnClick='capa(""capa_eua_lic|capa_eua_lic2"", ""img_lic_eua"", 2, ""hlic_eua"")'")
	
	num_cer_tri = campo("num_cer_tri", "text", "num_cer_tri", "30", "15", "14", num_cer_trir, "OnBlur='caractNoPermit(this)'") & hidden("num_cer_tri2", num_cer_tri2r)
	ast_num_cer_tri = tag("span", "ast_num_cer_tri", "visibility:hidden", marca_error(), "")	

	nac_tri = campo("nac_tri", "text", "nac_tri", "30", "15", "15", nac_trir, "OnBlur='caractNoPermit(this)'") & hidden("cod_nac_tri", "")
	ast_nac_tri = tag("span", "ast_nac_tri", "visibility:hidden", marca_error(), "")

	fec_venc_cer = campo_readonly("fec_venc_cer", "text", "fec_venc_cer", "30", "15", "16", fec_venc_cerr, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_venc_cer"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_cer = tag("span", "ast_fec_venc_cer", "visibility:hidden", marca_error(), "")
	'<input name="img_cer_tri" type="file" id="img_cer_tri" size="30" maxlength="30" />'
	'img_cer_tri = campo("img_cer_tri", "file", "img_cer_tri", "30", "20", "17", img_cer_trir, "")
	'ast_img_cer_tri = tag("span", "ast_img_cer_tri", "visibility:hidden", marca_error(), "")
	
	
	
	
	
	
	
	a = 1
	
	'if a > 0 then
		
		  		   
	  
	
		   
		   
		   img_cer_tri = campo_readonly("img_cer_tri", "text", "img_cer_tri", "30", "50", "20", img_cer_trir, "readonly", "") & hidden("img_cer_trir", img_cer_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=11"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=7&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")  
		   
		     'botonVer3 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=7&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			  botonVer3 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=7&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		   
		   
		
		'else
		
		
	'img_cer_tri = campo_readonly("img_cer_tri", "text", "img_cer_tri", "30", "50", "20", img_cer_trir, "readonly", "") & hidden("img_cer_trir", img_cer_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=11"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")
		   
		   
		 
		'end if
	
	
	
	
	
	
	
	
	
	'Esta linea se deja en comentario permanente
	'boton_cert = button("boton_cert", "Ver", "OnClick='capa(""capa_cert|capa_cert2|capa_cert3"", ""boton_cert"", 3)'") 
	imagenOjo_cert = img2("../imagenes/ojorg abrir.gif", "img_lic_cert", "img_lic_cert", "25", "25", "", "")
	hidden_cert = hidden("hcert", "Ver")
	boton_cert = enlace("", "javascript:void(0)", "text-decoration:none", "", imagenOjo_cert,  "OnClick='capa(""capa_cert|capa_cert2|capa_cert3|capa_cert4"", ""img_lic_cert"", 4, ""hcert"")'")

	'Nuevo
	
	'+++++++++++++
	'num_cer_eua_trir = ""
	'num_cer_eua_tri2r = ""
	'fec_exp_eua_cerr = ""
	'img_cer_eua_trir = ""
	'+++++++++++++
	
	num_cer_eua_tri = campo("num_cer_eua_tri", "text", "num_cer_eua_tri", "30", "15", "10", num_cer_eua_trir, "OnBlur='caractNoPermit(this)'") & hidden("num_cer_eua_tri2", num_cer_eua_tri2r)
	ast_num_cer_eua_tri = tag("span", "ast_num_cer_eua_tri", "visibility:hidden", marca_error(), "")
	fec_exp_eua_cer = campo_readonly("fec_exp_eua_cer", "text", "fec_exp_eua_cer", "30", "15", "12", fec_exp_eua_cerr, "readonly", "") & enlace("", "javascript:NewCssCal(""fec_exp_eua_cer"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_exp_eua_cer = tag("span", "ast_fec_exp_eua_cer", "visibility:hidden", marca_error(), "")
	'-<input name="img_cer_eua_tri" type="file" id="img_cer_eua_tri" size="50" maxlength="50" />'
	'img_cer_eua_tri = campo("img_cer_eua_tri", "file", "img_cer_eua_tri", "30", "20", "13", img_cer_eua_trir, "")
	'ast_img_cer_eua_tri = tag("span", "ast_img_cer_eua_tri", "visibility:hidden", marca_error(), "")
	
	
	
	
	
	
	a = 1
	
	'if a > 0 then
		
		  		   
	   
	
		   
		   img_cer_eua_tri = campo_readonly("img_cer_eua_tri", "text", "img_cer_eua_tri", "30", "50", "20", img_cer_eua_trir, "readonly", "") & hidden("img_cer_eua_trir", img_cer_eua_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=12"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=8&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'") 
		   
		     'botonVer4 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=8&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			  botonVer4 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=8&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		   
		   
		
		'else
		
		
	'img_cer_eua_tri = campo_readonly("img_cer_eua_tri", "text", "img_cer_eua_tri", "30", "50", "20", img_cer_eua_trir, "readonly", "") & hidden("img_cer_eua_trir", img_cer_eua_tri) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=12"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")
		   
		   
		 
		'end if
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	

	'Esta linea se deja en comentario permanente
	'boton_cert_eua = button("boton_cert", "Ver", "OnClick='capa(""capa_cert_eua|capa_cert_eua2|capa_cert_eua3"", ""boton_cert_eua"", 3)'") 
	imagenOjo_cert_eua = img2("../imagenes/ojorg abrir.gif", "img_lic_cert_eua", "img_lic_cert", "25", "25", "", "")
	hidden_cert_eua = hidden("hcerteua", "Ver")
	boton_cert_eua = enlace("", "javascript:void(0)", "text-decoration:none", "", imagenOjo_cert_eua,  "OnClick='capa(""capa_cert_eua|capa_cert_eua2|capa_cert_eua3|capa_cert_eua4"", ""img_lic_cert_eua"", 4, ""hcerteua"")'")

	
	'<input name="tel_cel" type="text" id="tel_cel" size="20" maxlength="20"/>'
	tel_cel = campo("tel_cel", "text", "tel_cel", "30", "15", "16", tel_celr, "OnBlur='esNum(this)'")
	ast_tel_cel = tag("span", "ast_tel_cel", "visibility:hidden", marca_error(), "")
	
	'<input name="e_mail" type="text" id="e_mail" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	e_mail = LCase(campo("e_mail", "text", "e_mail", "30", "50", "17", e_mailr, "OnBlur='caractNoPermit(this)' OnChange='errorCorreo(this)'"))
	e_mail = e_mail & LCase(hidden("e_mail2", e_mail2r))
	ast_e_mail = tag("span", "ast_e_mail", "visibility:hidden", marca_error(), "")
	
	num_pasap = campo("num_pasap", "text", "num_pasap", "30", "20", "18", num_pasapr, "OnBlur='caractNoPermit(this)'")
	num_pasap = num_pasap & hidden("num_pasap2", num_pasap2r)
	ast_num_pasap = tag("span", "ast_num_pasap", "visibility:hidden", marca_error(), "")
	
	'<input name="fec_venc" type="text" id="fec_venc" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	
	fec_venc = campo_readonly("fec_venc", "text", "fec_venc", "30", "20", "19", fec_vencr, "readonly", "")
	ast_fec_venc = tag("span", "ast_fec_venc", "visibility:hidden", marca_error(), "")
	
	'<a href="javascript:NewCssCal('fec_venc', 'ddMMyyyy')">'
	'<img src="../images_cal/cal.gif" width="16" height="16" alt="Pick a date" border="0"></a>'
	cal = enlace("", "javascript:NewCssCal(""fec_venc"", ""ddMMyyyy"")", "", "", imgcal, "")
	
	
	'img_pasap = campo("img_pasap", "file", "img_pasap", "30", "50", "20", img_pasapr, "")
	'ast_img_pasap = tag("span", "ast_img_pasap", "visibility:hidden", marca_error(), "")
	
	
	
		
	
	
	
	
	a = 1
	
	'if a > 0 then
		
		  		   
	   
	
		   
	
		   
		   img_pasap = campo_readonly("img_pasap", "text", "img_pasap", "30", "50", "20", img_pasapr, "readonly", "") & hidden("img_pasapr", img_pasap) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=7"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=3&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'") 
		   
		     'botonVer5 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=3&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			  botonVer5 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=3&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		   
		   
		
		'else
		
		
	'img_pasap = campo_readonly("img_pasap", "text", "img_pasap", "30", "50", "20", img_pasapr, "readonly", "") & hidden("img_pasapr", img_pasap) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=7"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")
		   
		   
		 
		'end if
	
	
	
	
	
	
	
	'<input name="num_visa_eua" type="text" id="num_visa_eua" size="20" maxlength="20" readonly="readonly"/>'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	
	html = img("../Imagenes/boton_agregar.jpg", "70", "19", "", "", "")
	boton_agregar = enlace("", "#", "text-decoration:none", "", html, "OnClick='envia()'")
	'<img src="../imagenes/boton_borrar.jpg" width="70" height="19" /'
	
	num_visa_eua = campo("num_visa_eua", "text", "num_visa_eua", "30", "20", "21", num_visa_euar, "OnBlur='caractNoPermit(this)'")
	num_visa_eua = num_visa_eua & hidden("num_visa_eua2", num_visa_eua2r)
	ast_num_visa_eua = tag("span", "ast_num_visa_eua", "visibility:hidden", marca_error(), "")
	
	'<input name="fec_venc_visa_eua" type="text" id="fec_venc_visa_eua" size="30" maxlength="30"  readonly="readonly"/>'
	fec_venc_visa_eua = campo_readonly("fec_venc_visa_eua", "text", "fec_venc_visa_eua", "30", "50", "22", fec_venc_visa_euar, "readonly", "")
	
	'<a href="javascript:NewCssCal('fec_venc_visa_eua', 'ddMMyyyy')">'
	'<img src="../images_cal/cal.gif" width="16" height="16" alt="Pick a date" border="0"></a>'
	cal2 = enlace("", "javascript:NewCssCal(""fec_venc_visa_eua"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_visa_eua = tag("span", "ast_fec_venc_visa_eua", "visibility:hidden", marca_error(), "")
	
	'<input name="img_visa_eua" type="file" id="324" size="20" maxlength="20"/>'
	
	'img_visa_eua = campo("img_visa_eua", "file", "img_visa_eua", "30", "50", "23", img_visa_euar, "")
	'ast_img_visa_eua = tag("span", "ast_img_visa_eua", "visibility:hidden", marca_error(), "")
	
	
	
	
	
	
	
	
	
	
	a = 1
	
	'if a > 0 then
		
		  		   
	   
	
		   
	
		   

		   
		   
		   img_visa_eua = campo_readonly("img_visa_eua", "text", "img_visa_eua", "30", "50", "20", img_visa_euar, "readonly", "") & hidden("img_visa_euar", img_visa_eua) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=8"",""Visa"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=4&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'") 
		   
		   
		     'botonVer6 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=4&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			  botonVer6 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=4&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
		   
		   
		
		'else
		
		
	'img_visa_eua = campo_readonly("img_visa_eua", "text", "img_visa_eua", "30", "50", "20", img_visa_euar, "readonly", "") & hidden("img_visa_euar", img_visa_eua) & button("bot_matricula", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=8"",""Visa"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'")
		   
		   
		 
		'end if
	
	
	
	'Tipo de tripulación'
	'<!--<select name="12" id="12">'
	'<option value="blanco" selected="selected"></option>'
	'<option value="PIC">PIC</option>'
	'<option value="COPILOTO">COPILOTO</option>'
	'<option value="AEROMOZ@">AEROMOZ@</option></select>-->'
	tip_trip = "" 'seleccion2("tip_trip", "tip_trip", "0|1|2|3", "Seleccione...|PIC|COPILOTO|DE CABINA", tip_tripr, "")
	ast_tip_trip = tag("span", "ast_tip_trip", "visibility:hidden", marca_error(), "")
	
	'<input type="checkbox" name="activo" id="activo"/>'
	'check_box(nombre, id, checked, exprJs)'
	if guard = "1" then
		if activor = "on" then
			activo = check_box("activo", "activo", "checked", "")
		else
			activo = check_box("activo", "activo", "", "")
		End if
	end if
	if guard <> "1" then
		activo = check_box("activo", "activo", "checked", "")
	End if
	
	thidden = thidden & hidden("env", "1")
	thidden = thidden & hidden("aleat", "")
	thidden = thidden & hidden("edit_inser_borr", val_inser)
	
	'<img src="../Imagenes/boton_agregar.jpg" width="70" height="19" />'
	'enlace(id, dir, estilo, clase, html, exprJs)'
	html = img("../Imagenes/boton_agregar.jpg", "70", "19", "", "", "")
	boton_agregar = enlace("", "#", "text-decoration:none", "", html, "OnClick='envia()'")
	'<img src="../imagenes/boton_borrar.jpg" width="70" height="19" /'
	html3 = img("../Imagenes/boton_borrar.jpg", "70", "19", "", "", "")
	'boton_borrar = enlace("", "#", "text-decoration:none", "", html3, "OnClick='elimina(document.form1.lista.value)'")
	boton_borrar = enlace("", "#", "text-decoration:none", "", html3, "OnClick='elimina(document.form1.num_trip.value)'")
	'boton_borrar = ""
	
	tabla = ""
	'tabla = tabla & "<table width='2600' border='2' cellpadding='0' cellspacing='0' bordercolor='#bcbec0' id='table_id' class='display' >"
	'tabla = tabla & "<thead>"
	'tabla = tabla & "<tr>"
	'tabla = tabla & "<th width='45' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
	'tabla = tabla & "<div align='center'>Selec</div></div></th>"
	'tabla = tabla & "<th width='160' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
    'tabla = tabla & "<div align='center'>Nombre</div></div></th>"
	'tabla = tabla & "<th width='160' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
   	'tabla = tabla & "<div align='center'>Apellido</div></div></th>"	
	'tabla = tabla & "<th width='70' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
	'tabla = tabla & "<div align='center'>C&eacute;dula</div></div></th>"
	'tabla = tabla & "<th width='110' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
	'tabla = tabla & "<div align='center'>Tel&eacute;fono</div></div></th>"
	'tabla = tabla & "<th width='200' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
	'tabla = tabla & "<div align='center'>Email</div></div></th>"
	'tabla = tabla & "<th width='350' rowspan='2' bgcolor='#91e1fa' class='Estilo25'>"
	'tabla = tabla & "<div align='center' class='Estilo32'>"
	'tabla = tabla & "<div align='center'>Direcci&oacute;n</div></div></th>"
	'tabla = tabla & "<th colspan='3' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<span class='Estilo32'>Pasaporte</span></div></th>"
    'tabla = tabla & "<th colspan='3' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Visa Americana</div></th>"
	''///////////////////
	'tabla = tabla & "<th colspan='3' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<span class='Estilo32'>Licencia</span></div></th>"
	''///////////////////
	'tabla = tabla & "<th colspan='3' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<span class='Estilo32'>Licencia EEUU</span></div></th>"
	''///////////////////
	'tabla = tabla & "<th colspan='3' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<span class='Estilo32'>Certificado</span></div></th>"
	''///////////////////
	''//////Nuevo/////////////
	'tabla = tabla & "<th colspan='4' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<span class='Estilo32'>Certificado EEUU (Applicant id)</span></div></th>"
	''///////////////////
	''///////////////////
	''tabla = tabla & "<td width='43' rowspan='2' bgcolor='#91e1fa' class='Estilo32'>"
	''tabla = tabla & "<div align='center' class='Estilo25'>Tipo</div></td>"
	''//////////////////
	'tabla = tabla & "<th width='65' rowspan='2' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>"
	'tabla = tabla & "<div align='center'><span class='Estilo32'>Activo</span></div>"
	'tabla = tabla & "</div></th></tr>"
	'tabla = tabla & "<tr>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'><span class='Estilo32'>Vence</span></div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Imagen</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Vence</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>imagen</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'><span class='Estilo32'>Vence</span></div></th>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Imagen</div></td>"
	'tabla = tabla & "<th width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></td>"
	'
	''tabla = tabla & "<td width='82' bgcolor='#91e1fa' class='Estilo32'>"
	''tabla = tabla & "<div align='center' class='Estilo25'>Vence.</div></td>"
	'
	'tabla = tabla & "<td colspan='2' width='57' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>imagen</div></td>"
	''////////////////////////////
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></td>"
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Vence</div></td>"
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>imagen</div></td>"
	''///////////////////////////
    '
	''////////////Nuevo///////////////
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>N&uacute;mero</div></td>"
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Exp.</div></td>"
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>Vence</div></td>"
	'tabla = tabla & "<td width='80' bgcolor='#91e1fa' class='Estilo32'>"
	'tabla = tabla & "<div align='center' class='Estilo25'>imagen</div></td></tr>"
	''////////////////////////////////
	'tabla = tabla & "</thead>"
	ind = 1
	'rstf.Open "Select t.*, p.Nom_pais From Siep_trip t, Siep_pais p Where (t.Nac_tri = p.Cod_pais) And (t.Num_clie = " & Session("Num_clie") & ")", cnnf, 1, 2
	
	'rstf.Open "Select t.*, p.Nom_pais From Siep_trip t  inner join Siep_pais p on t.Nac_tri = p.Cod_pais Where (t.Num_clie = " & Session("Num_clie") & ") Order By Nom_trip", cnnf, 1, 2
	
	hid_tab = ""
	'if  not rstf.EOF then
	'
	'
	'
	'
	'	'Do Until rstf.EOF
	'	Do Until rstf.EOF
	'	
	'	
	'	'rstf("Ape_trip")
	'	'response.End()
	'	 
	'		hid_tab = ""
	'		if err_nums > 0 And ed_in_bor = "1" then
	'			chequeado = lista_check(ind)
	'		Else
	'			chequeado = ""
	'		End if
	'		tabla = tabla & "<tr class='Estilo3'>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'>"
	'		tabla = tabla & "<label>"
	'		tabla = tabla & "<div align='center'>"
	'		tabla = tabla & check_box("selct_fila" & CStr(ind), "selct_fila" & CStr(ind), chequeado, "") & " " & hidden("selct_fila" & CStr(ind) & "_0", rstf("Num_trip")) & "</div></label></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & rstf("Nom_trip") & " " & hidden("selct_fila" & CStr(ind) & "_1", rstf("Nom_trip"))  & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & rstf("Ape_trip") & " " & hidden("selct_fila" & CStr(ind) & "_41", rstf("Ape_trip"))  & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_cedu_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_2", rstf("Num_cedu_tri")) & "</div></td>"
	'       tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Tlf_trip_mov"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_3", rstf("Tlf_trip_mov")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Ema_trip"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_4", rstf("Ema_trip")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Dir_trip"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_42", rstf("Dir_trip")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_pasa_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_5", rstf("Num_pasa_tri")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Fec_venc_pas"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_6", rstf("Fec_venc_pas")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_cont(rstf("Ima_pasa"), button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=3&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_7", rstf("Ima_pasa")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_visa_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_8", rstf("Num_visa_tri")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Fec_venc_vis"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_9", rstf("Fec_venc_vis")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_cont(rstf("Ima_visa"), button("visa" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=4&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_10", rstf("Ima_visa")) & "</div></td>"
	'	
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_lic_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_11", rstf("Num_lic_tri")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Fec_venc_lic"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_12", rstf("Fec_venc_lic")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_cont(rstf("Ima_lic_tri"), button("licencia" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=5&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_13", rstf("Ima_lic_tri")) & "</div></td>"			
	'		'+++++++
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_lic_eua_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_14", rstf("Num_lic_eua_tri")) & "</div></td>"		
	'		tabla = tabla & "<td bgcolor='#e6e8e9' colspan='2'><div align='center'>" & null_cont(rstf("Ima_lic_eua_tri"), button("licencia_eua" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=6&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_15", rstf("Ima_lic_eua_tri")) & "</div></td>"
	'		'+++++++
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_cer_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_16", rstf("Num_cer_tri")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Fec_venc_cer"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_17", rstf("Fec_venc_cer")) & "</div></td>"			
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_cont(rstf("Ima_cer_tri"), button("certificado" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=7&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_18", rstf("Ima_cer_tri")) & "</div></td>"		
	'		
	'		'///////////////////////Nuevo/////////////////////////
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Num_cer_eua_tri"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_19", rstf("Num_cer_eua_tri")) & "</div></td>"
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_bd(rstf("Fec_exp_eua_cer"), "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_20", rstf("Fec_exp_eua_cer")) & "</div></td>"
	'		
	'		
	'		
	'		
	'		
	'		
	'					
	'		if(IsNull(rstf("Fec_exp_eua_cer"))) then
	'			tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>N/D</div></td>"
	'		else
	'			tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & vencCerEUA(rstf("Fec_exp_eua_cer")) & "</div></td>"	
	'		end if
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & null_cont(rstf("Ima_cer_eua_tri"), button("certificado" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=8&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & " " & hidden("selct_fila" & CStr(ind) & "_21", rstf("Ima_cer_eua_tri")) & "</div></td>"	
    '
	'	
	'		hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_22", Replace(rstf("Nom_pais"), "Ã‘", "N")) & hidden("selct_fila" & CStr(ind) & "_23", rstf("Fec_nac_tri"))		
	'		
	'		
	'		
	'		if Not IsNull(rstf("Cod_pais_resd")) or Trim(rstf("Cod_pais_resd")) <> "" then
	'			rstr.Open "Select * From Siep_pais Where Cod_pais = " & rstf("Cod_pais_resd"), cnnf, 1, 2
	'		  	if rstr.EOF then
	'			
	'		
	'			
	'			
	'		  		Pais_res_trip=""
	'		  	else
	'		  		Pais_res_trip = rstr("Nom_pais")
	'		  	End if
	'		  	rstr.close
	'		else
	'	    	Pais_res_trip=""  
	'		end if
	'		
	'		hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_25", Replace(Pais_res_trip, "Ã‘", "N")) & hidden("selct_fila" & CStr(ind) & "_26", Pais_res_trip)					
	'		hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_27", rstf("Dir_trip")) & hidden("selct_fila" & CStr(ind) & "_28", rstf("Dir_trip"))
	'			if Not IsNull(rstf("Cod_esta_trip")) then
	'			rste.Open "Select * From Siep_esta Where Cod_esta = " & rstf("Cod_esta_trip"), cnnf, 1, 2
	'			if rste.EOF then
	'			
	'				
	'			
	'			
	'				Estado_trip = ""
	'			else
	'				Estado_trip = rste("Des_esta")
	'			End if
	'			rste.close
	'			else
	'				Estado_trip=""  
	'			end if
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_29", Replace(Estado_trip, "Ã‘", "N")) & hidden("selct_fila" & CStr(ind) & "_30", Estado_trip)						
	'	        if Not IsNull(rstf("Cod_ciu_trip")) then
	'		       rstc.Open "Select * From ciudad Where Cod_ciuda = " & rstf("Cod_ciu_trip"), cnnf, 1, 2
	'			   if not rstc.eof then 
	'		         ciudad_trip = rstc("nom_ciuda")
	'			   else 
    '                 ciudad_trip = ""					 
	'			   end if 	 
	'		       rstc.close
	'		    else
	'			
	'			
	'			
	'			   ciudad_trip=""  
	'		    end if
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_31", Replace(ciudad_trip, "Ã‘", "N")) & hidden("selct_fila" & CStr(ind) & "_32", ciudad_trip)			
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_33", rstf("Tip_visa")) & hidden("selct_fila" & CStr(ind) & "_34", rstf("Tip_visa"))				
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_35", rstf("Fec_exp_visa")) & hidden("selct_fila" & CStr(ind) & "_36", rstf("Fec_exp_visa"))				
	'	     hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_37", rstf("Ciu_exp_visa")) & hidden("selct_fila" & CStr(ind) & "_38", rstf("Ciu_exp_visa"))
	'		 
	'		 if rstf("sex_trip") = true then
	'			si_nos = "F"
	'		else
	'			si_nos = "M"
	'		End if
	'		 
	'		 
	'		 
	'		 
	'		 
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_39", si_nos) & hidden("selct_fila" & CStr(ind) & "_40", si_nos)						
	'		
	'		'///////////////////////mas Nuevo/////////////////////////
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_43", rstf("Ima_cer_tri")) & hidden("selct_fila" & CStr(ind) & "_44", rstf("Ima_cer_tri"))	
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_45", rstf("Ima_cer_eua_tri")) & hidden("selct_fila" & CStr(ind) & "_46", rstf("Ima_cer_eua_tri"))
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_47", rstf("Ima_lic_tri")) & hidden("selct_fila" & CStr(ind) & "_48", rstf("Ima_lic_tri"))	
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_49", rstf("Ima_lic_eua_tri")) & hidden("selct_fila" & CStr(ind) & "_50", rstf("Ima_lic_eua_tri"))
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_51", rstf("Ima_pasa")) & hidden("selct_fila" & CStr(ind) & "_52", rstf("Ima_pasa"))
	'		 hid_tab = hid_tab & hidden("selct_fila" & CStr(ind) & "_53", rstf("Ima_visa")) & hidden("selct_fila" & CStr(ind) & "_54", rstf("Ima_visa"))	
    '
	'		if rstf("Sta_acti") = true then
	'			si_no = "Si"
	'		else
	'			si_no = "No"
	'		End if	
	'		 
	'		tabla = tabla & "<td bgcolor='#e6e8e9'><div align='center'>" & si_no & " " & hidden("selct_fila" & CStr(ind) & "_24", si_no) & " " & hid_tab &  "</div></td></tr>"
	'		rstf.MoveNext
	'		ind = ind + 1
	'	Loop
	'End if
	'tabla = tabla & "</table>"
	htmlbus = img("../Imagenes/boton_aceptar.jpg", "70", "19", "", "", "")
	tabla = tabla & "<div align='left'>Filtro busqueda: " & campo("busquedaPas", "text", "bus_pas", "20", "20", "8","", "onkeypress='busPas()'") & enlace("", "javascript:void(0)", "text-decoration:none", "", htmlbus, "OnClick='submit()'") &"</div></br></br>"	
	tabla = tabla & "<table id='table_id' border='0' class='display' cellspacing='0' width='100%'>"
    tabla = tabla & "    <thead>"
    tabla = tabla & "        <tr>"
    tabla = tabla & "            <th>Nombre</th>"
	tabla = tabla & "			 <th>Apellido</th>"
    tabla = tabla & "            <th>C&eacute;dula</th>"
    tabla = tabla & "            <th>Tel&eacute;fono</th>"
    tabla = tabla & "            <th>Email</th>"
    tabla = tabla & "            <th>Direcci&oacute;n</th>"
    tabla = tabla & "            <th>Nro. Pasaporte</th>"
    tabla = tabla & "            <th>Venc. Pasaporte</th>"
    tabla = tabla & "            <th>Imagen Pasaporte</th>"
    tabla = tabla & "            <th>Nro. Visa EUA</th>"
 	tabla = tabla & "            <th>Venc. Visa EUA</th>"
 	tabla = tabla & "            <th>Imagen Visa EUA</th>"
 	tabla = tabla & "            <th>Nro. Licencia</th>"
 	tabla = tabla & "            <th>Venc. Licencia</th>"
 	tabla = tabla & "            <th>Imagen Licencia</th>"
 	tabla = tabla & "            <th>Nro. Lic. EUA</th>"
 	tabla = tabla & "            <th>Imagen Lic. EUA</th>"
 	tabla = tabla & "            <th>Nro. Cert.</th>"
 	tabla = tabla & "            <th>Venc. Cert.</th>"
 	tabla = tabla & "            <th>Imagen Cert.</th>"
 	tabla = tabla & "            <th>Nro. Cert. EUA</th>"
 	tabla = tabla & "            <th>Venc. Cert. EUA</th>"
 	tabla = tabla & "            <th>Imagen Cert. EUA</th>"
 	tabla = tabla & "            <th>Activo</th>"
 	tabla = tabla & "			 <th style='visibility:hidden'><span></span></th>"
 	tabla = tabla & "            </tr></thead>"
    
    'rstf.Open "Select t.*, p.Nom_pais From Siep_trip t  inner join Siep_pais p on t.Nac_tri = p.Cod_pais Where (t.Num_clie = " & Session("Num_clie") & ") Order By Nom_trip", cnnf, 1, 2
    'rstf.Open "Select t.*, p.Nom_pais From Siep_trip t  inner join Siep_pais p on t.Nac_tri = p.Cod_pais Where (t.Nom_trip like '%" & filtrobusqueda & "%') Order By Nom_trip", cnnf, 1, 2
	
    
	if filtrobusqueda <> "" then 
		     rstf.Open "Select t.*, p.Nom_pais From Siep_trip t  inner join Siep_pais p on t.Nac_tri = p.Cod_pais Where (t.Nom_trip like '%" & filtrobusqueda & "%') Order By Nom_trip", cnnf, 1, 2
	else
		     rstf.Open "Select s.* From Siep_trip s Where  s.Num_trip = 0 Order By s.Nom_trip", cnnf, 1, 2	 
	end if
	
	
	
	
	
	
	
	
	
	
	hid_tab = ""
    if  not rstf.EOF then
    	tabla = tabla & "    <tbody>"
    	Do Until rstf.EOF	
    		tabla = tabla & "        <tr>"
    		tabla = tabla & "            <td>" & rstf("Nom_trip") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Ape_trip") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Num_cedu_tri") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Tlf_trip_mov") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Ema_trip") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Dir_trip") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Num_pasa_tri") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Fec_venc_pas") & "</td>"
    		tabla = tabla & "            <td>" & null_cont(rstf("Ima_pasa"), button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=3&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Num_visa_tri") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Fec_venc_vis") & "</td>"
    		tabla = tabla & "            <td>" & null_cont(rstf("Ima_visa"), button("visa" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=4&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Num_lic_tri") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Fec_venc_lic") & "</td>"
    		tabla = tabla & "            <td>" & null_cont(rstf("Ima_lic_tri"), button("licencia" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=5&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
    		tabla = tabla & "            <td>" & rstf("Num_lic_eua_tri") & "</td>"
    		tabla = tabla & "            <td>" & null_cont(rstf("Ima_lic_eua_tri"), button("licencia_eua" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=6&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
			tabla = tabla & "			 <td>" & rstf("Num_cer_tri") & "</td>"
			tabla = tabla & "            <td>" & rstf("Fec_venc_cer") & "</td>"
			tabla = tabla & "            <td>" & null_cont(rstf("Ima_cer_tri"), button("certificado" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=7&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
    		tabla = tabla & "			 <td>" & rstf("Num_cer_eua_tri") & "</td>"
			if(IsNull(rstf("Fec_exp_eua_cer"))) then
				tabla = tabla & "<td>&nbsp;</td>"
			else
				tabla = tabla & "<td>" & vencCerEUA(rstf("Fec_exp_eua_cer")) & "</td>"	
			end if
			tabla = tabla & "			 <td>" & null_cont(rstf("Ima_cer_eua_tri"), button("certificado" & CStr(ind), "Ver", "OnClick='MM_openBrWindow(""obtener.asp?j=8&n=" & rstf("Num_trip") & """, ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'"),  "N/D") & "</td>"
			
			if rstf("Sta_acti") = true then
				si_no = "Si"
			else
				si_no = "No"
			End if	
			 
			tabla = tabla & "<td>" & si_no & "</td>"
			
			if Not IsNull(rstf("Cod_pais_resd")) or Trim(rstf("Cod_pais_resd")) <> "" then
				rstr.Open "Select * From Siep_pais Where Cod_pais = " & rstf("Cod_pais_resd"), cnnf, 1, 2
			  	if rstr.EOF then
			  		Pais_res_trip=""
			  	else
			  		Pais_res_trip = rstr("Nom_pais")
			  	End if
			  	rstr.close
			else
		    	Pais_res_trip=""  
			end if
			
			if Not IsNull(rstf("Cod_esta_trip")) then
				rste.Open "Select * From Siep_esta Where Cod_esta = " & rstf("Cod_esta_trip"), cnnf, 1, 2
				if rste.EOF then
					Estado_trip = ""
				else
					Estado_trip = rste("Des_esta")
				End if
				rste.close
			else
				Estado_trip=""  
			end if
			
			if Not IsNull(rstf("Cod_ciu_trip")) then
				rstc.Open "Select * From ciudad Where Cod_ciuda = " & rstf("Cod_ciu_trip"), cnnf, 1, 2
					if not rstc.eof then 
						ciudad_trip = rstc("nom_ciuda")
					else 
						ciudad_trip = ""					 
					end if 	 
				rstc.close
			else
				ciudad_trip=""  
			end if
	
			hid_tab = rstf("Sex_trip")& "|" & Replace(rstf("Nom_pais"), "Ã‘", "N") & "|" & rstf("Fec_nac_tri") & "|" & Replace(Pais_res_trip, "Ã‘", "N")
			hid_tab = hid_tab & "|" & Replace(Estado_trip, "Ã‘", "N")  & "|" & Replace(ciudad_trip, "Ã‘", "N")
			hid_tab = hid_tab & "|" & rstf("Tip_visa") & "|" & rstf("Fec_exp_visa") & "|" & rstf("Ciu_exp_visa") & "|" & rstf("Num_trip")
			
			hid_tab = hid_tab & "|" & null_bd(rstf("Ima_pasa"), "Nulo")  & "|" & null_bd(rstf("Ima_visa"), "Nulo")
			hid_tab = hid_tab & "|" & null_bd(rstf("Ima_lic_tri"), "Nulo") & "|" & null_bd(rstf("Ima_lic_eua_tri"), "Nulo")
			hid_tab = hid_tab & "|" & null_bd(rstf("Ima_cer_tri"), "Nulo") & "|" & null_bd(rstf("Ima_cer_eua_tri"), "Nulo")
			
			tabla = tabla & "            <td style='visibility:hidden'>" & hid_tab & "</td>"
			tabla = tabla & "        </tr>"
    		rstf.MoveNext
    	Loop
    	tabla = tabla & "</tbody>"
    End if
    tabla = tabla & "</table>"
	
	
	thidden = thidden & hidden("lista", ind)
	
	rstf.Close
	cnnf.Close
	Set rstf = nothing
	Set cnnf = nothing
	
	Set load = nothing
	
	'<img src="../imagenes/boton_editar.jpg" width="70" height="19" />'
	html2 = img("../Imagenes/boton_editar.jpg", "70", "19", "", "", "")
	'boton_editar = enlace("", "#", "text-decoration:none", "", html2,  "OnClick='llenaCampEdit(document.form1.lista.value)'")
	boton_editar = ""
	
	htrans = img("../imagenes/botontransferir.gif", "125", "28", "", "", "")
	transferir = enlace("", "transferencia_pasaj.asp", "text-decoration:none", "", htrans, "")
	
	ht = img("../imagenes/botonsalir.gif", "125", "28", "", "", "")
	salir = enlace("", "../salir.asp", "text-decoration:none", "", ht, "")
	
	if Err.Number <> 0 then
		Response.Write Err.Description & " " & Err.Number & " Los archivos no deben pesar mas de 100 Kb cada uno" 
		Error.Clear
		Response.End()
	End if
%>
<!--#include file="../template/tripulacion.html" -->

