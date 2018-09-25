<% @ LCID = 1034 %>


<!--'#include file="../modulos/vali_sesion.asp" -->



<!--#include file="../modulos/funcion.asp" -->
<!--#include file="../modulos/control.asp" -->
<!--#include file="../modulos/loader.asp"-->
<!--#include file="../js/funcion.js" -->
<!--#include file="../js/autosuggest.js" -->
<!--#include file="../js/datetimepicker_css.js" -->
<%
	On Error Resume Next		
%>

<%
'CONEXION ADO -- HUMBERTO ROJAS
'Option Explicit

'Se declaran las variables
Dim Conexion
Dim Cadena
Dim Rutafisica
Dim ADOPersonas
Dim MiId
Dim MiNombre

'Se crean dos objetos, una conexión y un recordset
Set Conexion = CreateObject("ADODB.Connection")
Set ADO = CreateObject("ADODB.Recordset")

'Se abre la conexión
Conexion.Open dsn()

'Se ejecuta la sentencia SQL
'ADO.Open "Select Id, Nombre from Personas", Conexion

%>	
	
	
<%	




	Call lista_arch_js2("siep\js\list_autogest_pais_cod.js", "pais_cod", "Cod_pais", "Nom_pais", "Siep_pais", "")
	Call lista_arch_js("siep\js\list_autogest_pais.js", "pais", "Nom_pais", "Siep_pais", "")
	Call lista_arch_js2("siep\js\list_autogest_pais_cod.js", "pais_cod", "Cod_pais", "Nom_pais", "Siep_pais", "")
	
	
	buscar = Request("b")
	viene_list = Request("viene_lista")
	
	
	Response.Buffer = true
	Dim load
	Set load = new Loader
	load.initialize
	
	cortesia_bd =  selectSimpleSQL2("Cor_susc", "Siep_susc", "Num_clie = " & Session("Num_clie"), 1)
	if cortesia_bd = true then
		cortesia = " (Suscripci&oacute;n de cortes&iacute;a)"
	else
		cortesia = ""
	End if 
	fuente = "<font color = #00ade5 ><b>Hola. Bienvenido " & Session("Nom_clie") & " " & cortesia & "</b></font>"
	html2 = img("../imagenes/botonsalir.gif", "125", "28", "", "", "")
	salir = enlace("", "../salir.asp", "border-style:none", "", html2, "")
	Nom_cli = tag_class("label", "", "", "", fuente, "")
	Dim ind
	z = load.getValue("aleat")
	
	'Insertar, borrar o editar'
	if val_inser = "" then
		val_inser = "0"
	End if
	
	busc_clie = load.getValue("camp_busc")
	ed_in_bor = load.getValue("edit_inser_borr")
	'ed_in_bor = 1
	nume_pasa = load.getValue("num_pasa")
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
	'guard = 1
	
	nom_aper = ""
	nom_ape2r = ""
	apellidor = ""
	apellido2r = ""
	cedular = ""
	cedula2r = ""
	fechnacr = ""
	fec_venc_visa_extrr = ""
	tel_celr = ""
	e_mailr = ""
	e_mail2r = ""
	direccr = ""
	tipo_visar = ""
	ocupar = ""
	ciu_expr = ""
	tipo_visar = ""
	ciudadr = ""
	nac_pasajr = ""
	cod_nac_pasajr = ""
	edo_vzla_pasajr = ""
	cod_edo_vzla_pasajr = ""
	res_pasajr = ""
	cod_res_pasajr = ""
	cod_ciudadr = ""
	num_pasapr = ""
	num_pasap2r = ""
	fec_vencr = ""
	img_pasapr = ""
	num_visa_euar = ""
	num_visa_eua2r = ""
	fec_venc_visa_euar = ""
	img_visa_euar = ""
	activor = "on"	
	sexor=""
	sexorF=""
	remitidor = ""
	ced_menorr = ""
	Pais_res_pas=""
	Estado=""
	
	
	'Set cnnLista = Server.CreateObject("ADODB.Connection")
	'Set rstLista = Server.CreateObject("ADODB.Recordset")
	'cnnLista.Open dsn()
	
	
	'Agregado por: Humberto Rojas
	'En caso de que venga de afuera
	if request.QueryString("vienelista") = 1 then
	'VL = 1
	'if VL = 1 then
	
	cedula = request.QueryString("cedula")		
    	buscar = request.QueryString("b")
	err_nums =  0
	edit_inser_borr = 1
	ed_in_bor = edit_inser_borr
	viene_lista = request.QueryString("vienelista")
	
	 
	'rstLista.Open "Select * From Siep_pasa Where Num_cedu_pas  like '%" & cedula & "%' and Num_clie = " & Session("Num_clie") & " order by nom_pasa desc ", cnnLista, 1, 2
	'//////////////Revisar---------------------
	ADO.Open "Select * From Siep_pasa Where Num_cedu_pas  like '%" & cedula & "%' and Num_clie = " & Session("Num_clie") & " order by nom_pasa desc ", Conexion, 1, 2
	
	if not ADO.eof then		   
		   Call alerta("Select * From Siep_pasa Where Num_cedu_pas  like '%" & cedula & "%' and Num_clie = " & Session("Num_clie") & " order by nom_pasa desc ")
		   cedular = cedula
		   nume_pasa = ADO(2)
		   		   
		   nom_aper = ADO(6)
		   
		   
		   apellidor = ADO(7)
		   sexo = ADO(26)
		   fechnacr = ADO(4)
		   nac_pasajr = selectSimpleSQL2("nom_pais", "Siep_pais", "Cod_pais = " & ADO(5) , 1)	
		   
		   
		   res_pasajr =  selectSimpleSQL2("nom_pais", "Siep_pais", "Cod_pais = " & ADO(28) , 1)	
			
		   edo_vzla_pasajr = selectSimpleSQL2("des_esta", "Siep_esta", "Cod_pais = " & ADO(28) & " and Cod_esta = " & ADO(29)   , 1)   
			
		   'response.write "edo " & edo_vzla_pasajr
		   if nac_pasajr = "EOF" then
		      nac_pasajr = ""
		   end if
		   
		   if res_pasajr = "EOF" then
		      res_pasajr = ""
		   end if
		   
		   if edo_vzla_pasajr = "EOF" then
		      edo_vzla_pasajr = ""
		   end if		   
		   
		   if edo_vzla_pasajr = "EOF" then
		      edo_vzla_pasajr = ""
		   end if	  	   
		   
		   tel_celr = ADO(9)
           e_mailr = ADO(10)
           direccr = ADO(8)
		   ocupar = ADO(27)
		   
		   ciu_expr = ADO(32)
		   tipo_visar = ADO(22)
		   ciudadr = selectSimpleSQL2("Nom_ciuda", "Ciudad", "cod_ciuda = " & ADO(11), 1)
		   
		   if ciudadr = "EOF" then
		      ciudadr = ""
		   end if	  	   
		   
		   num_pasapr = ADO(12)
		   fec_vencr = ADO(14)
		   img_pasapr = ADO(16)
		   num_visa_euar = ADO(13)
		   fec_venc_visa_euar = ADO(15)
		   img_visa_euar = ADO(17)		   
			
			
	val_inser = 1
	vienelista = hidden("vienelista",1)
	
	
	end if 
	
	ADO.Close
    'Conexion.Close	
	
	end if
	'Fin agregado por Humberto Rojas
	
	
								
				
	
	
	
	
	
	'Set cnnf = Server.CreateObject("ADODB.Connection")
	'Set rstf = Server.CreateObject("ADODB.Recordset")
	'Set rstw = Server.CreateObject("ADODB.Recordset")
	'Set rstr = Server.CreateObject("ADODB.Recordset")
	'Set rste = Server.CreateObject("ADODB.Recordset")
	'Set rstn = Server.CreateObject("ADODB.Recordset")
	'cnnf.Open dsn()
	
	if request.QueryString("vienelista") = "9999" then
	 'if request.QueryString("b") = 0 then
	   response.Write "buscar " & buscar
	   response.write " guard " & guard
	   response.write " refresco  " & refresco
	   response.write " err_nums  " & err_nums
	   response.write " ed_in_bor  " & ed_in_bor
	   response.end()
	   
	 end if
	
	if buscar = "0" then
		if guard = "1" then
			'Se verifica que se hizo un submit en vez de un'
			'Refresh'
			if refresco = false then
				Call crear_carpeta(Session("Nom_clie") & "_" & Session("Num_clie"))
				err_nums =  0
				nom_aper = limpia(load.getValue("nom_ape"))
				nom_ape2r = limpia(load.getValue("nom_ape2"))
				apellidor = limpia(load.getValue("apellido"))
				apellido2r = limpia(load.getValue("apellido2"))
				cedular = limpia(load.getValue("cedula"))
				cedula2r = limpia(load.getValue("cedula2"))
				ced_menorr = limpia(load.getValue("ced_menor"))
				fechnacr = limpia(load.getValue("fechnac"))
				fec_venc_visa_extrr = limpia(load.getValue("fec_venc_visa_extr"))
				tel_celr = limpia(load.getValue("tel_cel")) 
				e_mailr = limpia(load.getValue(LCase(("e_mail"))))
				e_mail2r = limpia(load.getValue(LCase(("e_mail2"))))
				direccr = limpia(load.getValue("direcc"))
				tipo_visar = limpia(load.getValue("tipo_visa"))
				ocupar = limpia(load.getValue("ocupa")) 
				ciu_expr = limpia(load.getValue("ciu_exp"))  
				ciudadr = limpia(load.getValue("ciudad"))  
				nac_pasajr = limpia(load.getValue("nac_pasaj"))
				edo_vzla_pasajr = limpia(load.getValue("edo_vzla_pasaj"))
				cod_nac_pasajr = limpia(load.getValue("cod_nac_pasaj"))
				res_pasajr = limpia(load.getValue("res_pasaj"))
				cod_res_pasajr = limpia(load.getValue("cod_res_pasaj"))
				

				
				if (CInt(ed_in_bor) <> 2) And (nac_pasajr <> "") And (cod_nac_pasajr <> "") then
				
				
						
			
					if (edo_vzla_pasajr <> "") then
				
						'rstf.Open  "Select Cod_esta From Siep_esta Where (Cod_pais = " & cod_nac_pasajr & ") And  (Des_esta = '" & edo_vzla_pasajr & "')", cnnf, 1, 2
						ADO.Open "Select Cod_esta From Siep_esta Where (Cod_pais = " & cod_nac_pasajr & ") And  (Des_esta = '" & edo_vzla_pasajr & "')", Conexion, 1, 2
						if ADO.EOF then	
							Cod_esta = ""
						else
							Cod_esta = ADO("Cod_esta")
						End if
						ADO.Close
                        'Conexion.Close					
	
						if Cod_esta = "" then				
							'rstf.Open  "Siep_esta", cnnf, 1, 2
							ADO.Open "Siep_esta", Conexion, 1, 2
							ADO.AddNew
							'rstf.AddNew
							ADO("Cod_pais") = cod_nac_pasajr
							ADO("Des_esta") = edo_vzla_pasajr
			
							ADO.Update 							
							ADO.Close
                            'Conexion.Close
				
							
							ADO.Open  "Select Last(Cod_esta) As codE From Siep_esta", Conexion, 1, 2
							'rstf.Open  "Select Last(Cod_esta) As codE From Siep_esta", cnnf, 1, 2
							
							 if not ADO.EOF then	
							   Cod_esta = ""							
							   Cod_esta = ADO("codE")
							   ADO.Close
							 end if  
						End if
					End if
					
							

			
					if (ciudadr <> "") then
						'rstf.Open  "Select Cod_ciuda From Ciudad Where (Cod_pais = " & cod_nac_pasajr & ") And  (Nom_ciuda = '" & ciudadr & "')", cnnf, 1, 2
						ADO.Open  "Select Cod_ciuda From Ciudad Where (Cod_pais = " & cod_nac_pasajr & ") And  (Nom_ciuda = '" & ciudadr & "')", Conexion, 1, 2
					
						if ADO.EOF then	
							Cod_ciu = ""
						
						else
							Cod_ciu = ADO("Cod_ciuda")
						End if
						
				
					
						ADO.Close '********************************************ERROR (22-06-2015) - Operation is not allowed when the object is closed. 3704
						if Cod_ciu = "" then
							ADO.Open  "Select Max(Cod_ciuda) As CodMax From Ciudad", Conexion, 1, 2
							if not ADO.eof then
							   CodMax = CInt(ADO("CodMax")) + 1 '-> El código no es autonumérico
							   ADO.Close
							end if
							'/////??????????->Comentado por irving Houli 17-09-2015
							'response.write "CODIGO CIUDAD" & CodMax
							'response.end()								
							ADO.Open  "Ciudad", Conexion
							ADO.AddNew
							ADO("Cod_ciuda") = CStr(CodMax)
							ADO("Cod_pais") = cod_nac_pasajr							
							ADO("Nom_ciuda") = ciudadr
							ADO.Update
							ADO.Close							
							Cod_ciu = CodMax
							
						End if
					End if
				End if
				
		

				cod_edo_vzla_pasajr = Cod_esta 'limpia(load.getValue("cod_edo_vzla_pasaj"))				
				cod_ciudadr = Cod_ciu 'limpia(load.getValue("cod_ciud"))
				num_pasapr = limpia(load.getValue("num_pasap"))
				num_pasap2r = limpia(load.getValue("num_pasap2"))
				fec_vencr = limpia(load.getValue("fec_venc"))
				img_pasapr = limpia(load.getValue("img_pasap"))
				num_visa_euar = limpia(load.getValue("num_visa_eua"))
				num_visa_eua2r = limpia(load.getValue("num_visa_eua2"))
				fec_venc_visa_euar = limpia(load.getValue("fec_venc_visa_eua"))
				img_visa_euar = limpia(load.getValue("img_visa_eua"))
				activor =  load.getValue("activo")						
				sexor =  load.getValue("sexo")
				sexorF =  load.getValue("sexo2")
				remitidor = load.getValue("remitido")
				'///////////////////////////////////'
				'Solo ciudad y nombre no pueden ser null'
				'No se admiten pasaporte, cédula, visa y/o correo repetidos '
				'para un mismo cliente si no son null'
				'////////////////////////////////////'
			
				if tel_celr = "" then
					tel_celr = null
				End if
				
				
		
				if cedular <> "" And (cedular <> cedula2r) then
					if Tabla_Vacia("Siep_pasa", " (Num_cedu_pas = '" & cedular & "') And (Num_pasa_orig is Null) And (Num_clie = " & Session("Num_clie") & ")") = false then
						call alerta("La cédula " & cedular & " se encuentra registrada ")
						err_nums =  err_nums + 1
					End if
				End if
				if fechnacr="" then
					fechnacr=null
				end if
				if cedular = "" And ced_menorr <> "on" then
					cedular = null
				End if
				if ced_menorr = "on"  then
					cedular = "V" & Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now()) & "_Menor"
				End if
				if e_mailr <> "" And (e_mailr <> e_mail2r) then
					if Tabla_Vacia("Siep_pasa", " Ema_pasa = '" & e_mailr & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El correo " & e_mailr & " se encuentra registrado ")
						err_nums =  err_nums + 1
					End if
				End if
				if e_mailr = "" then
					e_mailr = null
				End if
				if num_pasapr <> "" And (num_pasapr <> num_pasap2r) then
					if Tabla_Vacia("Siep_pasa", " Num_pasa_pas = '" & num_pasapr & "' And Num_clie = " & Session("Num_clie")) = false then
						call alerta("El pasaporte " & num_pasapr & " se encuentra registrado ")
						err_nums =  err_nums + 1
					End if
				End if
				if fec_venc_visa_extrr="" then
				    fec_venc_visa_extrr=null
				end if
				if e_mailr = "" And CInt(ed_in_bor) = 0 then
					num_pasapr = null
				End if
				if num_visa_euar <> "" And (num_visa_euar <> num_visa_eua2r) then
					if Tabla_Vacia("Siep_pasa", " Num_visa_pas = '" & num_visa_euar & "' And Num_clie = " & Session("Num_clie")) = false then
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
				if ciu_expr="" then
					ciu_expr=null
				end if
				if sexor <> "0" and sexor <>"1" then
					sexor ="0"
				end if

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
					if fileSize > 1000000 then
						Call alerta("No se puede subir un archivo de mas de 1 MB")
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
						pathToFile = CStr(Server.mapPath(carpeta) & "\" & fileName)
						Dim fileUploaded
						fileUploaded = load.saveToFile (Session("Nom_clie") & "_" & Session("Num_clie") & "_" & now(), pathToFile)
						'img_pasapr = pathToFile
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
					if fileSize2 > 1000000 then
						Call alerta("No se puede subir un archivo de mas de 1 MB")
						img_visa_euar = "" 
						err_nums =  err_nums + 1
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
						pathToFile2 = CStr(Server.mapPath(carpeta2) & "\" & fileName2)
						Dim fileUploaded2
						fileUploaded2 = load.saveToFile (Session("Nom_clie") & "_" & Session("Num_clie") & "_" & now(), pathToFile2)'
						'img_visa_euar = pathToFile2
					End if
				End if
				if err_nums = 0 then
				
				   'response.write "ed_in_bor " & ed_in_bor				
					Select case CInt(ed_in_bor)
						case 0, 1
					
							if CInt(ed_in_bor) = 0 then							   
								ADO.Open  "Siep_pasa", Conexion
								ADO.AddNew
								ADO("Fec_clie_reg") = CStr(Session("Fec_clie_reg"))
								ADO("Num_clie") = CStr(Session("Num_clie"))
							Else
								ADO.Open  "Select * From Siep_pasa Where Num_pasa = " & nume_pasa, Conexion, 1, 2
								'rstf.Open  "Select * From Siep_pasa Where Num_pasa like '%29318244%'", cnnf, 1, 2
								
								if not ADO.eof then
								   val_inser = "0"
								else
								    val_inser = ""   
								end if   
							End if
							    ADO("Nom_pasa") = CStr(nom_aper)
							    ADO("Ape_pasa") = CStr(apellidor)
								if(cod_ciudadr <> "") then
									ADO("Cod_ciud") = CStr(cod_ciudadr)
								else
									ADO("Cod_ciud") = null
								End if
							    ADO("Num_cedu_pas") = cedular
							    ADO("Fec_nac_pas") = fechnacr							
							    ADO("Fec_venc_visa_pas") = fec_venc_visa_extrr
								if(cod_nac_pasajr <> "") then
									ADO("Nac_pasaj") = cod_nac_pasajr
								else
									ADO("Nac_pasaj") = null
								End if
								if(cod_edo_vzla_pasajr <> "") then
									ADO("Est_dir_pas") = cod_edo_vzla_pasajr
								else
									ADO("Est_dir_pas") = null
								end if
								if(cod_res_pasajr <> "") then
									ADO("Pais_res_pas") = cod_res_pasajr
								else
									ADO("Pais_res_pas") = null
								end if
								ADO("Dir_pasa") = direccr
								ADO("Tipo_vis_pas") = tipo_visar
								ADO("ocu_pas") = ocupar
								ADO("Ciud_exp_vis_pas") = ciu_expr
								ADO("Tlf_movi_pas") = tel_celr
								ADO("Ema_pasa") = LCase(e_mailr)
								ADO("Num_pasa_pas") = num_pasapr
								ADO("Num_visa_pas") = num_visa_euar
								ADO("Fec_venc_pas") = fec_vencr
								ADO("Fec_venc_vis") = fec_venc_visa_euar
								if Not IsNull(img_pasapr) And img_pasapr <> "" then
									ADO("Ima_pasa") = img_pasapr
									ADO("Bin_pas").AppendChunk fileData
									ADO("Tip_cont_pas") = contentType
								End if
								if Not IsNull(img_visa_euar) And img_visa_euar <> "" then
									ADO("Ima_visa") = img_visa_euar
									ADO("Bin_visa").AppendChunk fileData2
									ADO("Tip_cont_vis") = contentType2
								End if
								if activor = "on" then
									ADO("Sta_acti") = "1"
								else
									ADO("Sta_acti") = "0"
								End if							
								ADO("Sex_pas") = sexor
								ADO("Rem_pas") = remitidor
													
								ADO.Update 
								ADO.Close
								
						case 2
						
							'nums = load.GetValue("lista")
							msj_err = 0
							msj = "El o los pasajero(s) "
							'No hay una lista de checkboxes -> Irving 08-2015
							'No hace falta el for
							'for fila = 1 to nums-1
								num_pas = "selct_fila" & CStr(fila) & "_0"
								selec = "selct_fila" & CStr(fila)
								'if load.GetValue(selec) = "on" then
									'if Tabla_Vacia("Siep_mapa",  "Num_pasa  = " & load.GetValue(num_pas)) = true then
									if Tabla_Vacia("Siep_mapa",  "Num_pasa  = " & nume_pasa) = true then
										'rstf.Open "Delete From Siep_pasa Where Num_pasa  = " & load.GetValue(num_pas), cnnf, 1, 2
										ADO.Open "Delete From Siep_pasa Where Num_pasa  = " & nume_pasa, Conexion, 1, 2
									else
										msj_err = msj_err + 1
										'msj = msj & load.GetValue("selct_fila" & fila & "_1") & ", "
										msj = msj & CStr(nom_aper) & " " & CStr(apellidor) & " "
									End if
								'End if
							'next
							if msj_err > 0 then
								'msj = left(msj, len(msj)-2)
								msj = msj & " no puede ser eliminado. Solo se puede colocar en estatus inactivo"
								Call alerta(msj)
							End if
							val_inser = "0"
					End Select
					
					ced_menorr = ""
					nume_pasa = ""
					nom_aper = ""
					nom_ape2r = ""
					apellidor = ""
					apellido2r = ""
					cedular = ""
					fechnacr = ""
					fec_venc_visa_extrr = ""
					cedula2r = ""
					tel_celr = ""
					e_mailr = ""
					e_mail2r = ""
					direccr = ""
					tipo_visar = ""
					ocupar = ""
					ciu_expr = ""
					ciudadr = ""
					nac_pasajr = ""
					cod_nac_pasajr = ""
					edo_vzla_pasajr = ""
					cod_edo_vzla_pasajr = ""
					res_pasajr = ""
					cod_res_pasajr = ""
					num_pasapr = ""
					num_pasap2r = ""
					fec_vencr = ""
					img_pasapr = ""
					num_visa_euar = ""
					num_visa_eua2r = ""
					fec_venc_visa_euar = ""
					img_visa_euar = ""
					activor = "on"
					sexor = ""
					sexorF= ""
					remitidor = ""
				Else
					'if CInt(ed_in_bor) = 1 then
						'm = load.getValue("lista")
						'Redim lista_check(m-1)
						'For j = 1 to m-1
						'	if load.getValue("selct_fila" & j) = "on" then
						'		lista_check(j) = "checked"
						'	Else
						'		lista_check(j) = ""
						'	End if
						'Next
					'End if
					val_inser = ed_in_bor
				End if
			End if
		End if
	End if
	
	
	
		
					
	
	
	
	imgcal = img("../images_cal/cal.gif", "16", "16", "Pick a date", "", "")
	num_pasa = hidden("num_pasa", nume_pasa)
	num_pasa = num_pasa & hidden("viene_lista", viene_lista)
	'<input name="nom_ape" type="text" id="nom_ape" size="20" maxlength="20"/>'
	nom_ape = campo("nom_ape", "text", "nom_ape", "30", "50", "1", nom_aper, "OnBlur='caractNoPermit(this)'")
	nom_ape = nom_ape & hidden("nom_ape2", nom_ape2r)
	ast_nom_ape = tag("span", "ast_nom_ape", "visibility:hidden", marca_error(), "")
	
	'<input name="apellido" type="text" id="apellido" size="20" maxlength="20"/>'
	apellido = campo("apellido", "text", "apellido", "30", "50", "1", apellidor, "OnBlur='caractNoPermit(this)'")
	apellido = apellido & hidden("apellido2", apellido2r)
	ast_apellido = tag("span", "ast_apellido", "visibility:hidden", marca_error(), "")
	sexo = radio("sexo", "sexo", "0", "") 
	sexo2 = radio("sexo", "sexo", "1",  "")


	
	'<input name="cedula" type="text" id="cedula" size="20" maxlength="20"/>'
	'cedula = campo("cedula", "text", "cedula", "15", "15", "2", cedular, "OnBlur='verCedula(this)'")
	cedula = campo("cedula", "text", "cedula", "30", "15", "2", cedular, "")
	cedula = cedula & hidden("cedula2", cedula2r)  & "</br>" & check_box("ced_menor", "ced_menor", "", "OnClick='deshabHabilitaCed()'")
	ast_cedula = tag("span", "ast_cedula", "visibility:hidden", marca_error(), "")
	
	'<input name="cedula" type="text" id="cedula" size="20" maxlength="20"/>'
	fechnac = campo_readonly("fechnac", "text", "fechnac", "30", "20", "3", fechnacr, "readonly", "")
	cal = enlace("", "javascript:NewCssCal(""fechnac"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fechnac = tag("span", "ast_fechnac", "visibility:hidden", marca_error(), "")
	
	
	fec_venc_visa_extr = campo_readonly("fec_venc_visa_extr", "text", " fec_venc_visa_extr", "30", "20", "3",  fec_venc_visa_extrr, "readonly", "")
	cal5 = enlace("", "javascript:NewCssCal("" fec_venc_visa_extr"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_visa_extr = tag("span", "ast_ fec_venc_visa_extr", "visibility:hidden", marca_error(), "")
	
	nac_pasaj = campo("nac_pasaj", "text", "nac_pasaj", "30", "15", "4", nac_pasajr, "OnBlur='caractNoPermit(this)'") & hidden("cod_nac_pasaj", "")
	ast_nac_pasaj = tag("span", "ast_nac_pasaj", "visibility:hidden", marca_error(), "")

	edo_vzla_pasaj = campo("edo_vzla_pasaj", "text", "edo_vzla_pasaj", "30", "40", "4", edo_vzla_pasajr, "OnBlur='caractNoPermit(this)'") & 	hidden("cod_edo_vzla_pasaj", "")
	ast_edo_vzla_pasaj = tag("span", "ast_edo_vzla_pasaj", "visibility:hidden", marca_error(), "")
	
	
	'response.Write res_pasaj

	res_pasaj = campo("res_pasaj", "text", "res_pasaj", "30", "15", "4", res_pasajr, "OnBlur='caractNoPermit(this)'") & hidden("cod_res_pasaj", "")
	ast_res_pasaj = tag("span", "ast_res_pasaj", "visibility:hidden", marca_error(), "")
	
	'<input name="tel_cel" type="text" id="tel_cel" size="50" maxlength="20"/>'
	tel_cel = campo("tel_cel", "text", "tel_cel", "30", "15", "5", tel_celr, "OnBlur='esNum(this)'")
	ast_tel_cel = tag("span", "ast_tel_cel", "visibility:hidden", marca_error(), "")
	
	'<input name="e_mail" type="text" id="e_mail" size="50" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	e_mail = LCase(campo("e_mail", "text", "e_mail", "30", "50", "6", e_mailr,"OnBlur='caractNoPermit(this)' OnChange='errorCorreo(this)'"))
	e_mail = e_mail & LCase(hidden("e_mail2", e_mail2r))
	ast_e_mail = tag("span", "ast_e_mail", "visibility:hidden", marca_error(), "")
	
	'<input name="direcc" type="text" id="direcc" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	direcc = campo("direcc", "text", "direcc", "30", "100", "7", direccr, "OnBlur='caractNoPermit(this)'")
	ast_direcc = tag("span", "ast_direcc", "visibility:hidden", marca_error(), "")
	
	ocupa = campo("ocupa", "text", "ocupa", "30", "100", "7", ocupar, "OnBlur='caractNoPermit(this)'")
	ast_ocupa = tag("span", "ast_ocupa", "visibility:hidden", marca_error(), "")
	
	ciu_exp = campo("ciu_exp", "text", "ciu_exp", "30", "100", "7", ciu_expr, "OnBlur='caractNoPermit(this)'")
	ast_ciu_exp = tag("span", "ast_ciu_exp", "visibility:hidden", marca_error(), "")
	
	tipo_visa = campo("tipo_visa", "text", "tipo_visa", "30", "100", "7", tipo_visar, "OnBlur='caractNoPermit(this)'")
	ast_tipo_visa = tag("span", "tipo_visa", "visibility:hidden", marca_error(), "")

	'<input name="ciudad" type="text" id="ciudad" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	ciudad = campo("ciudad", "text", "ciudad", "30", "40", "8", ciudadr, "OnBlur='caractNoPermit(this)'") & hidden("cod_ciud", "")
	ast_ciudad = tag("span", "ast_ciudad", "visibility:hidden", marca_error(), "")
	
	'<input name="num_pasap" type="text" id="num_pasap" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	'num_pasap = campo("num_pasap", "text", "num_pasap", "30", "20", "7", num_pasapr, "")'
	'ast_num_pasap = tag("span", "ast_num_pasap", "visibility:hidden", marca_error(), "")'
	
	'<input name="num_pasap" type="text" id="num_pasap" size="30" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	num_pasap = campo("num_pasap", "text", "num_pasap", "30", "20", "9", num_pasapr, "OnBlur='caractNoPermit(this)'")
	num_pasap = num_pasap & hidden("num_pasap2", num_pasap2r)
	ast_num_pasap = tag("span", "ast_num_pasap", "visibility:hidden", marca_error(), "")
	
	'<input name="fec_venc" type="text" id="fec_venc" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	fec_venc = campo_readonly("fec_venc", "text", "fec_venc", "30", "20", "10", fec_vencr, "readonly", "")
	ast_fec_venc = tag("span", "ast_fec_venc", "visibility:hidden", marca_error(), "")
	
	'<a href="javascript:NewCssCal('fec_venc', 'ddMMyyyy')">'
	'<img src="../images_cal/cal.gif" width="16" height="16" alt="Pick a date" border="0"></a>'
	cal2 = enlace("", "javascript:NewCssCal(""fec_venc"", ""ddMMyyyy"")", "", "", imgcal, "")
	
	'<input name="img_pasap" type="file" id="img_pasap" size="20" maxlength="20" />'
	'img_pasap = campo("img_pasap", "file", "img_pasap", "30", "20", "11", img_pasapr, "")
	'ast_img_pasap = tag("span", "ast_img_pasap", "visibility:hidden", marca_error(), "")
	
	
	
	
	
	 img_pasap = campo_readonly("img_pasap", "text", "img_lic_tri", "30", "50", "20", img_pasapr, "readonly", "") & hidden("img_pasapr", img_pasap) & button("bot_pasaporte", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=0"",""Pasaporte"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")  
		   
		     'botonVer2 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
			 botonVer1 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion3(""obtener.asp?j=1&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
	
	'<input name="num_visa_eua" type="text" id="num_visa_eua" size="20" maxlength="20" readonly="readonly"/>'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	num_visa_eua = campo("num_visa_eua", "text", "num_visa_eua", "30", "20", "12", num_visa_euar, "OnBlur='caractNoPermit(this)'")
	num_visa_eua = num_visa_eua & hidden("num_visa_eua2", num_visa_eua2r)
	ast_num_visa_eua = tag("span", "ast_num_visa_eua", "visibility:hidden", marca_error(), "")
	
	'<input name="fec_venc_visa_eua" type="text" id="fec_venc_visa_eua" size="30" maxlength="20"  readonly="readonly"/>'
	fec_venc_visa_eua = campo_readonly("fec_venc_visa_eua", "text", "fec_venc_visa_eua", "30", "20", "11", fec_venc_visa_euar, "readonly", "")
	
	'<a href="javascript:NewCssCal('fec_venc_visa_eua', 'ddMMyyyy')">'
	'<img src="../images_cal/cal.gif" width="16" height="16" alt="Pick a date" border="0"></a>'
	cal3 = enlace("", "javascript:NewCssCal(""fec_venc_visa_eua"", ""ddMMyyyy"")", "", "", imgcal, "")
	ast_fec_venc_visa_eua = tag("span", "ast_fec_venc_visa_eua", "visibility:hidden", marca_error(), "")
	
	'<input name="img_visa_eua" type="file" id="324" size="20" maxlength="20"/>'
	'img_visa_eua = campo("img_visa_eua", "file", "img_visa_eua", "30", "20", "13", img_visa_euar, "")
	'ast_img_visa_eua = tag("span", "ast_img_visa_eua", "visibility:hidden", marca_error(), "")
	img_visa_eua = campo_readonly("img_visa_eua", "text", "img_lic_tri", "30", "50", "20", img_visa_euar, "readonly", "") & hidden("img_visa_euar", img_visa_eua) & button("bot_visa", "Adjuntar", "OnClick='MM_openBrWindow(""subir_archivo.asp?v=1"",""Visa"",""scrollbars=0,resizable=0,location=0,status=0,scrollbars=0,"",""450"",""150"")'") '& button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")  
		   
		     'botonVer2 = button2("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion(""obtener.asp?j=5&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'","button")
			 
	botonVer2 = button("pasp" & CStr(ind), "Ver", "OnClick='MM_openBrWindowNewVersion3(""obtener.asp?j=2&n="", ""winName"",""scrollbars=yes, menubar=no, resizable=yes"",""1000"",""450"")'")
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
		'if sexor ="on" then
	'		sexo2 = radio("sexo", "sexo", "1",  "checked")
'		else
		    'sexo = radio("sexo", "sexo", "0", "checked")
	'	end if		
	End if
	remitidor = ""
	remitido = textarea2("remitido", "30", "2", "100", remitidor, "OnBlur='caractNoPermit(this)' style='width: 216px; height: 47px;'")'
	
	'<img src="../Imagenes/boton_agregar.jpg" width="70" height="19" />'
	'enlace(id, dir, estilo, clase, html, exprJs)'
	html = img("../Imagenes/boton_agregar.jpg", "70", "19", "", "", "")
	boton_agregar = enlace("", "#", "text-decoration:none", "", html, "OnClick='envia()'")
	'<img src="../imagenes/boton_borrar.jpg" width="70" height="19" /'
	html3 = img("../Imagenes/boton_borrar.jpg", "70", "19", "", "", "")
	'boton_borrar = enlace("", "#", "text-decoration:none", "", html3, "OnClick='elimina(document.form1.lista.value)'")
	boton_borrar = enlace("", "#", "text-decoration:none", "", html3, "OnClick='elimina(document.form1.num_pasa.value)'")
	
	'html4 = img("../Imagenes/boton_aceptar.jpg", "70", "19", "", "", "")
	'camp_busc = campo("camp_busc", "text", "camp_busc", "20", "100", "14", "", "OnBlur='caractNoPermit(this)'")
	'boton_buscar = enlace("", "#", "text-decoration:none", "", html4, "OnClick='buscar()'")
	
	
	'***************************************************************************************************************************************
	      'PRODUCCION	
		  
		 filtrobusqueda = request.QueryString("busquedaPas") 
        
         if filtrobusqueda <> "" then            		 
	        
		     strSqln = "Select s.*, p.Cod_pais, p.Nom_pais From Siep_pasa s, Siep_pais p  Where (p.Cod_pais = s.Nac_pasaj) And  (s.Num_clie = " & Session("Num_clie") & ")  And (s.Nom_Pasa like '%" & filtrobusqueda & "%') Order By s.Nom_pasa"	 
		 
		 else
		    strSqln = "Select s.* From Siep_pasa s Where  s.Num_pasa = 0 Order By s.Nom_pasa"	 
		 end if

         
		  
		ADO.Open strSqln, Conexion
	
	    htmlbus = img("../Imagenes/boton_aceptar.jpg", "70", "19", "", "", "")
		bus_pas = "<table id='busqueda_pas' class='display'  >"
        bus_pas = bus_pas& "<thead>"
		bus_pas = bus_pas & "<tr>"		
		bus_pas = bus_pas & "Filtro busqueda: " & campo("busquedaPas", "text", "bus_pas", "20", "20", "8","", "onkeypress='busPas()'")
		bus_pas = bus_pas &  enlace("", "javascript:void(0)", "text-decoration:none", "", htmlbus, "OnClick='submit()'")
		bus_pas = bus_pas & "</tr>"
		bus_pas = bus_pas & "</thead>"		
	
	
	    tn = "<table id='table_id' class='display'  >"
		tn = tn & "<thead>"
		tn = tn & "<tr>"		
		tn = tn & "<th><span class='Estilo25' style='width:400px;'>Nombre / Apellido</span></th>"	
		tn = tn & "<th><span class='Estilo25'>Cedula</span></th>"
		tn = tn & "<th><span class='Estilo25'>Teléfono</span></th>"	
		
		
		tn = tn & "<th><span class='Estilo25' style='width:200px;'>Email</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Direccion</span></th>"
		
		tn = tn & "<th><span class='Estilo25' style='width:100px;'>Ciudad</span></th>"		
		tn = tn & "<th><span class='Estilo25'>Número Pasaporte</span></th>"	
		tn = tn & "<th><span class='Estilo25' style='width:100px;'>Vencimiento Pasaporte</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Imagen Pasaporte</span></th>"	
		
		tn = tn & "<th><span class='Estilo25' style='width:100px;'>Número Visa</span></th>"		
		tn = tn & "<th><span class='Estilo25' style='width:100px;'>Vencimiento Visa</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Imagen Visa</span></th>"	
		
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Apellido</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Ocupacion</span></th>"
		
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Pais de Residencia</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Pais de 	Nacionalidad</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Estado</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'>Fecha de Nacimiento</span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'></span></th>"
		tn = tn & "<th><span class='Estilo25' style='width:800px;'></span></th>"
		
		tn = tn & "</tr>"
		tn = tn & "</thead>"
		
		
		
		
		if not ADO.EOF then
		
			tn = tn & "<tbody>"		
				'
				Do Until ADO.EOF
				   tn = tn & "<tr "&estilo&">"
				   'tn = tn & "<td  class='Estilo4' style='color:#000;'>" & rstx("Nom_trip") & "</td>"
				    tn = tn & "<td  class='Estilo4' style='color:#000;width:400px;'>"&ADO("Nom_pasa")&" "&ADO("Ape_pasa")&"</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;'>" & ADO("Num_cedu_pas") & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;'>" & ADO("Tlf_movi_pas") & "</td>"	
									
					tn = tn & "<td  class='Estilo4' style='color:#000;width:200px;'>" & ADO("Ema_pasa") & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & ADO("Dir_pasa") & "</td>"
					
					
					If Not IsNull(ADO("Cod_ciud")) then
				      Ciuda = selectSimpleSQL2("Nom_ciuda", "Ciudad", "cod_ciuda = " & ADO("Cod_ciud"), 1)
			        Else
				      Ciuda = ""
			        End if
					
				    tn = tn & "<td  class='Estilo4' style='color:#000;width:100x;'>" & Ciuda & "</td>"
					
					
					
					
					
					
					tn = tn & "<td  class='Estilo4' style='color:#000;'>" & ADO("Num_pasa_pas") & "</td>"					
				    tn = tn & "<td  class='Estilo4' style='color:#000;width:100px;'>" & ADO("Fec_venc_pas") & "</td>"
					
					
					if ADO("Ima_pasa") <> "" then
					
					    tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'><a href='" & ADO("Ima_pasa") & "' target='_blank'>VER</a></td>"
					else
					
					    tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'></td>"
					
					end if
					
					
					tn = tn & "<td  class='Estilo4' style='color:#000;width:100px;'>" & ADO("Num_visa_pas") & "</td>"					
				    tn = tn & "<td  class='Estilo4' style='color:#000;width:100px;'>" & ADO("Fec_venc_vis") & "</td>"
					
					if ADO("Ima_visa") <> "" then
					
					   tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'><a href='" & ADO("Ima_visa") & "' target='_blank'>VER</a></td>"
					   
					 else
					 
					   tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>&nbsp;</td>"
					 
					 end if   
					
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & ADO("Ape_pasa") & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & ADO("ocu_pas") & "</td>"
					
					
					
					
					
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & selectSimpleSQL2("nom_pais", "Siep_pais", "Cod_pais = " & ADO("Pais_res_pas") , 1) & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & selectSimpleSQL2("nom_pais", "Siep_pais", "Cod_pais = " & ADO("Nac_pasaj") , 1) & "</td>"					
					
					If Not IsNull(ADO("Nac_pasaj")) and Not IsNull(ADO("Est_dir_pas"))  then
				      estado = selectSimpleSQL2("des_esta", "Siep_esta", "Cod_pais = " & ADO("Nac_pasaj") & " and Cod_esta = " & ADO("Est_dir_pas"),1)
			        Else
				      estado = ""
			        End if					
					
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & estado & "</td>"
					
					
					
					
					'tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" &  rstn("Est_dir_pas")  & "</td>"
					
					
				    tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & ADO("Fec_nac_pas") & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;'>" & ADO("Nom_pasa") & "</td>"
					tn = tn & "<td  class='Estilo4' style='color:#000;width:800px;visibility:hidden'>" & ADO("Num_pasa") & "|" & ADO("Sex_pas") & "|" & ADO("Sta_acti") & "</td>"
					
					
				   tn = tn & "</tr>"
				   ADO.MoveNext
				Loop
				
			tn = tn & "</tbody>	"
			   
		End if
		
		
	tn = tn & "</table>"	
	
	
	ADO.Close
	
	
	
	
	
	
	
	'*****************************************************************************************************************************************
	
	
	
	ind = 1
	
	
	
	
	
	
	
	Conexion.Close
	Set ADO = nothing	
	Set Conexion = nothing	
	env = hidden("env", "1")
	aleat = hidden("aleat", "")
	lista = hidden("lista", ind)
	edit_inser_borr =  hidden("edit_inser_borr", val_inser)
		'edit_inser_borr = campo("edit_inser_borr", "text", "edit_inser_borr", "30", "50", "1", 0, "")
	
	
	'<img src="../imagenes/boton_editar.jpg" width="70" height="19" />'
	html2 = img("../Imagenes/boton_editar.jpg", "70", "19", "", "", "")
	boton_editar = enlace("", "#", "text-decoration:none", "", html2,  "OnClick='llenaCampEdit(document.form1.lista.value)'")	
	
	htrans = img("../imagenes/botontransferir.gif", "125", "28", "", "", "")
	transferir = enlace("", "transferencia_pasaj.asp", "text-decoration:none", "", htrans, "")
	
	ht = img("../imagenes/botonsalir.gif", "125", "28", "", "", "")
	salir = enlace("", "../salir.asp", "text-decoration:none", "", ht, "")
	
	
	
	
	
	
	
	
	Set load = Nothing
	'Set Upload = Nothing'
	if Err.Number <> 0 then
		Response.Write Err.Description & " " & Err.Number
		Response.End()
		Error.Clear
	End if

%>
<!--#include file="../template/pasajero.html" -->

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 