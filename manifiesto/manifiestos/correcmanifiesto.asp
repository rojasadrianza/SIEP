<% @ LCID = 1034 %>
<!--#include file="../modulos/vali_sesion.asp"-->
<!--#include file="../modulos/funcion.asp" -->
<!--#include file="../modulos/control.asp" -->
<!--#include file="../js/funcion.js" -->
<%
	'On Error Resume Next
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'////////////////////////////////////////////////////////////
	'////////////////////////////////////////////////////////////
	'El archivo registromanifiesto_guar.asp está muy sobrecargado
	'así que se va a tener que hacer todas las consultas por acá
	'de nuevo. Tampoco es muy conveniente pasar demasiada información 
	'por get
	'////////////////////////////////////////////////////////////
	'////////////////////////////////////////////////////////////
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
	bck = Request("bck")
	regreso = Request("r")
	chk_pasap = Request("c1")
	chk_trip = Request("c2")
	modi = Request("v")
	lst_pasa = Request("ps") 
	cantTri = Request("ctr") 
	lst_tri = Request("tr") 
	c_tr = Request("c_tr")
	num_per_tot = Request("num_per_tot")
	fec_salida = Request("dat")
	selecpasinput = request("selecpasinput") 
	selectripinput = request("selectripinput")	
	canttrip = request("canttrip") 
	cantpas = request("cantpas")
	
	
	'response.write "selecpasinput " & selecpasinput &" "& cantpas
	'response.write "selectripinput " & selectripinput &" "& canttrip
	'response.write "1 " & bck & "</br>"
	'response.write "2 " & regreso & "</br>"
	'response.write "3 " & chk_pasap & "</br>"
	'response.write "4 " & chk_trip & "</br>"
	'response.write "5 " & modi & "</br>"
	'response.write "6 " & lst_pasa & "</br>"
	'response.write "7 " & cantTri & "</br>"
	'response.write "8 " & lst_tri & "</br>"
	'response.write "9 " & c_tr & "</br>"
	
	'response.End()
	
	

	
	lista_reg = Split(regreso, "|")
	'for each x in lista_reg
	'	response.write(x & "<br />")
	'next

	aeropStr = lista_reg(1) & "‡" & lista_reg(3)
	aerop_lst = Split(aeropStr, "‡")
	'for each y in aerop_lst
    '	response.write(y & "<br />")
	'next
 	
	pasap_msj_err = ""
	tripu_msj_err = ""
 	aeronave = ""
	aeronave = lista_reg(0)
	aerop_sal = aerop_lst(0) & "/" & aerop_lst(1)
	aerop_lleg = aerop_lst(2) & "/" & aerop_lst(3)		
 	Set cnnf = Server.CreateObject("ADODB.Connection")
	Set rstf = Server.CreateObject("ADODB.Recordset")
	cnnf.Open dsn()
	'Código aeropuerto de salida aerop_lst(0)'
	rstf.Open  "Select Cod_aero, Cod_ciud From Siep_airp Where Cod_aero = '" & aerop_lst(0) & "'", cnnf, 1, 2	
	Cod_ciud_sal = rstf("Cod_ciud") 
	rstf.Close
	'Código aeropuerto de llegada aerop_lst(2)'

	rstf.Open  "Select Cod_aero, Cod_ciud From Siep_airp Where Cod_aero = '" & aerop_lst(2) & "'", cnnf, 1, 2
	Cod_ciud_lle = rstf("Cod_ciud") 
	rstf.Close
	
	rstf.Open  "Select Cod_pais From Ciudad Where Cod_ciuda = " & Cod_ciud_sal , cnnf, 1, 2
	Cod_pais_sal = rstf("Cod_pais") 
	rstf.Close
	
	rstf.Open  "Select Cod_pais From Ciudad Where Cod_ciuda = " & Cod_ciud_lle , cnnf, 1, 2
	Cod_pais_lle = rstf("Cod_pais") 
	rstf.Close
	
	rstf.Open  "Select Ntr_aero, Pue_aero From Siep_aero Where (Sig_aero = '" & aeronave & "') and (Num_clie = " & Session("Num_clie") & ")" , cnnf, 1, 2
	if IsNull(rstf("Ntr_aero")) then
		ctripulantes = "nulo"
	else
		ctripulantes = rstf("Ntr_aero")
	End if 
	rstf.Close
	
	
	numero_tripul = ""
	if c_tr = "no" then
		numero_tripul = "<tr>"
		numero_tripul = numero_tripul & "<td><strong>N&uacute;mero de tripulantes</strong></td></tr>"
		numero_tripul = numero_tripul & "<tr>"
		if ctripulantes = "nulo" then
			numero_tripul = numero_tripul & "<td>El n&uacute;mero de tripulantes no puede ser nulo (revisar aeronave) </td></tr>"
		else
			numero_tripul = numero_tripul & "<td>El n&uacute;mero de tripulantes no puede ser menor de " & ctripulantes & "</td></tr>"
		End if
	End if
	
	if cantTri <> "0" then
		cond = " Num_trip in " & lst_tri
	Else
		cond = " Num_trip = " & lst_tri
	End If
	

	
	apellido_pasaj = ""
	apellido_trip = "" 

	if Cod_pais_lle = "1" And Cod_pais_sal = "1" then
		if lst_pasa <> "nn" then
			rstf.Open  "Select Num_pasa, Nom_pasa, Ape_pasa, Num_cedu_pas, Num_pasa_pas, Pais_res_pas, Nac_pasaj, Fec_venc_pas, Num_cedu_pas From Siep_pasa Where Num_pasa in (" & lst_pasa & ")", cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_pasa")) Or Len(rstf("Ape_pasa")) = 0 Or rstf("Ape_pasa") = "" then
					pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & " Sin Apellido"  & "<br />"
				'Else
				'	apellido_pasaj = " " & rstf("Ape_pasa")
				End If
				'///////18-02-2014 -> se hizo una exepción con esta validación
				'Pasajero extranjero
				'if (IsNull(rstf("Pais_res_pas")) Or Len(Trim(rstf("Nac_pasaj"))) = 0 Or Trim(rstf("Nac_pasaj")) = "") then 
				' 	pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & " Sin país de residencia"  & "<br />"
				'End if
				'if ((rstf("Pais_res_pas") <> "1") and (Not IsNull(rstf("Pais_res_pas")))) then 
				'	if IsNull(rstf("Nac_pasaj")) Or Len(rstf("Nac_pasaj")) = 0 Or rstf("Nac_pasaj") = "" then
				'		pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & " Sin país de nacimiento"  & "<br />"
				'	End If
				'	if IsNull(rstf("Num_pasa_pas")) Or Len(rstf("Num_pasa_pas")) = 0 Or rstf("Num_pasa_pas") = "" then
				'		pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & apellido_pasaj & " Sin pasaporte"  & "<br />"
				'	Else
				'		if Not IsNull(rstf("Fec_venc_pas")) then
				'			if DateDiff("d", Now(), rstf("Fec_venc_pas")) <= 30 then
				'				pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & apellido_pasaj & " Con pasaporte pr&oacute;ximo a vencerse"  & "<br />"
				'			End if
				'		End if
				'	End If		
				'End If
				'Pasajero venezolano
				'if rstf("Pais_res_pas") = "1" then 	
				'	if IsNull(rstf("Num_cedu_pas")) Or Len(rstf("Num_cedu_pas")) = 0 Or rstf("Num_cedu_pas") = "" then
				'		pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & apellido_pasaj & " Sin cédula"  & "<br />"
				'	End If
				'End if
				rstf.MoveNext
			Loop
			rstf.Close
		End if
		'///////////////////////////////////
		
		if Mid(aeronave, 1, 2) = "YV" then
			rstf.Open  "Select Num_trip, Nom_trip, Ape_trip, Num_lic_tri, Num_cer_tri, Fec_venc_lic, Fec_venc_cer From Siep_trip Where " & cond, cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_trip")) Or Len(rstf("Ape_trip")) = 0 Or rstf("Ape_trip") = "" then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & " Sin apellido"  & "<br />"
				Else
					apellido_trip = " " & rstf("Ape_trip")
				End If
				if IsNull(rstf("Num_lic_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin licencia venezolana"  & "<br />"
				Else
					if IsNull(rstf("Fec_venc_lic")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Licencia venezolana sin fecha de vencimiento"  & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						'if DateDiff("d", Now(), rstf("Fec_venc_lic")) <= 30 then
						if DateDiff("d", fec_salida, rstf("Fec_venc_lic")) < 0 then	
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Licencia venezolana vencida"  & "<br />"
						End if
					End if
				End if
				'Se valida si el certificado de Vzla está vencido
				'Irving -> 16-09-2015
				if IsNull(rstf("Num_cer_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin certificado médico de Venezuela" & "<br />"
				Else
					if IsNull(rstf("Fec_venc_cer")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con certificado médico de Venezuela sin fecha de vencimiento" & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_cer")) < 0 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " certificado médico de Venezuela vencido" & "<br />"
						End if
					End if
				End if
				rstf.MoveNext
			Loop
			rstf.Close
		End if		
		
		if Mid(aeronave, 1, 1) = "N" then
			rstf.Open  "Select Num_trip, Nom_trip, Ape_trip, Num_cedu_tri, Num_lic_eua_tri, Num_cer_eua_tri, Fec_exp_eua_cer From Siep_trip Where " & cond, cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_trip")) Or Len(rstf("Ape_trip")) = 0 Or rstf("Ape_trip") = "" then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & " Sin apellido"  & "<br />"
				Else
					apellido_trip = " " & rstf("Ape_trip")
				End If
				if IsNull(rstf("Num_lic_eua_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin licencia de EUA"  & "<br />"
				End if
				if IsNull(rstf("Num_cer_eua_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin certificado medico de EUA"  & "<br />"
				Else
					if IsNull(rstf("Fec_exp_eua_cer")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Cerificado de EUA sin fecha de vencimiento"  & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d",fec_salida, vencCerEUA(rstf("Fec_exp_eua_cer"))) <= 30 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con certificado medico de EUA vencido"  & "<br />"
						End if
					End if
				End if
				rstf.MoveNext
			Loop
			rstf.Close
		End if
		'///////////////////////////////////
	End if
	
	msj_pal = ""
	msj_pal2 = ""
	if Cod_pais_lle <> "1" Or Cod_pais_sal <> "1" then
		if lst_pasa <> "nn" then
			rstf.Open  "Select Num_pasa, Nom_pasa, Ape_pasa, Num_pasa_pas, Pais_res_pas, Nac_pasaj, Fec_venc_pas, Num_cedu_pas, Num_visa_pas, Fec_venc_vis From Siep_pasa Where Num_pasa in (" & lst_pasa & ")", cnnf , 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_pasa")) Or Len(rstf("Ape_pasa")) = 0 Or rstf("Ape_pasa") = "" then
					msj_pal = msj_pal & rstf("Nom_pasa") & " Sin Apellido" & "<br />" 
				Else
					apellido_pasaj = " " & rstf("Ape_pasa")
				End If
				if IsNull(rstf("Nac_pasaj")) Or Len(rstf("Nac_pasaj")) = 0 Or rstf("Nac_pasaj") = "" then
					msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Sin país de nacionalidad" & "<br />"
				End If
				if IsNull(rstf("Num_pasa_pas")) Or Len(rstf("Num_pasa_pas")) = 0 Or rstf("Num_pasa_pas") = "" then
					msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Sin pasaporte" & "<br />"
				Else
					if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
						msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Con pasaporte sin fecha de vencimiento" & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_pas")) < 0 then
							msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Con pasaporte vencido" & "<br />"
						End if
					End if
				End If
				
				'Visa EUA
				if Cod_pais_lle = "4" then
					if IsNull(rstf("Num_visa_pas")) then
						msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Sin visa de EUA" & "<br />"
					Else
						if IsNull(rstf("Fec_venc_vis")) then
							msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Sin fecha de vencimiento visa de EUA" & "<br />"
						Else
							'Da error si el documento está vencido
							' Now() -> fec_salida
							if DateDiff("d", fec_salida, rstf("Fec_venc_vis")) < 0 then
								msj_pal2 = msj_pal2 & rstf("Nom_pasa") & apellido_pasaj & " Con visa de EUA vencida" & "<br />"
							End if
						End if
					End if
				end if
				
				'Paises del mercosur
				if Cod_pais_lle = "1" Or Cod_pais_sal = "1" then	
					if ((mercosur(Cod_pais_lle) = 1) Or (mercosur(Cod_pais_sal) = 1)) then
						msj_pal2 = ""	
					End if
				End if
				'if (IsNull(rstf("Pais_res_pas")) Or Len(Trim(rstf("Nac_pasaj"))) = 0 Or Trim(rstf("Nac_pasaj")) = "") then 
				'	pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & " Sin país de residencia"  & "<br />"
				'End if
				'if rstf("Pais_res_pas") = "1" then 	
				'	if IsNull(rstf("Num_cedu_pas")) Or Len(rstf("Num_cedu_pas")) = 0 Or rstf("Num_cedu_pas") = "" then
				'		pasap_msj_err = pasap_msj_err & rstf("Nom_pasa") & apellido_pasaj & " Sin cédula de identidad" & "<br />"
				'	End If
				'End if
				rstf.MoveNext
			Loop
			rstf.Close
			pasap_msj_err = msj_pal & msj_pal2
		End if		
		if Mid(aeronave, 1, 2) = "YV" then
			rstf.Open  "Select Num_cedu_tri, Nom_trip, Ape_trip, Num_pasa_tri, Fec_venc_pas, Nac_tri, Fec_nac_tri, Num_lic_tri, Num_lic_tri, Num_cer_tri, Fec_venc_lic, Fec_venc_cer, Fec_venc_pas From Siep_trip Where " & cond, cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_trip")) Or Len(rstf("Ape_trip")) = 0 Or rstf("Ape_trip") = "" then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & " Sin apellido" & "<br />"
				Else
					apellido_trip = " " & rstf("Ape_trip")
				End If
				if IsNull(rstf("Num_pasa_tri")) then
					tripu_msj_err = tripu_msj_err& rstf("Nom_trip") & apellido_trip & " Sin pasaporte" & "<br />"
				End if
				if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con pasaporte sin fecha de vencimiento" & "<br />"
				Else
					'Ahora no se procesa si está vencido (sept 2015) -> cliente
					' Now() -> fec_salida
					if DateDiff("d", fec_salida, rstf("Fec_venc_pas")) < 0 then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con pasaporte vencido" & "<br />"
					End if
				End if
				
				if IsNull(rstf("Num_lic_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin licencia venezolana" & "<br />"
				Else
					if IsNull(rstf("Fec_venc_lic")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con licencia venezolana sin fecha de vencimiento" & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_lic")) < 0 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con licencia venezolana vencida" & "<br />"
						End if
					End if
				End if
				
				'Se valida si el certificado de Vzla está vencido
				'Irving -> 16-09-2015
				if IsNull(rstf("Num_cer_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin certificado médico de Venezuela" & "<br />"
				Else
					if IsNull(rstf("Fec_venc_cer")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con certificado médico de Venezuela sin fecha de vencimiento" & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_cer")) < 0 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " certificado médico de Venezuela vencido" & "<br />"
						End if
					End if
				End if
				
				rstf.MoveNext
			Loop
			rstf.Close
		End if	
		
		if Mid(aeronave, 1, 1) = "N" then
			rstf.Open  "Select Num_pasa_tri, Nom_trip, Ape_trip, Fec_venc_pas, Nac_tri, Fec_nac_tri, Num_lic_eua_tri, Num_cer_eua_tri, Fec_exp_eua_cer, Fec_venc_pas, Num_visa_tri, Fec_venc_vis From Siep_trip Where " & cond, cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_trip")) Or Len(rstf("Ape_trip")) = 0 Or rstf("Ape_trip") = "" then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin apellido" & "<br />"
				Else
					apellido_trip = " " & rstf("Ape_trip")
				End If
				if IsNull(rstf("Num_pasa_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin pasaporte" & "<br />"
				Else
					if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con pasaporte sin fecha de vencimiento" & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_pas")) < 0 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con pasaporte vencido" & "<br />"
						End if
					End if
				End if
		
				'Response.Write " " & IsNull(rstf("Fec_venc_pas")) & " " & Len(rstf("Fec_venc_pas")) = 0 & " " & rstf("Fec_venc_pas")
				if IsNull(rstf("Num_lic_eua_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin licencia de EUA" & "<br />"
				End if
				if IsNull(rstf("Num_cer_eua_tri")) then
					tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin certificado medico de EUA" & "<br />"
				Else
					if IsNull(rstf("Fec_exp_eua_cer")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin fecha de expedición del certificado medico de EUA " & "<br />"
					Else
						'Ahora no se procesa si está vencido (sept 2015) -> cliente
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, vencCerEUA(rstf("Fec_exp_eua_cer"))) < 0 then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con certificado médico de EUA vencido" & "<br />"
						End if
					End if
				End if
				'Se valida la visa si el destino es EEUU
				'Irving -> 16-09-2015 
				if Cod_pais_lle = "4" then
					if IsNull(rstf("Num_visa_tri")) then
						tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin visa de EUA" & "<br />"
					Else
						if IsNull(rstf("Fec_venc_vis")) then
							tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Sin fecha de vencimiento de la visa de EUA" & "<br />"
						Else
							'Ahora no se procesa si está vencido (sept 2015) -> cliente
							' Now() -> fec_salida
							if DateDiff("d", fec_salida, rstf("Fec_venc_vis")) < 0 then
								tripu_msj_err = tripu_msj_err & rstf("Nom_trip") & apellido_trip & " Con visa de EUA vencida" & "<br />"
							End if
						End if
					End if
				end if
				rstf.MoveNext
			Loop
			rstf.Close
		End if	
	End if
	if (tripu_msj_err = "") then
		tripu_msj_err = "----"
	End if
	if (pasap_msj_err = "") then
		pasap_msj_err = "----"
	End if
	
	 	
 	if modi = "-" then		
		direccion = "registromanifiesto.asp?bck=1&r=" & regreso & "&c1=" & chk_pasap & "&c2=" & chk_trip & "&num_per_tot=" & num_per_tot & "&selecpasinput=" & selecpasinput & "&selectripinput=" & selectripinput & "&canttrip=" & canttrip & "&cantpas=" & cantpas
	else
		direccion = "modimanifiesto.asp?v=" & modi & "&num_per_tot=" &num_per_tot
	End if
	
	ht = img("../imagenes/boton_volver.jpg", "70", "19", "", "", "")
	volver = enlace("", direccion, "text-decoration:none", "", ht, "")
 	
%>
<!--#include file="../template/correcmanifiesto.html" -->
