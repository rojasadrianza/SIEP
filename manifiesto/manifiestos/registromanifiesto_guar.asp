<% @ LCID = 1034 %>
<!--#include file="../modulos/vali_sesion.asp"-->
<!--#include file="../modulos/funcion.asp" -->
<%
	'On Error Resume Next
	aeronave = Request.Form("aeronave")
	ae_orig = Request.Form("aerop_orig")
	ae_lleg = Request.Form("aerop_lleg")
	fec_salida = Request.Form("fec_salida")
	'fec_llegada = Request.Form("fec_llegada")
	fec_llegada =""
	'------------------------------------------
	selecpasinput = request.form("selecpasinput") 
	selectripinput = request.form("selectripinput")	
	canttrip = request.form("canttrip")
	cantpas = request.form("cantpas")
	
	
	
	
	
	if Instr(ae_orig,"|") > 0 then
		
		ls_aerop_orig = Split(ae_orig, "|")
		aerop_orig = ls_aerop_orig(1)	
		regr_aerop_orig = ls_aerop_orig(0) & "‡" & ls_aerop_orig(1)
	else
		aerop_orig = ae_orig
		regr_aerop_orig = ae_orig
	End if
	

	if Instr(ae_lleg,"|") > 0 then
		
		ls_aerop_lleg = Split(ae_lleg, "|")
		aerop_lleg = ls_aerop_lleg(1)	
		regr_aerop_lleg = ls_aerop_lleg(0) & "‡" & ls_aerop_lleg(1)
	else
		aerop_lleg = ae_lleg
		regr_aerop_lleg = ae_lleg
	End if
	'------------------------
	num_per_tot = Request.Form("num_per_tot") 
	num_per = Request.Form("num_per")
	num_pas = Request.Form("cantpas")
	num_tri = Request.Form("canttrip")
	
        if num_pas = "" then 
           num_pas = 0
        end if

        if num_tri = "" then
           num_tri = 0
        end if

        pic_list = Request.Form("pic_list")
	hora = Request.Form("hor_salida")
	minuto = Request.Form("min_salida")
	if((hora <> "--") And (minuto <> "--")) then	
		hor_salida = hora & ":" & minuto
	else
		hor_salida = ""
	End if
	hor_llegada=""
	num_tri_bd = Request.Form("num_tri_bd")
	num_pas_bd = Request.Form("num_pas_bd")
	modi = Request.Form("modifica")
	'Copiloto'
	copic_list = Request.Form("copic_list")
	regreso = ""
	regreso = aeronave & "|" & regr_aerop_orig & "|" & fec_salida & "|" & regr_aerop_lleg & "|" & num_per & "|" 
	regreso = regreso & hor_salida & "|" & num_pas  & "|" & num_tri & "|" & pic_list & "|" & copic_list & "|" & Request.Form("copic_list_rev")
	
	Set cnnf = Server.CreateObject("ADODB.Connection")
	Set rstf = Server.CreateObject("ADODB.Recordset")
	cnnf.Open dsn()
	
	'Código aeropuerto de salida'
	rstf.Open  "Select Cod_aero, Cod_ciud From Siep_airp Where Nom_aero = '" & aerop_orig & "'", cnnf, 1, 2	
	Cod_aero_sal = rstf("Cod_aero")
	Cod_ciud_sal = rstf("Cod_ciud") 
	rstf.Close
	'Código aeropuerto de llegada'

	rstf.Open  "Select Cod_aero, Cod_ciud From Siep_airp Where Nom_aero = '" & aerop_lleg & "'", cnnf, 1, 2
	Cod_aero_lle = rstf("Cod_aero")
	Cod_ciud_lle = rstf("Cod_ciud") 
	rstf.Close
	
	rstf.Open  "Select Cod_pais From Ciudad Where Cod_ciuda = " & Cod_ciud_sal , cnnf, 1, 2
	Cod_pais_sal = rstf("Cod_pais") 
	rstf.Close
	
	rstf.Open  "Select Cod_pais From Ciudad Where Cod_ciuda = " & Cod_ciud_lle , cnnf, 1, 2
	Cod_pais_lle = rstf("Cod_pais") 
	rstf.Close
	
	'manifiesto'
	
	'El código de Venezuela es 1'
	err_cant_tri = 0
	err_ape_pasaj = 0
	err_nac_pasaj = 0
	err_ced_pasaj = 0
	err_pasap_pasaj = 0
	err_visa_pasaj = 0
	err_vvisa_pasaj = 0
	
	err_ape_trip = 0
	err_pasap_trip = 0
	'Piloto y copiloto
	err_visa_cab = 0
	err_lic_cab = 0
	err_licv_cab = 0
	err_licEua_cab = 0
	err_cerEua_cab = 0
	err_ced_cab = 0
	err_cer_cab = 0 
	err_cerv_cab = 0 
	err_visa_cab = 0 
	err_visav_cab = 0 
	'cod_ced_err_cab = ""
	
	'///////////////
	'Consulta aeronave
	'response.write num_tri & "</br>"
		
		'response.End()
	
	rstf.Open "Select Ntr_aero, Pue_aero From Siep_aero Where (Sig_aero = '" & aeronave & "') and (Num_clie = " & Session("Num_clie") & ")", cnnf, 1, 2		
	c_tr = "si"
	if(IsNull(rstf("Ntr_aero"))) then
		c_tr = "no" 
		err_cant_tri = 1
	Else
		Cant_trip = rstf("Ntr_aero")
                num_tri = 0
		
		'SE LE SUMA EL PILOTO***********************HUMBERTO ROJAS
		num_tri = num_tri + 1
		'SE LE SUMA EL COPILOTO*********************HUMBERTO ROJAS
		if pic_list <> "" then
		   num_tri = num_tri + 1
		end if
		
		
		if(CInt(num_tri) < CInt(Cant_trip)) then
			c_tr = "no" 
			err_cant_tri = 1
		End if
	End if 
	rstf.Close
	'Else
	'lst_pasap_pasaj = "("
	'lst_pasap_trip = pic_list & ","
	'acum2 = 0
	'tripCanti = 0
	'for i = 1 to num_tri_bd
	'	if Request.Form("tri_" & i) = "on" then
			'if Request.Form("tri_hid_" & i) = "" then
			'	err_pasap_trip = err_pasap_trip + 1
			'else
	'			lst_pasap_trip = lst_pasap_trip & Request.Form("tri_hid_" & i) & ","
	'			chk_trip = chk_trip & "tri_" & i & "|"
	'			acum2 = acum2 + 1
			'End if
	'	End if 
	'next
	'if chk_trip <> "" then
	'	chk_trip = left(chk_trip, len(chk_trip)-1)
	'End if
	
	
	
	
	'----------------HUMBERTO ROJAS 03/2015
	'response.write Request.form("selectripinput") & "</br>"
	'a=mid(Request.form("selectripinput"),1, (len(Request.form("selectripinput"))-1))
    if Request.form("selectripinput") <> "" then
		lst_pasap_trip = mid(Request.form("selectripinput"),1, (len(Request.form("selectripinput"))-1))
		chk_trip = replace(lst_pasap_trip,",","|")	
	end if	
	
	if lst_pasap_trip <> "" then	
		   acum2 = 1
	end if
	
'response.write Request.form("selecpasinput") & "</br>"
	if Request.form("selecpasinput")  <> "" then
		lst_pasap_pasaj = mid(Request.form("selecpasinput"),1, (len(Request.form("selecpasinput"))-1))
		chk_pasap = replace(lst_pasap_pasaj,",","|")
	end if
	
	if lst_pasap_pasaj <> "" then	
		   acum1 = 1
	end if
	
	'response.write  lst_pasap_pasaj
	'response.End()	
	'--------------------------------------
	
	Request.Form("pic_list")
	
	
    
	if tripCanti > 0 then
		
		cond = " Num_trip in (" & pic_list & "," & copic_list & ")"
		tripuLst = "&ctr=1&tr=(" & lst_pasap_trip & "," & copic_list & ")"	
	Else
		cond = " Num_trip = " & pic_list
		tripuLst = "&ctr=0&tr=" & pic_list
	End If
	if acum1 > 0 then
		pasajList = "&ps=" & lst_pasap_pasaj
	Else
		pasajList = "&ps=nn"
	End if
	
	
	''response.write Cod_pais_lle &" "& Cod_pais_sal
	'response.End()
	
	
	
	
	
	if Cod_pais_lle = "1" And Cod_pais_sal = "1" then
		if acum1 > 0 then
			rstf.Open  "Select Num_pasa, Nom_pasa, Ape_pasa, Num_cedu_pas, Num_pasa_pas, Pais_res_pas, Nac_pasaj, Fec_venc_pas, Num_cedu_pas From Siep_pasa Where Num_pasa in (" & lst_pasap_pasaj & ")", cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_pasa")) Or Len(rstf("Ape_pasa")) = 0 Or rstf("Ape_pasa") = "" then
					err_ape_pasaj = err_ape_pasaj + 1
				End If
				'///////18-02-2014 -> se hizo una exepción con esta validación
				'Pasajero extranjero
				'if rstf("Pais_res_pas") <> "1" then 
				'	if IsNull(rstf("Nac_pasaj")) Or Len(rstf("Nac_pasaj")) = 0 Or rstf("Nac_pasaj") = "" then
				'		err_nac_pasaj = err_nac_pasaj + 1
				'	End If
				'	if IsNull(rstf("Num_pasa_pas")) Or Len(rstf("Num_pasa_pas")) = 0 Or rstf("Num_pasa_pas") = "" then
				'		err_pasap_pasaj = err_pasap_pasaj + 1
				'		'/////////////////////////////////
				'		'Ya no se usa la restricción de la fecha próxima del vencimiento del 
				'		'pasaporte
				'		'////////////////////////////////
				'		'Else
				'		'	if Not IsNull(rstf("Fec_venc_pas")) then
				'		'		if DateDiff("d", Now(), rstf("Fec_venc_pas")) <= 30 then
				'		'			err_pasap_pasaj = err_pasap_pasaj + 1
				'		'		End if
				'		'	End if
				'	End If		
				'End If
				'Pasajero venezolano
				'if rstf("Pais_res_pas") = "1" then 	
				'	if IsNull(rstf("Num_cedu_pas")) Or Len(rstf("Num_cedu_pas")) = 0 Or rstf("Num_cedu_pas") = "" then
				'		err_ced_pasaj = err_ced_pasaj + 1
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
					err_ape_trip = err_ape_trip + 1
				End If
				if IsNull(rstf("Num_lic_tri")) then
					err_lic_cab = err_lic_cab + 1
				End if
				if IsNull(rstf("Fec_venc_lic")) then
					err_licv_cab = err_licv_cab + 1
				Else
					'Da error si el documento está vencido (Sept 2015) -> Cliente
					' Now() -> fec_salida
					'if DateDiff("d", Now(), rstf("Fec_venc_lic")) < 0 then
					if DateDiff("d", fec_salida, rstf("Fec_venc_lic")) < 0 then
						err_licv_cab = err_licv_cab + 1
					End if
				End if
				'Se valida si el certificado de Vzla está vencido
				'Irving -> 16-09-2015
				if IsNull(rstf("Num_cer_tri")) then
					err_cer_cab = err_cer_cab + 1
				Else
					if IsNull(rstf("Fec_venc_cer")) then
						err_cerv_cab = err_cerv_cab + 1
					Else
						'Da error si el documento está vencido
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_cer")) < 0 then
							err_cerv_cab = err_cerv_cab + 1
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
					err_ape_trip = err_ape_trip + 1
				End If
				if IsNull(rstf("Num_lic_eua_tri")) then
					err_licEua_cab = err_licEua_cab + 1
				End if
				if IsNull(rstf("Num_cer_eua_tri")) then
					err_cerEua_cab = err_cerEua_cab + 1
				Else
					'Da error si el documento está vencido (Sept 2015) -> Cliente
					' Now() -> fec_salida
					if DateDiff("d", fec_salida, vencCerEUA(rstf("Fec_exp_eua_cer"))) < 0 then
						err_cerEua_cab = err_cerEua_cab + 1
					End if
				End if
				rstf.MoveNext
			Loop
			rstf.Close
		End if
		'///////////////////////////////////
	End if
	
	'response.write "Select Num_pasa, Nom_pasa, Ape_pasa, Num_pasa_pas, Pais_res_pas, Nac_pasaj, Fec_venc_pas, Num_cedu_pas From Siep_pasa Where Num_pasa in " & lst_pasap_pasaj
	'response.End()
	'response.Write acum1 &" ***** "
	'response.write lst_pasap_pasaj
	'response.End()
	
	
	if Cod_pais_lle <> "1" Or Cod_pais_sal <> "1" then
		if acum1 > 0 then
			rstf.Open  "Select Num_pasa, Nom_pasa, Ape_pasa, Num_pasa_pas, Pais_res_pas, Nac_pasaj, Fec_venc_pas, Num_cedu_pas, Num_visa_pas, Fec_venc_vis From Siep_pasa Where Num_pasa in (" & lst_pasap_pasaj & ")", cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Nac_pasaj")) Or Len(rstf("Nac_pasaj")) = 0 Or rstf("Nac_pasaj") = "" then
					err_nac_pasaj = err_nac_pasaj + 1
				End If
				if IsNull(rstf("Ape_pasa")) Or Len(rstf("Ape_pasa")) = 0 Or rstf("Ape_pasa") = "" then
					err_ape_pasaj = err_ape_pasaj + 1
				End If
				if IsNull(rstf("Num_pasa_pas")) Or Len(rstf("Num_pasa_pas")) = 0 Or rstf("Num_pasa_pas") = "" then
					err_pasap_pasaj = err_pasap_pasaj + 1
				End If	
				if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
					err_pasap_pasaj = err_pasap_pasaj + 1
					'/////////////////////////////////
					'Ya no se usa la restricción de la fecha próxima del vencimiento del 
					'pasaporte
					'////////////////////////////////
					'//////////////////////////////
					'Else
					'	if DateDiff("d", Now(), rstf("Fec_venc_pas")) <= 30 then
					'		err_pasap_pasaj = err_pasap_pasaj + 1
					'	End if
					'//////////////////////////////
				End if
				
				'Visa EUA
				if Cod_pais_lle = "4" then
					if IsNull(rstf("Num_visa_pas")) then
						err_visa_pasaj = err_visa_pasaj + 1
					Else
						if IsNull(rstf("Fec_venc_vis")) then
							err_vvisa_pasaj =err_vvisa_pasaj+ 1
						Else
							'Da error si el documento está vencido
							' Now() -> fec_salida
							if DateDiff("d", fec_salida, rstf("Fec_venc_vis")) < 0 then
								err_vvisa_pasaj = err_vvisa_pasaj + 1
							End if
						End if
					End if
				end if
				
				'Paises del mercosur
				if Cod_pais_lle = "1" Or Cod_pais_sal = "1" then	
					if ((mercosur(Cod_pais_lle) = 1) Or (mercosur(Cod_pais_sal) = 1)) then
						err_pasap_pasaj = 0	
					End if
				End if
				'if rstf("Pais_res_pas") = "1" then 	
				'	if IsNull(rstf("Num_cedu_pas")) Or Len(rstf("Num_cedu_pas")) = 0 Or rstf("Num_cedu_pas") = "" then
				'		err_ced_pasaj = err_ced_pasaj + 1
				'	End If
				'End if
				rstf.MoveNext
			Loop
			rstf.Close
		End if
				
		if Mid(aeronave, 1, 2) = "YV" then
			rstf.Open  "Select Num_cedu_tri, Nom_trip, Ape_trip, Num_pasa_tri, Fec_venc_pas, Nac_tri, Fec_nac_tri, Num_lic_tri, Num_lic_tri, Num_cer_tri, Fec_venc_lic, Fec_venc_cer, Fec_venc_pas From Siep_trip Where " & cond, cnnf, 1, 2
			Do Until rstf.EOF
				if IsNull(rstf("Ape_trip")) Or Len(rstf("Ape_trip")) = 0 Or rstf("Ape_trip") = "" then
					err_ape_trip = err_ape_trip + 1
				End If
				if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
					err_pasap_trip = err_pasap_trip + 1
				Else
					'if DateDiff("d", Now(), rstf("Fec_venc_pas")) <= 30 then
					'Da error si el documento está vencido (Sept 2015) -> Cliente
					' Now() -> fec_salida
					if DateDiff("d", fec_salida, rstf("Fec_venc_pas")) < 0 then
						err_pasap_trip = err_pasap_trip + 1
					End if
				End if
				if IsNull(rstf("Num_lic_tri")) then
					err_lic_cab = err_lic_cab + 1
				Else
					if IsNull(rstf("Fec_venc_lic")) then
						err_licv_cab = err_licv_cab + 1
					Else
						'Da error si el documento está vencido
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_lic")) < 0 then
							err_licv_cab = err_licv_cab + 1
						End if
					End if
				End if
				'Se valida si el certificado de Vzla está vencido
				'Irving -> 16-09-2015
				if IsNull(rstf("Num_cer_tri")) then
					err_cer_cab = err_cer_cab + 1
				Else
					if IsNull(rstf("Fec_venc_cer")) then
						err_cerv_cab = err_cerv_cab + 1
					Else
						'Da error si el documento está vencido
						' Now() -> fec_salida
						if DateDiff("d", fec_salida, rstf("Fec_venc_cer")) < 0 then
							err_cerv_cab = err_cerv_cab + 1
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
					err_ape_trip = err_ape_trip + 1
				End If
				'Response.Write " " & IsNull(rstf("Fec_venc_pas")) & " " & Len(rstf("Fec_venc_pas")) = 0 & " " & rstf("Fec_venc_pas")
				if IsNull(rstf("Fec_venc_pas")) Or Len(rstf("Fec_venc_pas")) = 0 Or rstf("Fec_venc_pas") = "" then
					err_pasap_trip = err_pasap_trip + 1
				Else
					'if DateDiff("d", Now(), rstf("Fec_venc_pas")) <= 30 then
					'Da error si el documento está vencido (Sept 2015) -> Cliente
					' Now() -> fec_salida
					if DateDiff("d", fec_salida, rstf("Fec_venc_pas")) < 0 then
						err_pasap_trip = err_pasap_trip + 1
					End if
				End if
				if IsNull(rstf("Num_lic_eua_tri")) then
					err_licEua_cab = err_licEua_cab + 1
				End if
				if IsNull(rstf("Num_cer_eua_tri")) then
					err_cerEua_cab = err_cerEua_cab + 1
				Else
					'Da error si el documento está vencido (Sept 2015) -> Cliente
					' Now() -> fec_salida
					if DateDiff("d", fec_salida, vencCerEUA(rstf("Fec_exp_eua_cer"))) < 0 then
						err_cerEua_cab = err_cerEua_cab + 1
					End if
				End if
				'Se valida la visa si el destino es EEUU
				'Irving -> 16-09-2015 
				if Cod_pais_lle = "4" then
					if IsNull(rstf("Num_visa_tri")) then
						err_visa_cab = err_visa_cab + 1
					Else
						if IsNull(rstf("Fec_venc_vis")) then
							err_visav_cab =err_visav_cab+ 1
						Else
							'Da error si el documento está vencido
							' Now() -> fec_salida
							if DateDiff("d", fec_salida, rstf("Fec_venc_vis")) < 0 then
								err_visav_cab = err_visav_cab + 1
							End if
						End if
					End if
				end if
				rstf.MoveNext
			Loop
			rstf.Close
		End if		
	End if
	'End if
	
	
	'response.write err_ape_pasaj & "</br>"
	'response.write err_ape_trip & "</br>"
	'response.write err_nac_pasaj & "</br>"
	'response.write err_ced_pasaj & "</br>"
	'response.write err_pasap_pasaj & "</br>"
	'response.write err_pasap_trip & "</br>"
	'response.write err_licv_cab & "</br>"
	'response.write err_visa_cab & "</br>"
	'response.write err_cerEua_cab & "</br>"
	'response.write err_licEua_cab & "</br>"
	'response.write err_cant_tri & "</br>"
    'response.End()
	
	'response.write  lst_pasap_pasaj
	'response.End()	
	
	'if (err_pasap_pasaj + err_pasap_trip + err_licv_cab + err_licEua_cab + err_visa_cab = 0) then
	if (err_ape_pasaj + err_ape_trip + err_nac_pasaj + err_ced_pasaj + err_pasap_pasaj + err_pasap_trip + err_licv_cab + err_visa_cab + err_cerEua_cab + err_licEua_cab + err_cant_tri + err_cer_cab + err_cerv_cab + err_visa_cab + err_visav_cab + err_visa_cab + err_visav_cab = 0) then
	
	   'response.write "modi " & Request.Form("modifica")
	  ' response.End()
	
	
		if modi = "-" then
			rstf.Open  "Siep_mani", cnnf, 1, 2
			rstf.AddNew
			rstf("Fec_clie_reg") = Session("Fec_clie_reg")
			rstf("Num_clie") = Session("Num_clie")
		else
			rstf.Open  "Select * from Siep_mani Where Num_mani = " & modi, cnnf, 1, 2
		End if 
		rstf("Cod_aeron") = aeronave
		rstf("Cod_aero_sal") = Cod_aero_sal
		rstf("Cod_aero_lle") = Cod_aero_lle
		if hor_salida <> "" then
			rstf("Fec_vuel_sal") = fec_salida & " " & hor_salida
			rstf("Hor_vuel_sal") = hor_salida
		else
			rstf("Fec_vuel_sal") = fec_salida
			rstf("Hor_vuel_sal") = null
		End if
		if fec_llegada="" then
			fec_llegada = null
		end if
		'rstf("Fec_vuel_lle") = fec_llegada & " " & hor_llegada
		rstf("Fec_vuel_lle") = fec_llegada
		rstf.Update
		num_mani = rstf("Num_mani")
		rstf.Close
		
		'response.Write num_mani
		'response.End()
		
		
		if modi <> "-" then
			Call elimina("Siep_mapa", "Num_mani = " & modi)
			Call elimina("Siep_matr", "Num_mani = " & modi)
		End if
		
		'Pasajeros  (manifiesto)'
		'for i = 1 to num_pas_bd
			'if Request.Form("pas_" & i) = "on" then
			'	rstf.Open  "Siep_mapa", cnnf, 1, 2
			'	rstf.AddNew
			'	rstf("Fec_clie_reg") = Session("Fec_clie_reg")
			'	rstf("Num_clie") = Session("Num_clie")
			'	rstf("Num_mani") = num_mani
			'	rstf("Num_pasa") = Request.Form("pas_hid_" & i)
			'	rstf.Update
			'	rstf.Close
			'end if
		'next
		
		'Pasajeros (NUEVA CAPTURA DESARROLLADA POR HUMBERTO ROJAS 2015)
		'RESPONSE.WRITE Request.Form("selecpas") 
		'response.End()
			
		
		Dim pasajeros1(), I 
		I = 0 
		lst_pasap_pasaj2 = ""
		Arraypasajeros =Split(lst_pasap_pasaj,",")
		For Each Valor In Arraypasajeros 
		
			Redim Preserve pasajeros1(I) 
			pasajeros1(I) = Valor 
			
			    rstf.Open  "Siep_mapa", cnnf, 1, 2
				rstf.AddNew
				rstf("Fec_clie_reg") = Session("Fec_clie_reg")
				rstf("Num_clie") = Session("Num_clie")
				rstf("Num_mani") = num_mani
				rstf("Num_pasa") = pasajeros1(I)
				rstf.Update
				rstf.Close
				
				lst_pasap_pasaj2 = lst_pasap_pasaj2 & pasajeros1(I) & ","			
			
			I = I + 1 
		Next  
		
		ps = lst_pasap_pasaj2
		chk_pasap = I + 1
		
		'response.Write "pasajeros " & ps
		'response.End()
		
		
		'***************************************************************
		
		'Inserta el piloto (manifiesto)'
		rstf.Open  "Siep_matr", cnnf, 1, 2
		rstf.AddNew
		rstf("Fec_clie_reg") = Session("Fec_clie_reg")
		rstf("Num_clie") = Session("Num_clie")
		rstf("Num_mani") = num_mani
		rstf("Num_trip") = pic_list
		rstf("Pic") = "1"
		rstf.Update
		rstf.Close
		if copic_list <> "0" then
			rstf.Open  "Siep_matr", cnnf, 1, 2
			rstf.AddNew
			rstf("Fec_clie_reg") = Session("Fec_clie_reg")
			rstf("Num_clie") = Session("Num_clie")
			rstf("Num_mani") = num_mani
			rstf("Num_trip") = copic_list
			rstf("Copic") = "1"
			rstf.Update
			rstf.Close
		End if
		
		
		'Tripulacion (NUEVA CAPTURA DESARROLLADA POR HUMBERTO ROJAS 2015)
			
				
		Dim tripulacion(), I2 
		I2 = 0 
		lst_tri = ""
		Arraytripulacion =Split(lst_pasap_trip,",")
		For Each Valor In Arraytripulacion 
			Redim Preserve tripulacion(I2) 
			tripulacion(I2) = Valor 
			
			    rstf.Open  "Siep_matr", cnnf, 1, 2
				rstf.AddNew
				rstf("Fec_clie_reg") = Session("Fec_clie_reg")
				rstf("Num_clie") = Session("Num_clie")
				rstf("Num_mani") = num_mani
				rstf("Num_trip") = tripulacion(I2)
				rstf.Update
				rstf.Close
						
			    lst_tri = lst_tri & tripulacion(I2) & ","	
			I2 = I2 + 1 
		Next  
		
		tr = lst_tri 
		chk_trip = I2 + 1
		ctr = I2 + 1
		'***************************************************************
		
		'response.Write "modi " & modi
		
		direccion = ""
		if modi = "-" then
			if CInt(num_pas) = 0 then
				if Tabla_Vacia("Avion", " Cod_avion = '" & aeronave & "'") = false then
					direccion1 = "registropierna.asp?v=" & num_mani
					direccion2 = "finregistromanifiesto.asp?v=" & num_mani
					Call confirmar("¿Desea registrar este vuelo en piernas libres?",direccion1, direccion2)
				else
					direccion = "finregistromanifiesto.asp?v=" & num_mani & "&p=" & num_pas & "&t=" & num_tri
				End if
			else
				direccion = "finregistromanifiesto.asp?v=" & num_mani & "&p=" & num_pas & "&t=" & num_tri
			End if 
		else
			direccion = "finregistromanifiesto.asp?v=" & num_mani & "&p=" & num_pas & "&t=" & num_tri   
		End if
	else
		'msgBoxWarn = ". Revisar: "
		'If err_pasap_pasaj > 0 then
		'	msgBoxWarn = msgBoxWarn & "documento de identificación (tomar en cuenta C.I. Venezolana) de los pasajeros,"
		'End if
		'If err_pasap_trip > 0 then
		'	msgBoxWarn = msgBoxWarn &  " pasaporte de los tripulantes,"
		'End if
		''Piloto y copiloto
		''If err_visa_cab > 0 then
		''	msgBoxWarn = msgBoxWarn &  " visa de pilotos y copilotos,"	
		''End if
		'If err_licv_cab > 0 then
		'	msgBoxWarn = msgBoxWarn &  " licencia de pilotos y copilotos,"
		'End If
		'If err_cerEua_cab > 0 then
		'	msgBoxWarn = msgBoxWarn &  " certificado de EEUU de pilotos y copilotos,"	
		'End If
		''30-01-2014 Lic EEUU no tiene fecha de vencimiento
		'If err_licEua_cab > 0 then
		'	msgBoxWarn = msgBoxWarn &  " licencia de EUA de pilotos y copilotos,"
		'End if
		'If err_ced_cab > 0 then
		'	msgBoxWarn = msgBoxWarn &  " cédula de identidad de pilotos y copilotos "
		'End if
		'msgBoxWarn = Mid(msgBoxWarn, 1, Len(msgBoxWarn) -1)
		'Call alerta("Los pasajeros y/o tripulación en ciertas condiciones no pueden tener documentos vencidos, por vencerse o sin fecha de vencimiento en un vuelo internacional o nacional " & msgBoxWarn)
		
		Call alerta("El manifiesto no puede ser registrado ")
		direccion = "correcmanifiesto.asp?bck=1&r=" & regreso & "&c1=" & chk_pasap & "&c2=" & chk_trip & tripuLst & pasajList & "&c_tr=" & c_tr & "&v=" & modi & "&num_per_tot= " & num_per_tot & "&dat=" & fec_salida & "&selecpasinput= " & selecpasinput & "&selectripinput=" & selectripinput & "&canttrip= " & canttrip & "&cantpas=" & cantpas
	End if
	'/////////////////////////'
	cnnf.Close
	Set rstf = nothing
	Set cnnf = nothing  
	if Err.Number <> 0 then
		Response.Write Err.Description
		Error.Clear
		Response.End()
	End if
	if direccion <> "" then
		call redir(direccion)
	end if
	'Response.Redirect(direccion)'
%>
