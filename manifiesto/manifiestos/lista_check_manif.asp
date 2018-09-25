<!--#include file="../modulos/funcion.asp"-->
<!--#include file="../modulos/control.asp"-->
<%

   	'HUMBERTO UPDATE 2018    
	filtrobusqueda = request.form("busquedaPas")
	filtrobusquedaTrip = request.form("busquedaTrip")
    
	
	vertrip = request.form("vertrip")
    verpas = request.form("verpas")
	
	'response.write "vertrip " & vertrip
	'response.write "verpas " & verpas
	
	
	'if vertrip = 1 then
	 '  verpas = 0
	'else 
	 '  vertrip = 0
	'end if   
	
	pilotoViene = request.form("pic_list")
	copilotoViene = request.form("copic_list")
	
    'response.write "pilotoViene " & pilotoViene 
    'response.write "copilotoViene " & copilotoViene

   'FIN FUMBERTO 2018
	num_pas_bd = 0
	num_tri_bd = 0
	bck = Request("bck")
	LstNumPasaMarc = ""
	LstNumTripMarc = ""
	bsq = ""
	'**************************************************
	Sub tablaConCheck(cadSql, color)	
		rstx.Open cadSql, cnnx, 1, 2	
		num_pas_bd = rstx.RecordCount
		if rstx.EOF then
			list_pas = list_pas & "<tr><td align='left'><span class='Estilo4'>No hay pasajeros disponibles</span></td></tr>"
		else
			Do Until rstx.EOF
				fec_venc_pas = "" 
				fec_venc_vis = ""
				disab = ""
				if Not IsNull(rstx("Fec_venc_pas")) then 
					if  DateDiff("d", Now(), rstx("Fec_venc_pas"))< 0 then
						fec_venc_pas =  " - <span class='Estilo4'>Pasaporte vencido</span>"
					End if
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) <= 30 And DateDiff("d", Now(), rstx("Fec_venc_pas")) >= 0 then
						fec_venc_pas =  " - <span style='color:#FFBF00'>Pasaporte por vencerse</span>"
					End if
					'///////////////////////////////////
					'Este caso de pasaporte por vencerse no se usa
					'///////////////////////////////////
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) > 30  then
						disab = ""
					End if
				else
					fec_venc_pas = " - <span class='Estilo4'>Pasaporte sin fecha de vencimiento</span>"
					disab = ""
				End if
				if Not IsNUll(rstx("Num_visa_pas")) then
					if Not IsNull(rstx("Fec_venc_vis")) then 
						disab = ""
						if  DateDiff("d", Now(), rstx("Fec_venc_vis"))< 0 then
							fec_venc_vis = " - <span class='Estilo4'><b>Visa vencida</b></span>"
						End if
						if  DateDiff("d", Now(), rstx("Fec_venc_vis")) <= 30 And DateDiff("d", Now(), rstx("Fec_venc_pas")) >= 0 then
							fec_venc_vis = " - <span style='color:#FFBF00'><b>Visa por vencerse</b></span>"
						End if
						if  DateDiff("d", Now(), rstx("Fec_venc_vis")) > 30  then
							disab = ""
						End if
					else
						fec_venc_vis = " - <span class='Estilo4'>Visa sin fecha de vencimiento</span>"
						disab = ""
					End if
				Else
					fec_venc_vis = " - <span class='Estilo4'>Sin visa de EEUU</span>"
				End if
				'if bck = "1" then' 
					'chk_pasaj = lista_checbox("pas_" & indi,c_pasap)' 
				'else'
					'chk_pasaj = ""'
				'End if'
				
				if modimanifiesto = true then
					chequeado = ""
					if pasajero <> "" then
						for k = 0 to UBound(list_pasajero)
							if Trim(list_pasajero(k)) = Trim(rstx("Num_pasa")) then 
								chequeado = "checked"
							End if
						next
					End if
				End if
				
				if(Request("t") <> "") then
					if(Request.Form("pas_" & indi) = "on") then
						LstNumPasaMarc = LstNumPasaMarc & Request.Form("pas_hid_" & indi) & "|"
					End if
				End if
				
				if rstx("Sta_acti") = "False" then
					disab = "disabled"
				End if 
				
				if rstx("Sta_acti") = "False" then
					disab = "disabled"
				End if
				if modimanifiesto = true then	
					list_pas = list_pas & "<tr><td align='left'  " & color & "><span class='Estilo3'>&nbsp;" & check_box_disab("pas_" & indi, "pas_" & indi, chequeado, disab, "OnClick='chk_pasj(" & """pas_" & indi & """)'") & "&nbsp;" & rstx("Nom_pasa") & "&nbsp;" & rstx("Ape_pasa") & "&nbsp;" & fec_venc_pas & "&nbsp;" & fec_venc_vis & " " & hidden("pas_hid_" & indi, rstx("Num_pasa")) & hidden("fec_nac_pasa_" & indi, rstx("Fec_nac_pas")) & "</span></td></tr>"
				else
					elchk = ""
					if bck <> "" then
						'elchk = lista_checbox("pas_" & indi,c_pasap)
						elchk = lista_checbox(rstx("Num_pasa"),c_pasap)	
					else
						elchk = chequeadoReg
					End if
					list_pas = list_pas & "<tr><td align='left'  " & color & "><span class='Estilo3'>&nbsp;" & check_box_disab("pas_" & indi, "pas_" & indi, elchk, disab, "OnClick='chk_pasj(" & """pas_" & indi & """)'") & "&nbsp;" & rstx("Nom_pasa") & "&nbsp;" & rstx("Ape_pasa") & "&nbsp" & fec_venc_pas & "&nbsp;" & fec_venc_vis & " " & hidden("pas_hid_" & indi, rstx("Num_pasa")) & hidden("fec_nac_pasa_" & indi, rstx("Fec_nac_pas")) & "</span></td></tr>"
				End if
				rstx.MoveNext
				indi = indi + 1
			Loop
		End if
		rstx.Close
	End Sub
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	Sub tablaConCheckTri()
		'Tripulación con checkbox'
		'rstx.Open "Select Num_trip, Nom_trip, Ape_trip, Num_visa_tri, Fec_venc_pas, Fec_venc_vis, Fec_venc_eua_lic, Fec_venc_lic, Fec_venc_cer From Siep_trip Where Num_clie = " & Session("Num_clie") & " And  (Sta_acti = true) Order By Nom_trip", cnnx, 1, 2
		rstx.Open "Select Num_trip, Nom_trip,Ape_trip, Sta_acti, Num_visa_tri, Fec_venc_pas, Fec_venc_vis, Fec_venc_lic, Fec_venc_cer, Fec_exp_eua_cer From Siep_trip Where Num_clie = " & Session("Num_clie") & " And  (Sta_acti = true) Order By Nom_trip", cnnx, 1, 2
		num_tri_bd = rstx.RecordCount
		if rstx.EOF then
			list_tri = list_pas & "<tr><td align='left'><span class='Estilo4'>No hay tripulacion disponible</span></td></tr>"
		else
			Do Until rstx.EOF
				fec_venc_pas_tri = "" 
				fec_venc_vis_tri = ""
				fec_venc_lic_tri = ""
				'fec_venc_lic_eua_tri = ""
				fec_venc_cer_eua_tri = ""
				if Not IsNull(rstx("Fec_venc_pas")) then 
					disabl = ""
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) < 0 then
						fec_venc_pas_tri =  " - <span class='Estilo4'>Pasaporte vencido</span>"
					End if
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) <= 30 And DateDiff("d",  Now(), rstx("Fec_venc_pas")) >= 0 then
						fec_venc_pas_tri =  " - <span style='color:#FFBF00'>Pasaporte por vencerse</span>"
						disabl = ""
					End if
					'if  DateDiff("d",  Now(), rstx("Fec_venc_pas")) > 30  then
					'	disabl = ""
					'End if
				else
					fec_venc_pas_tri = " - <span class='Estilo4'>Pasaporte sin fecha de vencimiento</span>"
					'disabl = "disabled"   And (DateDiff('d',now(),Fec_venc_eua_lic)>0) And '
				End if
				if Not IsNull(rstx("Num_visa_tri")) then
					if Not IsNull(rstx("Fec_venc_vis")) then 
						disabl = ""
						if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) < 0 then
							fec_venc_vis_tri = " - <span class='Estilo4'>Visa vencida</span>"
						End if
						if  DateDiff("d", Now(), rstx("Fec_venc_vis")) < 30 And DateDiff("d",Now(), rstx("Fec_venc_vis")) >= 0  then
							fec_venc_vis_tri = " - <span style='color:#FFBF00'>Visa por vencerse</span>"
						End if
						
						'if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) > 30 then
						'	disabl = ""
						'End if
					else
						disabl = ""
						fec_venc_vis_tri = " - <span class='Estilo4'>Visa sin fecha de vencimiento</span>"
					End if
				else
					fec_venc_vis_tri = " - <span class='Estilo4'>Sin visa de EEUU</span>"
				End if
				if Not IsNull(rstx("Fec_venc_lic")) then
					disabl = ""
					if  DateDiff("d",  Now(), rstx("Fec_venc_lic")) < 0 then
						fec_venc_lic_tri = " - <span class='Estilo4'>Licencia vencida</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_lic_tri = " - <span class='Estilo4'>Licencia sin fecha de vencimiento</span>"
				End if
				'////////////////Nuevo////////////////////////
				if Not IsNull(rstx("Fec_exp_eua_cer")) then
					disabl = ""
					if  DateDiff("d",  Now(), vencCerEUA(rstx("Fec_exp_eua_cer"))) < 0 then
						fec_venc_cer_eua_tri = " - <span class='Estilo4'>Certificado de EUA vencido</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_cer_eua_tri = " - <span class='Estilo4'>Certificado de EUA sin fecha de vencimiento</span>"
				End if
				'/////////////////////////////////////////////
				'30/01/2014 -> La licencia de EUA no lleva fecha
				'if Not IsNull(rstx("Fec_venc_eua_lic")) then
				'	if  DateDiff("d",  Now(), rstx("Fec_venc_eua_lic")) <= 0 then
				'		fec_venc_lic_eua_tri = " - <span class='Estilo4'>Licencia de EEUU vencida</span>"
				'	End if
				'else
				'	fec_venc_lic_eua_tri = " - <span class='Estilo4'>Licencia de EEUU sin fecha de vencimiento</span>"
				'End if
				if Not IsNull(rstx("Fec_venc_cer")) then
					disabl = ""
					if  DateDiff("d",  Now(), rstx("Fec_venc_cer")) < 0 then
						fec_venc_cer_tri = " - <span class='Estilo4'>Certificado vencido</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_cer_tri = " - <span class='Estilo4'>Certificado sin fecha de vencimiento</span>"
				End if
				
				if modimanifiesto = true then		
					chequeado_tri = ""
					if tripulacion <> "" then
						for k = 0 to UBound(list_tripulacion)
							'if list_tripulacion(k) = rstx("Nom_trip") & " " & rstx("Ape_trip") then
							if Trim(list_tripulacion(k)) = Trim(rstx("Num_trip")) then
								chequeado_tri = "checked"
							End if
						next
					End if
				End if
				if(Request("t") <> "") then
					if(Request.Form("tri_" & indi) = "on") then
						LstNumTripMarc = LstNumTripMarc & Request.Form("tri_hid_" & indi) & "|"
					End if
				End if
				
				if rstx("Sta_acti") = "False" then
					disab = "disabled"
				End if 
				if modimanifiesto = true then	
					list_tri = list_tri & "<tr><td align='left'><span class='Estilo3'>&nbsp;" & check_box_disab("tri_" & indi, "tri_" & indi, chequeado_tri, disabl, "OnClick='chk_trip(""" & "tri_" & indi & """)'") & "&nbsp;" & rstx("Nom_trip") & "&nbsp;" & rstx("Ape_trip") & "&nbsp;" & fec_venc_pas_tri & "&nbsp;" & fec_venc_vis_tri & fec_venc_lic_tri & fec_venc_lic_eua_tri & fec_venc_cer_eua_tri & "&nbsp;" & " " & hidden("tri_hid_" & indi, rstx("Num_trip")) & "</span></td></tr>"
				Else	
					list_tri = list_tri & "<tr><td align='left'><span class='Estilo3'>&nbsp;" & check_box_disab("tri_" & indi, "tri_" & indi, lista_checbox("tri_" & indi,c_trip), disabl, "OnClick='chk_trip(""" & "tri_" & indi & """)'") & "&nbsp;" & rstx("Nom_trip") & "&nbsp;" & rstx("Ape_trip") & "&nbsp" & fec_venc_pas_tri & "&nbsp;" & fec_venc_vis_tri & fec_venc_lic_tri & fec_venc_lic_eua_tri & fec_venc_cer_eua_tri & "&nbsp;" & " " & hidden("tri_hid_" & indi, rstx("Num_trip")) & "</span></td></tr>"
				End if
				rstx.MoveNext
				indi = indi + 1
			Loop
		End if
	End Sub
	
	'**************************************************
	
	busqueda = Request("t")
	resultBusq = Request.Form("camp_busc")
	
	'Esto solo se ejecuta
	'cuando se efectúa la
	'búsqueda
	if (resultBusq <> "") then
		rstx.Open "Select Num_pasa From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) And ((Nom_pasa Like '%" & resultBusq & "%') Or (Ape_pasa Like '%" & resultBusq & "%')) Order By Nom_pasa", cnnx, 1, 2
		if not rstx.EOF then
				remarcar_busc = 1
				id_pasa_lst = ""
				Do Until rstx.EOF
					id_pasa_lst = id_pasa_lst & rstx("Num_pasa") & ","
					rstx.MoveNext
				Loop
				id_pasa_lst = Left(id_pasa_lst,len(id_pasa_lst)-1)
				id_pasa_lst = "(" & id_pasa_lst & ")"		
		end if
		rstx.Close
		bsq = "1"
	End if

	'[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-
	
	''''''''''''''''''''''''''''''''''''''
	'//////////////////////////////////////
	'Esta porción de código
	'fué ordenada para que esté
	'en un solo archivo y no
	'en dos.
	'//////////////////////////////////////
	
	sqlStr = ""
	codigoAeron = ""
	aeroPtoOrigen = ""
	fechaSalidax = ""
	aeroPtoLlegadax = ""
	picListax = ""
	almasPersonax = ""
	numerPerTotx = ""
	cantPasajx = ""
	cantTripx = ""
		
	if modimanifiesto = true then
		sqlStr = "Select Sig_aero, Ntr_aero From Siep_aero Where Num_clie = " & Session("Num_clie") & " Order by Sig_aero"
		codigoAeron = Cod_aero
		aeroPtoOrigen = Cod_aero_sal & "|" & aero_sal
		fechaSalidax = FormatDateTime(Fec_vuel_sal,2)
		aeroPtoLlegadax = Cod_aero_lle & "|" & aero_lle
		picListax = piloto
		almasPersonax = personas
		numerPerTotx = personas
		cantPasajx = cant_pas
		cantTripx = cant_trip 
	else
		sqlStr = "Select Sig_aero, Ntr_aero From Siep_aero Where Num_clie = " & Session("Num_clie") & " And  Sta_acti = true"  & " Order by Sig_aero"
		codigoAeron = aeronaver
		aeroPtoOrigen = aerop_origr
		fechaSalidax = fec_salidar
		aeroPtoLlegadax = aerop_llegr
		picListax = pic_listr
		almasPersonax = almas
		numerPerTotx = tot_per
		cantPasajx = num_pasr
		cantTripx = tripu
	End if
	
	if (resultBusq <> "") then
		codigoAeron = Request.Form("aeronave")
		aeroPtoOrigen = Request.Form("aerop_orig")
		sig_aerop1 = Request.Form("sig_1")
		nom_aerop1 = Request.Form("nom_1")
		fechaSalidax = Request.Form("fec_salida")
		aeroPtoLlegadax = Request.Form("aerop_lleg")
		tiempoHora = Request.Form("hor_salida")
		tiempoMin = Request.Form("min_salida")
		sig_aerop2 = Request.Form("sig_2")
		nom_aerop2 = Request.Form("nom_2")
		picListax = Request.Form("pic_list")
		copic_listr = Request.Form("copic_list")
		copic_list_revr = Request.Form("copic_list_rev")
		almasPersonax =  Request.Form("num_per")
		numerPerTotx = Request.Form("num_per_tot")
		cantPasajx = Request.Form("num_pas")
		cantTripx = Request.Form("num_tri")
	End if
		
	'Selección de aeronaves'
	rstx.Open sqlStr, cnnx, 1, 2
	num = 0
	items1 = ""
	items2 = ""
	thidden_avion_trip = ""
	if not rstx.EOF then	
		if rstx.RecordCount = 1 then
			items1 = rstx("Sig_aero") & "|"
			items2 = rstx("Sig_aero") & "|"
			thidden_avion_trip = hidden("trip_1", rstx("Sig_aero") & "_" & rstx("Ntr_aero"))
			nums = 1
		else
			do until rstx.EOF
				nums = nums + 1
				items1 = items1 & rstx("Sig_aero") & "|"
				items2 = items2 & rstx("Sig_aero") & "|"
				thidden_avion_trip = thidden_avion_trip & hidden("trip_" & nums, rstx("Sig_aero") & "_" & rstx("Ntr_aero"))
				rstx.MoveNext
			loop 
			items1 = Left(items1,len(items1)-1)
			items2 = Left(items2,len(items2)-1)
		End if
	else
		items1 = "0|"
		items2 = "Sin datos..|"
	End if
	thidden_avion_trip = thidden_avion_trip & hidden("num_reg_trip", nums)
	rstx.Close
	
	codigoAeronViene = Request.Form("aeronave")
	if codigoAeronViene <> "" then
	   aeronav = seleccion2("aeronave", "aeronave", items1, items2, codigoAeronViene, "")
	else
      aeronav = seleccion2("aeronave", "aeronave", items1, items2, codigoAeron, "")
    end if	

	aeroOriViene = Request.Form("aerop_orig")
	if aeroOriViene <> "" then
	  aerop_orig = campo("aerop_orig", "text", "aerop_orig", "70", "60", "1", aeroOriViene, "OnBlur='caractNoPermit(this)'")
	else
	  aerop_orig = campo("aerop_orig", "text", "aerop_orig", "70", "60", "1", aeroPtoOrigen, "OnBlur='caractNoPermit(this)'")

    end if	
	ast_aerop_orig = tag("span", "ast_aerop_origen", "visibility:hidden", marca_error(), "")
	
	estiloTextBox = "background-color:transparent;font-size:14pt;"
	
	'<input type="text" readonly="readonly" id="sig_1" name="sig_1" class="nobord2" style="background-color:transparent;">
	sig_1 = campo_estilo_readonly("sig_1", "text", "sig_1", "nobord2", estiloTextBox, "10", "5", "2", sig_aerop1, "readonly", "")
	'<input type="text" readonly="readonly" id="nom_1" name="nom_1" class="nobord2" style="background-color:transparent;">
	nom_1 = campo_estilo_readonly("nom_1", "text", "nom_1", "nobord2", estiloTextBox, "70", "5", "3", nom_aerop1, "readonly", "")

	'<input name="fec_venc" type="text" id="fec_venc" size="20" maxlength="20" />'
	'campo(name, tipo, id, size, maxlength, tabindex, valor, exprJs)'
	fechaSalidaViene = Request.Form("fec_salida")
	if fechaSalidaViene <> "" then 
	   fec_salida = campo_readonly("fec_salida", "text", "fec_salida", "20", "20", "4", fechaSalidaViene, "readonly", "")
	else
       fec_salida = campo_readonly("fec_salida", "text", "fec_salida", "20", "20", "4", fechaSalidax, "readonly", "")
    end if
	ast_fec_salida = tag("span", "ast_fec_salida", "visibility:hidden", marca_error(), "")
	
	itemsh1 = "--|00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23"
	itemsh2 = "--|00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23"
	
	itemsm1 = "--|00"
	itemsm2 = "--|00"

	for i = 1 to 59
		if (i < 10) then
			itemsm1 = itemsm1 & "|0" & CStr(i)
			itemsm2 = itemsm2 & "|0" & CStr(i)
		else
			itemsm1 = itemsm1 & "|" & CStr(i)
			itemsm2 = itemsm2 & "|" & CStr(i)	
		end if
	next
	
	'itemsampm1 = "--|AM|PM"
	'itemsampm2 = "--|AM|PM"
	
	If modimanifiesto = true then
		if Hor_vuel_sal = "" then
			hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, "--", "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, "--", "") '& " " & seleccion2("am_pm_sal", "am_pm_sal", itemsampm1, itemsampm2, "--", "") 
		else
			  hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, formato_24_comb(FormatDateTime(Fec_vuel_sal, 3),1), "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, formato_24_comb(FormatDateTime(Fec_vuel_sal, 3),2), "")
		End if
	else
	   horaViene = Request.Form("hor_salida")
	   munitoViene = Request.Form("min_salida")
	   if horaViene <> "" then
		  'hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, formato_24_comb(FormatDateTime(horaViene, 3),1), "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, formato_24_comb(FormatDateTime(munitoViene, 3),2), "")
		  hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, horaViene, "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, munitoViene, "") '& " " & seleccion2("am_pm_sal", "am_pm_sal", itemsampm1, itemsampm2, AMPM, "") 
		else
		  'hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, formato_24_comb(FormatDateTime(Fec_vuel_sal, 3),1), "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, formato_24_comb(FormatDateTime(Fec_vuel_sal, 3),2), "")
		  hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, tiempoHora, "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, tiempoMin, "") '& " " & seleccion2("am_pm_sal", "am_pm_sal", itemsampm1, itemsampm2, AMPM, "") 
		end if
		'hor_salida = seleccion2("hor_salida", "hor_salida", itemsh1, itemsh2, tiempoHora, "") & " : " & seleccion2("min_salida", "min_salida", itemsm1, itemsm2, tiempoMin, "") '& " " & seleccion2("am_pm_sal", "am_pm_sal", itemsampm1, itemsampm2, AMPM, "") 
	End If
	'<a href="javascript:NewCssCal('fec_venc', 'ddMMyyyy')">'
	'<img src="../images_cal/cal.gif" width="16" height="16" alt="Pick a date" border="0"></a>'
	imgcal = img("../images_cal/cal.gif", "16", "16", "Pick a date", "", "")

	cal = enlace("", "javascript:NewCssCal(""fec_salida"", ""ddMMyyyy"");", "", "", imgcal, "")
	
	aeroLlegViene = Request.Form("aerop_lleg")
	if aeroLlegViene <> "" then
	   aerop_lleg = campo("aerop_lleg", "text", "aerop_lleg", "70", "60", "5", aeroLlegViene, "OnBlur='caractNoPermit(this)'")
	else
	    aerop_lleg = campo("aerop_lleg", "text", "aerop_lleg", "70", "60", "5", aeroPtoLlegadax, "OnBlur='caractNoPermit(this)'")
	end if
	
	
	ast_aerop_lleg = tag("span", "ast_aerop_lleg", "visibility:hidden", marca_error(), "")
	
	'<input type="text" readonly="readonly" id="sig_2" name="sig_2" class="nobord2" style="background-color:transparent;">
	sig_2 = campo_estilo_readonly("sig_2", "text", "sig_2", "nobord2", estiloTextBox, "10", "5", "7", sig_aerop2, "readonly", "")
	'<input type="text" readonly="readonly" id="nom_2" name="nom_2" class="nobord2" style="background-color:transparent;">
	nom_2 = campo_estilo_readonly("nom_2", "text", "nom_2", "nobord2", estiloTextBox, "70", "60", "8", nom_aerop2, "readonly", "")
				
	'Selección PICs'
	sql_pic = "Select Num_trip, Nom_trip, Ape_trip, Num_pasa_tri, Fec_venc_pas, Fec_venc_vis From Siep_trip Where "
	'if pilotoViene <> "" then
	   'sql_pic = sql_pic & " (Num_clie = " & Session("Num_clie") & ") And  (Sta_acti = true) And (Num_trip = " & pilotoViene & ")  Order By Nom_trip"
	'else  
	   sql_pic = sql_pic & " (Num_clie = " & Session("Num_clie") & ") And  (Sta_acti = true) Order By Nom_trip"
	'end if
	rstx.Open sql_pic, cnnx, 1, 2
	itemspic1 = ""
	itemspic2 = ""
    Pilotox = ""
    nombrepilotoLst = 0
    apellidopilotoLst = 0
	if not rstx.EOF then			
		v = 1	
		do until rstx.EOF
		'**************************************************************************************
		'* Fecha:13/04/2015
		'* Daniel manda a comentarizar la validación del pasaporte, fecha de vencimiento de pasaporte y visa
		'* para la tripulación debido a que pueden ser pilotos o copilostos venezolanos y hacer un vuelo 
		'* dentro de venezuela y no requiere pasaporte ni visa
		'**************************************************************************************
			'if rstx("Num_pasa_tri") = "" then
			'   	x = 0
			'Else
			'	if rstx("Fec_venc_pas") = "" then
			'		x = 0
					'Aviso de vencimiento no se usa
					'else	
					'	if (DateDiff("d",now(),rstx("Fec_venc_pas")) < 30) then
					'		x = 0
					'	else
					'		x = 1
					'	End if
			'	else
			'			x = 1
			'		End if
			'	End if
			'	'if rstx("Fec_venc_vis") = "" then
				'	v = 1
				'else
				'	if (DateDiff("d",now(),rstx("Fec_venc_vis")) < 30) then
				'		v = 0
				'	else
				'		v = 1
				'	End if
				'End if				
			'	if (x*v = 1) then
			'**************************************************************************************
			'* sigue pero hasta aqui en esta parte
			'**************************************************************************************
  					if Not IsNull(rstx("Nom_trip")) then
						nombrepilotoLst = 1
					else
						nombrepilotoLst = 0
					End if
					if Not IsNull(rstx("Ape_trip")) then
						apellidopilotoLst = 1
					else
						apellidopilotoLst = 0
					End if
    					if(nombrepilotoLst + apellidopilotoLst > 0) then
						Pilotox = rstx("Nom_trip") & " " & rstx("Ape_trip")
						itemspic1 = itemspic1 & rstx("Num_trip") & "|"
						itemspic2 = itemspic2 & Pilotox & "|"
					End if
			'**************************************************************************************
			'* sigue
			'**************************************************************************************
			'	End if
			'**************************************************************************************
			'* hasta aqui los comentarios del fecha 13/04/2015
			'**************************************************************************************
				rstx.MoveNext
			loop 
			if modimanifiesto = true then
				if rstx.RecordCount > 1 then 
					itemspic1 = Left(itemspic1,len(itemspic1)-1)
					itemspic2 = Left(itemspic2,len(itemspic2)-1)
				End if
			else
				itemspic1 = Left(itemspic1,len(itemspic1)-1)
				itemspic2 = Left(itemspic2,len(itemspic2)-1)
			End if
	else
		itemspic1 = "0"
		itemspic2 = "Sin datos.." 
	end if
	rstx.Close
	
	
	
	
	if pilotoViene <> "" then
	   pic_list = seleccion2("pic_list", "pic_list", itemspic1, itemspic2, pilotoViene, "")	   	
	else
       pic_list = seleccion2("pic_list", "pic_list", itemspic1, itemspic2, picListax, "")
    end if	
	
	
	
	'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	'Copilotos'
	nombrecopilotoLst = 0
    apellidocopilotoLst = 0
	sql_copic = "Select Num_trip, Nom_trip, Ape_trip, Num_pasa_tri, Fec_venc_pas, Fec_venc_vis From Siep_trip "
	'if copilotoViene <> "" then
	   'sql_copic = sql_copic & "Where (Num_clie = " & Session("Num_clie") & ") And  (Sta_acti = true) And (Num_trip = " & copilotoViene & ")  Order By Nom_trip"
	'else  
	   sql_copic = sql_copic & "Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) Order By Nom_trip"
	'end if
	rstx.Open sql_copic, cnnx, 1, 2
	itemscopic1 = ""
	itemscopic2 = ""
	Copilotox=""
	if not rstx.EOF then	
		v = 1
		do until rstx.EOF
			'**************************************************************************************
			'* Fecha:13/04/2015
			'* Daniel manda a comentarizar la validación del pasaporte, fecha de vencimiento de pasaporte y visa
			'* para la tripulación debido a que pueden ser pilotos o copilostos venezolanos y hacer un vuelo 
			'* dentro de venezuela y no requiere pasaporte ni visa
			'**************************************************************************************
			'if rstx("Num_pasa_tri") = "" then
			'	x = 0
			'Else
			'	if rstx("Fec_venc_pas") = "" then
			'	   x = 0
			'	else	
			'	   if (DateDiff("d",now(),rstx("Fec_venc_pas")) < 30) then
			'	      x = 0
			'	   else
			'	      x = 1
			'	   End if
			'	else (este incluso esta demás
			'	   x = 1
			'	End if
			'End if
			'if rstx("Fec_venc_vis") = "" then
			'	v = 1
			'else
			'	if (DateDiff("d",now(),rstx("Fec_venc_vis")) < 30) then
			'		v = 0
			'	else
			'		v = 1
			'	End if
			'End if
			'if (x*v = 1) then
			'**************************************************************************************
			'* sigue pero hasta aqui en esta parte
			'**************************************************************************************
				if Not IsNull(rstx("Nom_trip")) then
					nombrecopilotoLst = 1
				else
					nombrecopilotoLst = 0
				End if
				if Not IsNull(rstx("Ape_trip")) then
					apellidocopilotoLst = 1
				else
					apellidocopilotoLst = 0
				End if
				if (nombrecopilotoLst + apellidocopilotoLst > 0) then
					Copilotox = rstx("Nom_trip") & " " & rstx("Ape_trip")
					itemscopic1 = itemscopic1 & rstx("Num_trip") & "|"
					itemscopic2 = itemscopic2 & Copilotox & "|"
				End if
			'**************************************************************************************
			'* sigue
			'**************************************************************************************
			'End if
			'**************************************************************************************
			'* hasta aqui los comentarios del fecha 13/04/2015
			'**************************************************************************************
			rstx.MoveNext
		loop 
		itemscopic1 = "0|" & itemscopic1  
		itemscopic2 = "Seleccione..|" & itemscopic2 
		itemscopic1 = Left(itemscopic1,len(itemscopic1)-1)
		itemscopic2 = Left(itemscopic2,len(itemscopic2)-1)
		'End if'
	else
		itemscopic1 = "0"
		itemscopic2 = "Sin datos.." 
	End if
	rstx.Close
	
	if copilotoViene <> "" then
	   copic_list = seleccion2("copic_list", "copic_list", itemscopic1, itemscopic2, copilotoViene, "OnChange='copic();'") 	   	
	else
       copic_list = seleccion2("copic_list", "copic_list", itemscopic1, itemscopic2, copic_listr, "OnChange='copic();'")
	end if
	
	copic_list = copic_list & hidden("copic_list_rev", copic_list_revr)
	
	num_per = hidden("num_per", almasPersonax)
	
	
	
	if request("num_per_tot") <> "" then
	  
	   
	   num_per_tot = campo("num_per_tot", "text", "num_per_tot", "20", "20", "8", request("num_per_tot"), "OnBlur='esNum(this)'")
	   ast_num_per_tot = tag("span", "ast_num_per_tot", "visibility:hidden", marca_error(), "")
	
	 else  
	 
	   num_per_tot = campo("num_per_tot", "text", "num_per_tot", "20", "20", "8", numerPerTotx, "OnBlur='esNum(this)'")
	   ast_num_per_tot = tag("span", "ast_num_per_tot", "visibility:hidden", marca_error(), "")
	   
	 end if  
		
	num_pas = campo_readonly("num_pas", "text", "num_pas", "3", "2", "9", cantPasajx, "readonly", "")
	num_tri = campo_readonly("num_tri", "text", "num_tri", "3", "2", "10", cantTripx, "readonly", "")
	
	'''''''''''''''''''''''''''''''''''''
	
	'[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-[-

	htmlbus = img("../Imagenes/boton_aceptar.jpg", "70", "19", "", "", "")
	camp_busc = campo("camp_busc", "text", "camp_busc", "20", "100", "14", "", "OnBlur='caractNoPermit(this)'")
	boton_buscar = enlace("", "javascript:void(0)", "text-decoration:none", "", htmlbus, "OnClick='buscar()'")
	
	'Lista de pasajeros con checkbox'
	'--------------------------------------
    list_pas = ""
	indi = 1
'--------------------------------------
	
	if id_pasa_lst = "" then
		sqlPasajero = "Select Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) Order By Nom_pasa"
		Call tablaConCheck(sqlPasajero, "")
	else
		sqlPasajero = "Select Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) And (Num_pasa In " & id_pasa_lst & ") Order By Nom_pasa"
		Call tablaConCheck(sqlPasajero, "bgcolor='#00aff2'")
		sqlPasajero = "Select Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) And (Num_pasa Not In " & id_pasa_lst & ") Order By Nom_pasa"
		Call tablaConCheck(sqlPasajero, "")
	End if
'--------------------------------------


	'End if
	if (resultBusq <> "") then
		if(LstNumPasaMarc <> "") then
			LstNumPasaMarc = Left(LstNumPasaMarc,len(LstNumPasaMarc)-1)
		End if
	End if
	'---------------------------------------
	
	LstNumTripMarc = ""
	indi = 1
	list_tri = ""
	Call tablaConCheckTri

	if (resultBusq <> "") then
		if(LstNumTripMarc <> "") then
			LstNumTripMarc = Left(LstNumTripMarc,len(LstNumTripMarc)-1)
		End if
	End if
	thidden = hidden("busqueda", bsq)
	thidden = thidden & hidden("pasa_mar_bus", LstNumPasaMarc)
	thidden = thidden & hidden("trip_mar_bus", LstNumTripMarc)
	
               
        '------------------------------------------------------TRIPULACION-----------------------------
	
	    rstx.close
		
		bus_trip = "<table id='busqueda_trip' class='display'  >"
        bus_trip = bus_trip& "<thead>"
		bus_trip = bus_trip & "<tr>"		
		bus_trip = bus_trip & "Filtro busqueda: " & campo("busquedaTrip", "text", "bus_pas", "20", "20", "8","", "onkeypress='busPas()'")
		bus_trip = bus_trip &  enlace("", "javascript:void(0)", "text-decoration:none", "", htmlbus, "OnClick='submit()'")
		bus_trip = bus_trip & "</tr>"
		bus_trip = bus_trip & "</thead>"
	
	    tn = "<table id='table_id' class='display'  >"
		tn = tn & "<thead>"
		tn = tn & "<tr>"
		tn = tn & "<th style='display:none'>Num_trip</span></th>"
		tn = tn & "<th><span class='Estilo25'>Nombre</span></th>"
		tn = tn & "<th><span class='Estilo25'>Apellido</span></th>"
		tn = tn & "<th><span class='Estilo25'>Cedula</span></th>"
		tn = tn & "<th><span class='Estilo25'>Numero Pasaporte</span></th>"
		tn = tn & "<th><span class='Estilo25'>Pasaporte</span></th>"
	    tn = tn & "<th><span class='Estilo25'>Visa</span></th>"
		tn = tn & "<th><span class='Estilo25'>Licencia</span></th>"
		tn = tn & "<th><span class='Estilo25'>Certificado EUA</span></th>"
		tn = tn & "<th><span class='Estilo25'>Certificado</span></th>"
		'tn = tn & "<th><span class='Estilo25'>Seleccione</span></th>"
		tn = tn & "</tr>"
		tn = tn & "</thead>"		
		

		
		
		
			if filtrobusquedaTrip <> "" or tripulacion <> ""  then 
			   
			   if filtrobusquedaTrip <> "" then
			     rstx.Open "Select Num_trip, Nom_trip,Ape_trip, Sta_acti, Num_visa_tri, Fec_venc_pas, Fec_venc_vis, Fec_venc_lic, Fec_venc_cer, Fec_exp_eua_cer,Num_pasa_tri,Num_cedu_tri,Num_pasa_tri  From Siep_trip Where Num_clie = " & Session("Num_clie") & " And  (Sta_acti = true) And ((Nom_trip Like '%" & filtrobusquedaTrip & "%') Or (Ape_trip Like '%" & filtrobusquedaTrip & "%')) Order By Nom_trip ", cnnx, 1, 2
			   else
			     tripulacionRep = Replace(tripulacion,"|",",")
			     rstx.Open "Select Num_trip, Nom_trip,Ape_trip, Sta_acti, Num_visa_tri, Fec_venc_pas, Fec_venc_vis, Fec_venc_lic, Fec_venc_cer, Fec_exp_eua_cer,Num_pasa_tri,Num_cedu_tri,Num_pasa_tri  From Siep_trip Where Num_clie = " & Session("Num_clie") & " And  (Sta_acti = true) And Num_trip in (" & tripulacionRep & ")  Order By Nom_trip ", cnnx, 1, 2
			   end if
			
			
			else       
			   rstx.Open "Select top 1 Num_trip, Nom_trip,Ape_trip, Sta_acti, Num_visa_tri, Fec_venc_pas, Fec_venc_vis, Fec_venc_lic, Fec_venc_cer, Fec_exp_eua_cer,Num_pasa_tri,Num_cedu_tri,Num_pasa_tri  From Siep_trip Where Num_trip = 0 ", cnnx, 1, 2
			end if
		
		
		
		
		if not rstx.EOF then
		
		tn = tn & "<tbody>"		
			cantt = 0		
			Do Until rstx.EOF
			    fec_venc_pas_tri = "" 
				fec_venc_vis_tri = ""
				fec_venc_lic_tri = ""
				'fec_venc_lic_eua_tri = ""
				fec_venc_cer_eua_tri = ""
				fec_venc_cer_tri = ""
			
			   
			    if Request("c2") <> "" then
			
					
					c2 = split(Request("c2"), "|")					
					
					for each x in c2
						'id = x
						
						if int(rstx("Num_trip")) = int(x) then 
						   'estilo  = "class='odd selected'"
						   estilo  = "class='selected'"	
						   cantt = cantt + 1
						   optio = optio & "<option value='"&rstx("Num_trip")&"'>"&rstx("Nom_trip")&"</option>"
						   
						   EXIT FOR				   
						  
						else
						   estilo = ""   
						end if					 
					next  
					
					
			     end if
				 
				 
				 '------------------------------------------------------------------------------------------------
				 'para modificacion - HUMBERTO ROJAS
				 
				 if tripulacion <> "" then
				 
				 				'response.write "tripulacion " tripulacion
					
					
				 
				 
				        Arraytripulacion = split(tripulacion, "|")
						
						if ubound(Arraytripulacion) > 0 then
						
						
								For Each Valor In Arraytripulacion 	
								'for k = 0 to UBound(list_tripulacion)
									if Trim(Valor) = Trim(rstx("Num_trip")) then 
										'chequeado = "checked"
										estilo  = "class='selected'"	
										cantt = cantt + 1
										optio = optio & "<option value='"&rstx("Num_trip")&"'>"&rstx("Nom_trip")&" "&rstx("Ape_trip")&"</option>"
										modi = 1
										
										'response.write  optio
										exit for
										 
									else
										estilo = "" 
										
										
									End if
								next
								
						  else		
						  
						        if Trim(tripulacion) = Trim(rstx("Num_trip")) then 
						  
									 estilo  = "class='selected'"	
									 cantt = cantt + 1
									 optio = optio & "<option value='"&rstx("Num_trip")&"'>"&rstx("Nom_trip")&" "&rstx("Ape_trip")&"</option>"
									 modi = 1 
								else 	 
									 
									 estilo = "" 
								 
								 end if 
						  
						  
								
						  end if		
								
					End if
			
			   
			   
               tn = tn & "<tr "&estilo&">"
			   tn = tn & "<td style='display:none'>" & rstx("Num_trip") & "</td>"
               tn = tn & "<td  class='Estilo4' style='color:#000;'>" & rstx("Nom_trip") & "</td>"
			   tn = tn & "<td class='Estilo4' style='color:#000;'>" & rstx("Ape_trip") & "</td>"
			   tn = tn & "<td><span class='Estilo4' style='color:#000;'>" & rstx("Num_cedu_tri") & "</span></td>"
			   tn = tn & "<td><span class='Estilo4' style='color:#000;'>" & rstx("Num_pasa_tri") & "</span></td>"
			   tn = tn & "<td><span class='Estilo4' style='color:#000;'>"             
               
               
			   
			   
			   if Not IsNull(rstx("Fec_venc_pas")) then 
					disabl = ""
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) < 0 then
						fec_venc_pas_tri =  "<span class='Estilo4'>Pasaporte vencido</span>"
					End if
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) <= 30 And DateDiff("d",  Now(), rstx("Fec_venc_pas")) >= 0 then
						fec_venc_pas_tri =  "<span style='color:#FFBF00'>Por vencerse</span>"
						disabl = ""
					End if
					'if  DateDiff("d",  Now(), rstx("Fec_venc_pas")) > 30  then
					'	disabl = ""
					'End if
				else
					fec_venc_pas_tri = "<span class='Estilo4'>Sin fecha de vencimiento</span>"
					'disabl = "disabled"   And (DateDiff('d',now(),Fec_venc_eua_lic)>0) And '
				End if
			   
            
                 
               tn = tn & fec_venc_pas_tri & "</span></td>"
			   tn = tn & "<td><span class='Estilo4' style='color:#000;'>"
		
		
		       if Not IsNull(rstx("Num_visa_tri")) then
					if Not IsNull(rstx("Fec_venc_vis")) then 
						disabl = ""
						if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) < 0 then
							fec_venc_vis_tri = "<span class='Estilo4'>Vencida</span>"
						End if
						if  DateDiff("d", Now(), rstx("Fec_venc_vis")) < 30 And DateDiff("d",Now(), rstx("Fec_venc_vis")) >= 0  then
							fec_venc_vis_tri = " <span style='color:#FFBF00'>Por Vencerse</span>"
						End if
						
						'if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) > 30 then
						'	disabl = ""
						'End if
					else
						disabl = ""
						fec_venc_vis_tri = "<span class='Estilo4'>Sin fecha</span>"
					End if
				else
					fec_venc_vis_tri = " <span class='Estilo4'>Sin visa</span>"
				End if
		
		
		        tn = tn & fec_venc_vis_tri & "</span></td>"
				tn = tn & "<td><span class='Estilo4' style='color:#000;'>"
				
                'tn = tn & "</tr>"   
				
				if Not IsNull(rstx("Fec_venc_lic")) then
					disabl = ""
					if  DateDiff("d",  Now(), rstx("Fec_venc_lic")) < 0 then
						fec_venc_lic_tri = "<span class='Estilo4'>Vencida</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_lic_tri = "<span class='Estilo4'>Sin fecha</span>"
				End if
				
				
				
		        tn = tn & fec_venc_lic_tri & "</span></td>"	
				tn = tn & "<td><span class='Estilo4' style='color:#000;'>"
				
				
				if Not IsNull(rstx("Fec_exp_eua_cer")) then
					disabl = ""
					if  DateDiff("d",  Now(), vencCerEUA(rstx("Fec_exp_eua_cer"))) < 0 then
						fec_venc_cer_eua_tri = "<span class='Estilo4'>Vencido</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_cer_eua_tri = "<span class='Estilo4'>Sin fecha de vencimiento</span>"
				End if
				
				tn = tn & fec_venc_cer_eua_tri & "</span></td>"
				tn = tn & "<td><span class='Estilo4' style='color:#000;'>"
				
				if Not IsNull(rstx("Fec_venc_cer")) then
					disabl = ""
					if  DateDiff("d",  Now(), rstx("Fec_venc_cer")) < 0 then
						fec_venc_cer_tri = "<span class='Estilo4'>Vencido</span>"
					'else
					'	disabl = ""
					End if
				else
					disabl = ""
					fec_venc_cer_tri = "<span class='Estilo4'>Sin fecha</span>"
				End if
				
		
		        tn = tn & fec_venc_cer_tri & "</span></td>"
				
				
				
				
				
				
				
							 
                tn = tn & "</tr>"
				
				
				
		
		
		
		
		rstx.MoveNext
				
			Loop
		        
				tn = tn & "</tbody>	"
			
			
		End if
		
		
		tn = tn & "</table>"
		'response.write tn 			
	
	
	
	'End Sub
	
	
	'------------------------------------------------------PASAJEROS
        
        
      
	 
	
	
	
	
	    rstx.close
		bus_pas = "<table id='busqueda_pas' class='display'  >"
        bus_pas = bus_pas& "<thead>"
		bus_pas = bus_pas & "<tr>"		
		bus_pas = bus_pas & "Filtro busqueda: " & campo("busquedaPas", "text", "bus_pas", "20", "20", "8","", "onkeypress='busPas()'")
		bus_pas = bus_pas &  enlace("", "javascript:void(0)", "text-decoration:none", "", htmlbus, "OnClick='submit()'")
		bus_pas = bus_pas & "</tr>"
		bus_pas = bus_pas & "</thead>"
		
		
	
	    tn_pas = "<table id='table_id_pas' class='display'  >"		
		tn_pas = tn_pas & "<thead>"
		tn_pas = tn_pas & "<tr>"
		tn_pas = tn_pas & "<th style='display:none'>Num_pasa</span></th>"
		tn_pas = tn_pas & "<th><span class='Estilo25'>Nombre</span></th>"
		tn_pas = tn_pas & "<th><span class='Estilo25'>Apellido</span></th>"
		tn_pas = tn_pas & "<th><span class='Estilo25'>Cedula</span></th>"
		tn_pas = tn_pas & "<th><span class='Estilo25'>Número Pasaporte</span></th>"
		tn_pas = tn_pas & "<th><span class='Estilo25'>Pasaporte</span></th>"
	    tn_pas = tn_pas & "<th><span class='Estilo25'>Visa</span></th>"
		'tn_pas = tn_pas & "<th><span class='Estilo25'>Seleccione</span></th>"
		tn_pas = tn_pas & "</tr>"
		tn_pas = tn_pas & "</thead>"
		
	
	
		
			if filtrobusqueda <> "" or pasajero <> ""  then
			   
			   if filtrobusqueda <> "" then			   
			      rstx.Open "Select Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas,Num_cedu_pas, Num_pasa_pas,(Nom_pasa + ' ' + Ape_pasa) as nombrec From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) And ((Nom_pasa Like '%" & filtrobusqueda & "%') Or (Ape_pasa Like '%" & filtrobusqueda & "%'))  Order By Nom_pasa ", cnnx, 1, 2 
			   else 
			      pasajeroRep = Replace(pasajero,"|",",")
			      rstx.Open "Select Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas,Num_cedu_pas, Num_pasa_pas,(Nom_pasa + ' ' + Ape_pasa) as nombrec From Siep_pasa Where (Num_clie = " & Session("Num_clie") & ") And (Sta_acti = true) And Num_pasa in (" & pasajeroRep & ")   Order By Nom_pasa ", cnnx, 1, 2 
			   end if
			
			else
			   rstx.Open "Select top 1 Num_pasa, Nom_pasa, Ape_pasa, Sta_acti, Num_visa_pas, Fec_venc_pas, Fec_venc_vis, Fec_nac_pas,Num_cedu_pas, Num_pasa_pas,(Nom_pasa + ' ' + Ape_pasa) as nombrec From Siep_pasa Where Num_pasa = 0 ", cnnx, 1, 2               
			end if			
			

        
		if not rstx.EOF then
		
		tn_pas = tn_pas & "<tbody>"		
			cantp = 0 		
			Do Until rstx.EOF
			   fec_venc_pas_pasa = ""
               fec_venc_vis_pasa = ""
						
			
			   				 
				 '------------------------------------------------------------------------------------------------
				 'para retornar si ocurre un error - HUMBERTO ROJAS			 
				
			
				 
					 if Request("c1")  <> ""   then
						c1 = split(Request("c1"), "|")
						for each x2 in c1
							if int(rstx("Num_pasa")) = int(x2) then 
							   estilo2  = "class='selected'"	
							   cantp = cantp + 1
							   optio2 = optio2 & "<option value='"&rstx("Num_pasa")&"'>"&rstx("Nom_pasa")&"</option>"
							   EXIT FOR		
							else
							   estilo2 = ""   
							end if					 
						next  
					 end if
					 
				
				
				 
				 '------------------------------------------------------------------------------------------------
				 'para modificacion - HUMBERTO ROJAS
			  
				 
				 if pasajero <> "" then
				   
				       
						 Arraypasajero = split(pasajero, "|")
						 
						 if ubound(Arraypasajero) > 0 then
						 
								For Each Valor2 In Arraypasajero 	
								
								
									if Trim(Valor2) = Trim(rstx("Num_pasa")) then 
										'chequeado = "checked"
										estilo2  = "class='selected'"	
										cantp = cantp + 1
										optio2 = optio2 & "<option value='"&rstx("Num_pasa")&"'>"&rstx("Nom_pasa")&" "&rstx("Ape_pasa")&"</option>"
										modi = 1
										exit for
										'response.write  optio2
										 
									else
										estilo2 = "" 
										
										
									End if
								next
								
							else									
								
								if Trim(pasajero) = Trim(rstx("Num_pasa")) then 
										'chequeado = "checked"
										estilo2  = "class='selected'"	
										cantp = cantp + 1
										optio2 = optio2 & "<option value='"&rstx("Num_pasa")&"'>"&rstx("Nom_pasa")&" "&rstx("Ape_pasa")&"</option>"
										modi = 1										 
								 else
										estilo2 = "" 
										
										
								 End if 
								
							end if	
								
					End if
			
			
			
			   num_pas_bd = rstx.RecordCount
			
			   
               tn_pas = tn_pas & "<tr "&estilo2&">"
			   tn_pas = tn_pas & "<td style='display:none'>" & rstx("Num_pasa") & "</td>"
               tn_pas = tn_pas & "<td  class='Estilo4' style='color:#000;'>" & rstx("Nom_pasa") & "</td>"
			   tn_pas = tn_pas & "<td class='Estilo4' style='color:#000;'>" & rstx("Ape_pasa") & "</td>"
			   tn_pas = tn_pas & "<td><span class='Estilo4' style='color:#000;'>" & rstx("Num_cedu_pas") & "</span></td>"
			   tn_pas = tn_pas & "<td><span class='Estilo4' style='color:#000;'>" & rstx("Num_pasa_pas") & "</span></td>"
			   
			   tn_pas = tn_pas & "<td><span class='Estilo4' style='color:#000;'>"             
               
               
			   
			   
			   if Not IsNull(rstx("Fec_venc_pas")) then 
					disabl = ""
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) < 0 then
						fec_venc_pas_pasa =  "<span class='Estilo4'>Pasaporte vencido</span>"
					End if
					if  DateDiff("d", Now(), rstx("Fec_venc_pas")) <= 30 And DateDiff("d",  Now(), rstx("Fec_venc_pas")) >= 0 then
						fec_venc_pas_pasa =  "<span style='color:#FFBF00'>Por vencerse</span>"
						disabl = ""
					End if
					'if  DateDiff("d",  Now(), rstx("Fec_venc_pas")) > 30  then
					'	disabl = ""
					'End if
				else
					fec_venc_pas_pasa = "<span class='Estilo4'>Sin fecha de vencimiento</span>"
					'disabl = "disabled"   And (DateDiff('d',now(),Fec_venc_eua_lic)>0) And '
				End if
			   
            
                 
               tn_pas = tn_pas & fec_venc_pas_pasa & "</span></td>"
			   tn_pas = tn_pas & "<td><span class='Estilo4' style='color:#000;'>"
		
		
		       if Not IsNull(rstx("Num_visa_pas")) then
					if Not IsNull(rstx("Fec_venc_vis")) then 
						disabl = ""
						if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) < 0 then
							fec_venc_vis_pasa = "<span class='Estilo4'>Vencida</span>"
						End if
						if  DateDiff("d", Now(), rstx("Fec_venc_vis")) < 30 And DateDiff("d",Now(), rstx("Fec_venc_vis")) >= 0  then
							fec_venc_vis_pasa = " <span style='color:#FFBF00'>Por Vencerse</span>"
						End if
						
						'if  DateDiff("d",  Now(), rstx("Fec_venc_vis")) > 30 then
						'	disabl = ""
						'End if
					else
						disabl = ""
						fec_venc_vis_pasa = "<span class='Estilo4'>Sin fecha</span>"
					End if
				else
					fec_venc_vis_pasa = " <span class='Estilo4'>Sin visa</span>"
				End if
		
		
		        tn_pas = tn_pas & fec_venc_vis_pasa & "</span></td>"
				 'tn_pas = tn_pas &  "<td></td>"
				 
				'tn_pas = tn_pas & "<td style='display:none'>" & rstx("nombrec") & "</td>"
                tn_pas = tn_pas & "</tr>"   
		
		
		
		
		rstx.MoveNext
				
			Loop
		        
				tn_pas = tn_pas & "</tbody>	"
			
			
		End if
		
		
		tn_pas = tn_pas & "</table>"
                
	
	
	
	        selec =  "<div style='margin-top:10px; margin-bottom:10px; width:100%;float:left;margin-left: -180px;' >"
			selec = selec & "<div style='text-align:right;margin-right:30%' ><span class='Estilo32Copia'>Piloto: </span>" & pic_list  & "</div>"
			selec = selec & "<div  style='text-align:right;margin-right:30%'><span class='Estilo32Copia'>Copiloto: </span>" & copic_list  & "</div>"
			selec = selec & "</div >"
            selec = selec &  "<div id='' class='display' >"
			selec = selec & "<div style = 'float: left'>"
			selec = selec & "<div><span class='Estilo25'>Tripulación adicional seleccionada</span></div>"
			selec = selec & "<select name='selec' size='10' id='selec' multiple style='width:300px;'>"
			if Request("c2") <> ""  then
			   selec = selec & optio
			end if
			if tripulacion <> "" then
			      selec = selec & optio
			end if
			selec = selec & " </select>"
			selec = selec & "<div ><span class='Estilo25'>Cant. Trip.: </span> <input name='canttrip' type='text' id='canttrip' size='5' maxlength='2' value = "&cantt&"> </div>"			
			selec = selec & "</div>"
			selec = selec & "<input type='hidden' id = 'selectripinput' name='selectripinput' >"
			selec = selec & "<div style = 'float: left'>"			
			selec = selec & "<div><span class='Estilo25'>Pasajeros Seleccionados</span></div>"
			selec = selec & "<select name='selecpas' size='10' id='selecpas' multiple disabled style='width:300px;'>"			
			if Request("c1") <> ""  then
			   selec = selec & optio2
			end if
			if pasajero <> "" then
			      selec = selec & optio2
			end if
			selec = selec & " </select>"
			selec = selec & "<div ><span class='Estilo25'>Cant. Pas.: </span> <input name='cantpas' type='text' id='cantpas' size='5' maxlength='2' value = "&cantp&"> </div>"
			selec = selec & "</div>"
			selec = selec & "<input type='hidden' id = 'selecpasinput' name='selecpasinput' >"
			'selec = selec & "<input type='hidden' id = 'vertrip' name='vertrip' > "
           ' selec = selec & "<input type='hidden' id = 'verpas' name='verpas'  > " 
                       
            selec = selec & "</div>"
			
      
	'HUMBERTO UPDATE 2018   
	
	selecpasinputPrin = ""
	selectripinputPrin = ""
	
	if request.form("selecpasinput") <> "" then
	   selecpasinputPrin = request.form("selecpasinput")
    elseif request("selecpasinput") <> "" then
       selecpasinputPrin = request("selecpasinput")	
	end if 

    if request.form("selectripinput") <> "" then
	   selectripinputPrin = request.form("selectripinput")
    elseif request("selectripinput") <> "" then
       selectripinputPrin = request("selectripinput")	
	end if 
    
    if request.form("canttrip") <> "" then
	   canttripPrin = request.form("canttrip")
    elseif request("canttrip") <> "" then
       canttripPrin = request("canttrip")	
	end if  
    
	if request.form("cantpas") <> "" then
	   cantpasPrin = request.form("cantpas")
    elseif request("cantpas") <> "" then
       cantpasPrin = request("cantpas")	
	end if  
	
	
	
	if selecpasinputPrin <> "" OR selectripinputPrin <> "" then
	        if canttripPrin <> "" then  
	          cantt = canttripPrin
            else
			  cantt = 0
			end if 
            if cantpasPrin <> "" then  
	          cantp = cantpasPrin
            else
			  cantp = 0
			end if 	
	        selec = ""
	        selec =  "<div style='margin-top:10px; margin-bottom:10px; width:100%;float:left;margin-left: -180px;' >"
			selec = selec & "<div style='text-align:right;margin-right:30%' ><span class='Estilo32Copia'>Piloto: </span>" & pic_list  & "</div>"
			selec = selec & "<div  style='text-align:right;margin-right:30%'><span class='Estilo32Copia'>Copiloto: </span>" & copic_list  & "</div>"
			selec = selec & "</div >"
            selec = selec &  "<div id='' class='display' >"
			selec = selec & "<div style = 'float: left'>"
			selec = selec & "<div><span class='Estilo25'>Tripulación adicional seleccionada</span></div>"
			selec = selec & "<select name='selec' size='10' id='selec' multiple style='width:300px;'>"
			if selectripinputPrin <> "" then
                          

			   selectripinput = mid(selectripinputPrin,1, (len(selectripinputPrin)-1))
				rstx.close
				rstx.Open "Select Num_trip, Nom_trip, Ape_trip From Siep_trip Where  Num_trip in (" + selectripinput + ") Order By Nom_trip ", cnnx, 1, 2
				if not rstx.EOF then
				   While not rstx.EOF			       
					   selec = selec & "<option value=" & rstx("Num_trip") & ">" & rstx("Nom_trip") & " " & rstx("Ape_trip") & "</option>"
					   rstx.movenext
				  Wend		           
				end if
			end if
			selec = selec & " </select>"
			selec = selec & "<div ><span class='Estilo25'>Cant. Trip.: </span> <input name='canttrip' type='text' id='canttrip' size='5' maxlength='2' value = "&cantt&"> </div>"			
			selec = selec & "</div>"
			selec = selec & "<input type='hidden' id = 'selectripinput' name='selectripinput' >"
            'selec = selec & "<input type='hidden' id = 'vertrip' name='vertrip' value=1 > "
            'selec = selec & "<input type='hidden' id = 'verpas' name='verpas' value=0 > " 
			selec = selec & "<div style = 'float: left'>"			
			selec = selec & "<div><span class='Estilo25'>Pasajeros Seleccionados</span></div>"
			selec = selec & "<select name='selecpas' size='10' id='selecpas' multiple disabled style='width:300px;'>"
                        if selecpasinputPrin <> "" then
                           'vertrip = 0
                           'verpas = 1	
'response.write verpas &" "& vertrip		
			   selecpasinput = mid(selecpasinputPrin,1, (len(selecpasinputPrin)-1))
				rstx.close
				rstx.Open "Select Num_pasa, Nom_pasa, Ape_pasa From Siep_pasa Where  Num_pasa in (" + selecpasinput + ") Order By Nom_pasa ", cnnx, 1, 2
				if not rstx.EOF then
				   While not rstx.EOF			       
					   selec = selec & "<option value=" & rstx("Num_pasa") & ">" & rstx("Nom_pasa") & " " & rstx("Ape_pasa") & "</option>"
					   rstx.movenext
				  Wend		           
		        end if
			end if	
			selec = selec & " </select>"
			selec = selec & "<div ><span class='Estilo25'>Cant. Pas.: </span> <input name='cantpas' type='text' id='cantpas' size='5' maxlength='2' value = "&cantp&"> </div>"
			selec = selec & "</div>"
			selec = selec & "<input type='hidden' id = 'selecpasinput' name='selecpasinput' >"
            'selec = selec & "<input type='hidden' id = 'vertrip' name='vertrip' value=0 > "
            'selec = selec & "<input type='hidden' id = 'verpas' name='verpas' value=1 > " 
            selec = selec & "</div>"
	end if
%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
