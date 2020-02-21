'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~		  SCRIPT MADE BY JORDI PARÉ			~'
'~		 	|ACCESS RANGE SCRIPT|			~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~			  RICOH SANT CUGAT				~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~		TWITTER: @_xJPBx_					~'
'~		INSTAGRAM: jpareb					~'
'~		EMAIL: jparebernado@gmail.com		~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

set windows=CreateObject("WScript.shell")
set windows_pestana = CreateObject("Shell.Application")
set objeto_leer_doc = CreateObject("Scripting.FileSystemObject")
dim numero_de_paquetes_ping

s_raiz = "C:\Users\Jordi.pare\Desktop\temp\"
s_impresoras_incorrectas = s_raiz & "impresoras_incorrectas.csv"
s_impresoras_correctas = s_raiz & "impresoras_correctas.csv"
s_rango_impresoras = s_raiz & "rango_impresoras.csv"
s_csv_imp = s_raiz & "csv_imp.csv"
s_dns = s_raiz & "dns.csv"
s_telnet = s_raiz & "telnet.csv"


select case(objeto_leer_doc.FileExists(s_impresoras_incorrectas))
	
	Case 0
		set impresoras_incorrectas = objeto_leer_doc.OpenTextFile(s_impresoras_incorrectas,8, true)
		impresoras_incorrectas.WriteLine("IP,HOSTNAME,ACCESSRANGE 1.1-ACCESSRANGE 1.2,ACCESSRANGE 2.1-ACCESSRANGE 2.2,ACCESSRANGE 3.1-ACCESSRANGE 3.2,ACCESSRANGE 4.1-ACCESSRANGE 4.2,ACCESSRANGE 5.1-ACCESSRANGE 5.2,FECHA,HORA,PROBLEMA")
	Case -1
		set impresoras_incorrectas = objeto_leer_doc.OpenTextFile(s_impresoras_incorrectas,8, true)
end select

select case(objeto_leer_doc.FileExists(s_impresoras_correctas))
	
	Case 0
		set impresoras_correctas = objeto_leer_doc.OpenTextFile(s_impresoras_correctas,8, true)
		impresoras_correctas.WriteLine("IP,HOSTNAME,ACCESSRANGE 1.1-ACCESSRANGE 1.2,ACCESSRANGE 2.1-ACCESSRANGE 2.2,ACCESSRANGE 3.1-ACCESSRANGE 3.2,ACCESSRANGE 4.1-ACCESSRANGE 4.2,ACCESSRANGE 5.1-ACCESSRANGE 5.2,FECHA,HORA, ESTADO")
	Case -1
		set impresoras_correctas = objeto_leer_doc.OpenTextFile(s_impresoras_correctas,8, true)
end select

select case (objeto_leer_doc.FileExists(s_csv_imp))
	
	Case 0
		set csv_ip_o = objeto_leer_doc.OpenTextFile(s_csv_imp, 8, true)
		csv_ip_o.WriteLine("IP/HOSTNAME,ACCESSRANGE 1.1-ACCESSRANGE 1.2,ACCESSRANGE 2.1-ACCESSRANGE 2.2,ACCESSRANGE 3.1-ACCESSRANGE 3.2,ACCESSRANGE 4.1-ACCESSRANGE 4.2,ACCESSRANGE 5.1-ACCESSRANGE 5.2")
	Case -1
		set csv_ip_o = objeto_leer_doc.OpenTextFile(s_csv_imp, 8, true)
end select





set rango_impresoras_a = objeto_leer_doc.OpenTextFile(s_rango_impresoras,8, true)

telnet_exito = 0
funcion_servicio_cmd(1)
numero_de_paquetes_ping= 3
windows_pestana.MinimizeAll

function funcion_servicio_cmd(matar_iniciar)
	select case (matar_iniciar)
		Case "1"
			windows.run "taskkill /im cmd.exe", , True
		Case "0"
			windows.Run "%comspec%"
			WScript.sleep 2000
	end select
end function


function preguntar (caso_pregunta)

	select case (caso_pregunta)
		Case 0
			IPIM = InputBox("IP MAQUINA")
			IPIM_splited = Split(IPIM, ".")
			
			
			do while ip_correctas(IPIM_splited) <> 1
			
				IPIM = InputBox("IP MAQUINA")
				If IPIM = "" Then Exit function
				IPIM_splited = Split(IPIM, ".")
				
			loop
			
			IPU = InputBox("PRIMERA IP DEL RANGO")
			IPU_splited = Split(IPU, ".")
			
			do while ip_correctas(IPU_splited) <> 1
				IPU = InputBox("PRIMERA IP DEL RANGO")
				If IPU = "" Then Exit function
				IPU_splited = Split(IPU, ".")
			loop
			
			
			IPD = InputBox("SEGUNDA IP DEL RANGO")
			IPD_splited = Split(IPD, ".")
			
			do while ip_correctas(IPD_splited) <> 1
				IPD = InputBox("SEGUNDA IP DEL RANGO")
				If IPD = "" Then Exit function
				IPD_splited = Split(IPD, ".")
			loop
			
			access_range = InputBox("RANGE 1-5")
			if access_range=0 then
				WScript.Echo "LA OPCION NO ESTA ENTRE 1 Y 5"
			end if
			do while access_range >5 or access_range<1
			access_range = InputBox("RANGE 1-5")
			loop
			
			contrasena = InputBox("PASS DEL USER ADMIN")
			
			preguntar = IPIM &","& IPU &","& IPD &","& contrasena &","& access_range
			
		Case 3
			archivo = InputBox("¿DONDE ESTA EL ARCHIVO? ej. C:\users\jordi\desktop\printers.csv")
			select case (objeto_leer_doc.FileExists(archivo))
				Case -1
					
				Case ""
					MsgBox ("NO SE ENCUENTRA EL FICHERO")
					exit function
			end select
					
			contrasena = InputBox("PASS DEL USER ADMIN")
			
			preguntar = contrasena &","& archivo
	end select

end function

function ip_correctas (IP)
	IP_suma = uBound(IP) + 1
	
	for i = 0 to IP_suma - 1
		select case (IP(i))
			Case ""
				'MsgBox ("LA IP ES INCORRECTA")
				ip_correctas = 2
				exit function
		end select 
	next
	
	if IP_suma = 4 then
		ip_correctas = 1
	end if
	
	if IP_suma <> 4 then
		ip_correctas = 2		
	end if
end function

function datos_impresora_uno ()
			
	res = preguntar(0)
	arr_res = Split(res,",")
	
	IPIM = arr_res(0)
	IPU = arr_res (1)
	IPD = arr_res (2)
	contrasena = arr_res (3)
	access_range = arr_res (4)
	
	config contrasena, IPU, IPD, IPIM, access_range
	
end function

function nslookup (IPIM)
	funcion_servicio_cmd(0)
	windows.sendKeys "nslookup " & IPIM & " >> "&s_dns&" "
	
	windows.sendkeys ("{Enter}")
	WScript.sleep 2000
	set dns = objeto_leer_doc.OpenTextFile(s_dns)
	do until dns.AtEndOfStream
		linia_dns = dns.ReadLine
		if (InStr(linia_dns,"Name:")) > 0 then
			nslookup = Right(linia_dns, 38)
			exit do
		else
			nslookup = "no DNS"
		end if
	loop
	WScript.sleep 50
	dns.close
	WScript.sleep 50
	objeto_leer_doc.DeleteFile(s_dns)
	funcion_servicio_cmd(1)
	
end function

function config (contrasena, IPU, IPD, IPIM, access_range)
	posible_conexion = Not CBool(windows.run("ping -n " & numero_de_paquetes_ping & " " & IPIM,0,True))
	select case (posible_conexion)
		Case -1
			correcto_o_no = funcion_cambiar_config_maquina(contrasena, IPU, IPD, IPIM, access_range)
			select case (correcto_o_no)
				Case 1
					impresoras_correctasf IPIM, IPU, IPD, access_range
				Case 0
					tipo_error = "telnet"
					impresoras_incorrectasf IPIM, IPU, IPD, tipo_error, access_range
					
			end select
		Case 0
			tipo_error = "Ping"
			impresoras_incorrectasf IPIM, IPU, IPD, tipo_error, access_range
	end select
	
end function

function impresoras_incorrectasf (IPIM, IPU, IPD, tipo_error, access_range)
	
	a_IPU = Split(IPU,",")
	a_IPD = Split(IPD,",")
	rango = ""
	a_access_range = Split(access_range,",")
	rango_no_existe = "1,2,3,4,5"
	a_rango_no_existe = Split(rango_no_existe,",")
	noarray = ""
	Dim rango_esta (4)

	for y = 0 to uBound(rango_esta)
		
		rango_esta(y) = 0
		
	next
	
	for i = 0 to uBound(a_access_range)
		select case (a_access_range(i))
			case "1"
				
				rango_esta(0) = 1
			case "2"
				
				rango_esta(1) = 1
			case "3"
				
				rango_esta(2) = 1
			case "4"
				
				rango_esta(3) = 1
			case "5"
				
				rango_esta(4) = 1
				
		end select
		
	next
	pos_u = 0
	for z = 0 to uBound(rango_esta)
		
		if rango_esta(z) = 1 then
			rango = rango & a_IPD(pos_u) & " - " & a_IPD(pos_u) & ", "
			pos_u = pos_u +1
		else
			rango = rango & ", "
		end if
		
	next
	
	
	impresoras_incorrectas.WriteLine IPIM & ", " & trim(nslookup (IPIM)) & ", " & rango & Date & ", " & Time & ", " & tipo_error
	

end function

function impresoras_correctasf (IPIM, IPU, IPD, access_range)
	a_IPU = Split(IPU,",")
	a_IPD = Split(IPD,",")
	rango = ""
	a_access_range = Split(access_range,",")
	rango_no_existe = "1,2,3,4,5"
	a_rango_no_existe = Split(rango_no_existe,",")
	noarray = ""
	Dim rango_esta (4)

	for y = 0 to uBound(rango_esta)
		
		rango_esta(y) = 0
		
	next
	
	for i = 0 to uBound(a_access_range)
		select case (a_access_range(i))
			case "1"
				
				rango_esta(0) = 1
			case "2"
				
				rango_esta(1) = 1
			case "3"
				
				rango_esta(2) = 1
			case "4"
				
				rango_esta(3) = 1
			case "5"
				
				rango_esta(4) = 1
				
		end select
		
	next
	pos_u = 0
	for z = 0 to uBound(rango_esta)
		
		if rango_esta(z) = 1 then
			rango = rango & a_IPD(pos_u) & " - " & a_IPD(pos_u) & ", "
			pos_u = pos_u +1
		else
			rango = rango & ", "
		end if
		
	next
	
	impresoras_correctas.WriteLine IPIM & ", " & trim(nslookup (IPIM)) & ", " & rango & Date & ", " & Time & ", " & "Correcto"
end function

function funcion_cambiar_config_maquina (contrasena, IPU, IPD, IPIM, access_range)
		
		funcion_servicio_cmd(0)
		windows.sendKeys "telnet -f " & ""&s_telnet&" " & IPIM
		WScript.sleep 4000
		windows.sendkeys ("{Enter}")
		WScript.sleep 2000
		windows.sendkeys "admin"
		windows.sendkeys ("{Enter}")
		WScript.sleep 2000
		windows.sendkeys (contrasena)
		windows.sendkeys ("{Enter}")
		WScript.sleep 2000
		
		select case (InStr(",",IPU))
			
				
			Case 0
				a_IPU = Split(IPU,",")
				a_IPD = Split(IPD,",")
				a_access_range = Split(access_range,",")
				for i = 0 to uBound(a_access_range)
					windows.SendKeys "access "& a_access_range(i) &" range "& a_IPU(i) &" "& a_IPD(i) &""
					WScript.sleep 2000
					windows.SendKeys ("{Enter}")
					WScript.Sleep 2000
				next
			Case Else
				windows.SendKeys "access "& a_access_range(i) &" range "& a_IPU(i) &" "& a_IPD(i) &""
				WScript.sleep 2000
		end select
		
		
		windows.SendKeys ("{Enter}")
		WScript.Sleep 2000
		windows.SendKeys "logout" 
		windows.SendKeys ("{Enter}")
		WScript.Sleep 2000
		windows.SendKeys "yes" 
		windows.SendKeys ("{Enter}")
		WScript.Sleep 2000
		windows.SendKeys ("{Enter}")
			
		set log_de_telnet = objeto_leer_doc.OpenTextFile(s_telnet)
			
		do until log_de_telnet.AtEndOfStream
			linia_log_telnet = log_de_telnet.ReadLine
			if (StrComp(linia_log_telnet,"Now, Save data.")) = 0 then
				telnet_exito = 1
			end if		
		loop
		
		
		log_de_telnet.close
		objeto_leer_doc.DeleteFile(s_telnet)
		funcion_cambiar_config_maquina = telnet_exito
		funcion_servicio_cmd(1)
		telnet_exito = 0
end function


function csv_impresoras()

	info_no_split = preguntar(3)
	tota_la_info = Split(info_no_split,",")
	contrasena = tota_la_info(0)
	ruta_arch_man = tota_la_info(1)
	
	array_con_las_ip = leer_csvf(s_csv_imp)
	array_con_las_ip_s = Split(array_con_las_ip,";")
	length_array = uBound(array_con_las_ip_s)
	
	
	
	for i = 2 to length_array
		contrasena = ""
		IPU = ""
		IPD = ""
		access_range = ""
		IPUi = ""
		IPDi = ""
		access_rangei = ""
		be_mal = 0
		
		dentro_linia = Split(array_con_las_ip_s(i),",")
		
		length_linia = uBound(dentro_linia)
		pos = 0
		for y = 1 to length_linia
			ip_ok = 1
			ip_falla = 99
			select case (dentro_linia(y))
				Case ""
				Case Else
					IP_no_split = Split(dentro_linia(y),"-")
					for k = 0 to uBound(IP_no_split)
						if ((ip_correctas(Split(IP_no_split(k),".")) = 1) and (ip_ok <> 0)) then
							ip_ok = 1
						else
							ip_ok = 0
							ip_falla = y
							pos = y
						end if
					next
					
					
					
					select case (ip_ok)
						Case 0
							be_mal = 1
							if 0 = y then
								select case (IPUi)
									Case ""
										IPUi = IP_no_split(0)
									Case Else
										IPUi = IPUi & "," & IP_no_split(0)
								end select
								
							end if
							
							if 1 = y then
								select case (IPDi)
									Case ""
										IPDi = IP_no_split(1)
									Case Else
										IPDi = IPDi &","& IP_no_split(1)
								end select
								
							end if
							
							select case(access_rangei)
								Case ""
									access_rangei = y
								Case Else
									access_rangei = access_rangei & "," & y
							end select
							
						Case 1
							be_mal = 0
							if ((0 <> y and 1 <> y) or ip_falla = 99) then
								select case (IPU)
									Case ""
										IPU = IP_no_split(0)
									Case Else
										IPU = IPU & "," & IP_no_split(0)
								end select
															
							end if
							
							if ((1 <> y and 0 <> y) or ip_falla = 99) then
								select case (IPD)
									Case ""
										IPD = IP_no_split(1)
									Case Else
										IPD = IPD &","& IP_no_split(1)
								end select
								
							end if
							
							if ((1 <> y and 0 <> y) or ip_falla = 99) then
								select case (access_range)
									Case ""
										access_range = y
									case Else
										access_range = access_range & "," & y
								end select
							end if
							
					end select
				end select
				ip_falla = 0

		next
		
		select case (be_mal)
			Case 0
				
				'impresoras_incorrectasf dentro_linia(0), IPUi, IPDi, "RANGO MAL ESCRITO", access_rangei
				WScript.sleep 2000
				config contrasena, IPU, IPD, dentro_linia(0), access_range
			Case 1
				WScript.sleep 2000
				config contrasena, IPU, IPD, dentro_linia(0), access_range
		end select
	next
		
end function



function leer_csvf(archivo)
	Set direcciones_ip = objeto_leer_doc.OpenTextFile(archivo)
	Dim array1()
	 
	pos = 0
	
	Do Until  direcciones_ip.AtEndOfStream
		ip_csv = direcciones_ip.ReadLine
		select case (ip_csv)
			Case ""
			Case Else
				String_ips = String_ips &";"& ip_csv
				pos = pos +1
		end select
		
	loop
	leer_csvf = String_ips
	direcciones_ip.close
end function

function rango_impresoras()

	WScript.Echo "EN PROCESO DE CREACION"
	exit function

end function

function menu()
	opcion = InputBox("1- UNA MAQUINA" & vbCrLf & "2-RANGO DE MAQUINAS" & vbCrLf & "3-ARCHIVO CSV")
	select case (opcion)
		Case 1
			datos_impresora_uno()
			impresoras_incorrectas.close
		impresoras_correctas.close
		Case 2
			rango_impresoras()		
		case 3
			csv_impresoras()
			
			impresoras_incorrectas.close
	impresoras_correctas.close
	end select
	MsgBox ("UN CORDIAL ADIOS :)")
	
end function

menu()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~		  SCRIPT MADE BY JORDI PARÉ			~'
'~		 	|ACCESS RANGE SCRIPT|			~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~			  RICOH SANT CUGAT				~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~		TWITTER: @_xJPBx_					~'
'~		INSTAGRAM: jpareb					~'
'~		EMAIL: jparebernado@gmail.com		~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
