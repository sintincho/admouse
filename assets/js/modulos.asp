<%
Function CharFix(sIn) 
        Dim oIn
        Set oIn = CreateObject("ADODB.Stream") 
        oIn.Open 
        oIn.CharSet = "WIndows-1252" 
        oIn.WriteText sIn 
        oIn.Position = 0 
        oIn.CharSet = "UTF-8" 
        CharFix = oIn.ReadText 
        oIn.Close 
End Function
'======================================================================================================================
'======================================================================================================================
Function fregexp(tex, bus, trae)
	StringToSearch = tex
	Set RegularExpressionObject = New RegExp
		With RegularExpressionObject
			.Pattern = bus
			.IgnoreCase = True
			.Global = True
		End With

		Set expressionmatch = RegularExpressionObject.Execute(StringToSearch)
			If expressionmatch.Count > 0 Then
				For Each expressionmatched in expressionmatch
					wVarCad = expressionmatched.Value & "," & wVarCad
					cont = cont + 1
				Next
			end if
		Set expressionmatch = nothing
	Set RegularExpressionObject = nothing

	if trae = 1 then
		fregexp = cont
	elseif trae = 2 then
		fregexp = wVarCad
	end if
End Function
'======================================================================================================================
'======================================================================================================================
wVarNombre = request.form("wVarNombre")
wVarApellido = request.form ("wVarApellido")
wVarPais = request.form ("wVarPais")
wVarLocalidad = request.form ("wVarLocalidad")
wVarStatus = request.form("wVarStatus")
wVarEmail = request.form("wVarEmail")
wVarTelefono = request.form ("wVarTelefono")
wVarComentario = request.form ("wVarComentario")

if wVarNombre = "" then
		response.Write("<script language='JavaScript'>")
		response.Write("document.getElementById('NombreError').innerHTML = 'Debe ingresar un Nombre';")
		response.Write("document.getElementById('NombreError').style.display = '';")
		response.Write("document.getElementsByName('wVarNombre')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
if(fregexp(wVarNombre, "^([a-z A-ZÑñÁÉÍÓÚáéíóú]{2,60})$" , 1)) = false then
		response.Write("<script type='text/javascript'>")
		response.Write("document.getElementById('NombreError').innerHTML = 'Debe ingresar un Nombre valido';")
		response.Write("document.getElementById('NombreError').style.display = '';")
		response.Write("document.getElementsByName('wVarNombre')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('ApellidoError').innerHTML = '';")
response.Write("document.getElementById('ApellidoError').style.display = 'none';")
response.Write("</script>")

if wVarApellido = "" then
		response.Write("<script language='JavaScript'>")
		response.Write("document.getElementById('ApellidoError').innerHTML = 'Debe ingresar un Nombre';")
		response.Write("document.getElementById('ApellidoError').style.display = '';")
		response.Write("document.getElementsByName('wVarApellido')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
if(fregexp(wVarApellido, "^([a-z A-ZÑñÁÉÍÓÚáéíóú]{2,60})$" , 1)) = false then
		response.Write("<script type='text/javascript'>")
		response.Write("document.getElementById('ApellidoError').innerHTML = 'Debe ingresar un Nombre valido';")
		response.Write("document.getElementById('ApellidoError').style.display = '';")
		response.Write("document.getElementsByName('wVarApellido')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('ApellidoError').innerHTML = '';")
response.Write("document.getElementById('ApellidoError').style.display = 'none';")
response.Write("</script>")

if wVarPais = "-1" then
    response.Write("<script language='JavaScript'>")
	response.Write("document.getElementById('PaisError').innerHTML = '<font color=#8C0108><b>Debe seleccionar un Pais</b></font>';")
	response.Write("document.getElementsByName('wVarPais')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('PaisError').style.display = '';")
	response.Write("document.getElementsByName('wVarPais')[0].focus();")
	response.Write("grecaptcha.reset();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.end
end if
if(fregexp(wVarPais, "^([a-z A-ZÑñÁÄÉËÍÏÓÖÚÜáäéëíïóöúü]{2,60})$" , 1)) = false then
	response.Write("<script type='text/javascript'>")
	response.Write("document.getElementById('PaisError').innerHTML = '<font color=#8C0108><b>Debe seleccionar un Pais valido</b></font>';")
	response.Write("document.getElementsByName('wVarPais')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('PaisError').style.display = '';")
	response.Write("document.getElementsByName('wVarPais')[0].focus();")
	response.Write("grecaptcha.reset();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.End()
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('PaisError').innerHTML = '';")
response.Write("document.getElementsByName('wVarPais')[0].style.backgroundColor='#FFFFFF';")
response.Write("document.getElementById('PaisError').style.display = 'none';")
response.Write("</script>")

if wVarLocalidad = "" then
		response.Write("<script language='JavaScript'>")
		response.Write("document.getElementById('LocalidadError').innerHTML = 'Debe ingresar una Localidad';")
		response.Write("document.getElementById('LocalidadError').style.display = '';")
		response.Write("document.getElementsByName('wVarLocalidad')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
if(fregexp(wVarLocalidad, "^([a-z A-ZÑñÁÉÍÓÚáéíóú]{2,60})$" , 1)) = false then
		response.Write("<script type='text/javascript'>")
		response.Write("document.getElementById('LocalidadError').innerHTML = 'Debe ingresar una Localidad valida';")
		response.Write("document.getElementById('LocalidadError').style.display = '';")
		response.Write("document.getElementsByName('wVarLocalidad')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('LocalidadError').innerHTML = '';")
response.Write("document.getElementById('LocalidadError').style.display = 'none';")
response.Write("</script>")

if wVarStatus = "-1" then
	response.Write("<script language='JavaScript'>")
	response.Write("document.getElementById('StatusError').innerHTML = '<font color=#8C0108><b>Debe seleccionar un Status</b></font>';")
	response.Write("document.getElementsByName('wVarStatus')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('StatusError').style.display = '';")
	response.Write("document.getElementsByName('wVarStatus')[0].focus();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.end
end if
if(fregexp(wVarStatus, "^(Particular)$|^(Distribuidor)$|^(Docente)$|^(Rehabilitacion)$|^(ONG)$|^(Institucion)$" , 1)) = false then
	response.Write("<script type='text/javascript'>")
	response.Write("document.getElementById('StatusError').innerHTML = '<font color=#8C0108><b>Debe seleccionar una Status valido</b></font>';")
	response.Write("document.getElementsByName('wVarStatus')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('StatusError').style.display = '';")
	response.Write("document.getElementsByName('wVarStatus')[0].focus();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.End()
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('StatusError').innerHTML = '';")
response.Write("document.getElementsByName('wVarStatus')[0].style.backgroundColor='#FFFFFF';")
response.Write("document.getElementById('StatusError').style.display = 'none';")
response.Write("</script>")

if wVarEmail = "" then
		response.Write("<script language='JavaScript'>")
		response.Write("document.getElementById('EmailError').innerHTML = 'Debe ingresar un E-mail';")
		response.Write("document.getElementById('EmailError').style.display = '';")
		response.Write("document.getElementsByName('wVarEmail')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
if(fregexp(wVarEmail, "^[_a-z0-9-]+(.[_a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)(.[a-z]{2,4})$" , 1)) = false then
		response.Write("<script type='text/javascript'>")
		response.Write("document.getElementById('EmailError').innerHTML = 'Debe ingresar un E-mail valido';")
		response.Write("document.getElementById('EmailError').style.display = '';")
		response.Write("document.getElementsByName('wVarEmail')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('EmailError').innerHTML = '';")
response.Write("document.getElementById('EmailError').style.display = 'none';")
response.Write("</script>")

if wVarTelefono = "" then
    response.Write("<script language='JavaScript'>")
	response.Write("document.getElementById('TelefonoError').innerHTML = '<font color=#8C0108><b>Debe ingresar un Telefono</b></font>';")
	response.Write("document.getElementsByName('wVarTelefono')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('TelefonoError').style.display = '';")
	response.Write("document.getElementsByName('wVarTelefono')[0].focus();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.end
end if
if(fregexp(wVarTelefono, "^([ 0-9+-]{4,20})$" , 1)) = false then
	response.Write("<script type='text/javascript'>")
	response.Write("document.getElementById('TelefonoError').innerHTML = '<font color=#8C0108><b>Debe ingresar un Telefono valido</b></font>';")
	response.Write("document.getElementsByName('wVarTelefono')[0].style.backgroundColor='#FFC1C1';")
	response.Write("document.getElementById('TelefonoError').style.display = '';")
	response.Write("document.getElementsByName('wVarTelefono')[0].focus();")
	response.Write("document.getElementById('loader').style.display='none';")
	response.Write("</script>")
	response.End()
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('TelefonoError').innerHTML = '';")
response.Write("document.getElementsByName('wVarTelefono')[0].style.backgroundColor='#FFFFFF';")
response.Write("document.getElementById('TelefonoError').style.display = 'none';")
response.Write("</script>")

if wVarComentario = "" then
		response.Write("<script language='JavaScript'>")
		response.Write("document.getElementById('ComentarioError').innerHTML = 'Debe ingresar un Asunto';")
		response.Write("document.getElementById('ComentarioError').style.display = '';")
		response.Write("document.getElementsByName('wVarComentario')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
if(fregexp(wVarComentario, "^([a-z A-ZÑñÁÉÍÓÚáéíóú]{2,60})$" , 1)) = false then
		response.Write("<script type='text/javascript'>")
		response.Write("document.getElementById('ComentarioError').innerHTML = 'Debe ingresar un Asunto valido';")
		response.Write("document.getElementById('ComentarioError').style.display = '';")
		response.Write("document.getElementsByName('wVarComentario')[0].focus();")
		response.Write("grecaptcha.reset();")
		response.Write("</script>")
		response.end
end if
response.Write("<script language='JavaScript'>")
response.Write("document.getElementById('ComentarioError').innerHTML = '';")
response.Write("document.getElementById('ComentarioError').style.display = 'none';")
response.Write("</script>")

if wVarCap = "" then
	response.Write("<script type='text/javascript'>")
	response.Write("document.getElementById('CaptaError').innerHTML = '<b>Debe completar el reCAPTCHA</b>';")
	response.Write("document.getElementById('CaptaError').style.display = '';")
	response.Write("document.getElementById('CaptaError').focus();")
	response.Write("grecaptcha.reset();")
	response.Write("</script>")
	response.end
end if

wVarSecret = "6LeK0vYiAAAAALwzVAzGy-oi1js7thTGhJvQ6w6b"
wVarIP = request.ServerVariables("REMOTE_ADDR")
wVarURL = "https://www.google.com/recaptcha/api/siteverify?" & "secret=" & wVarSecret & "&response=" & wVarCap & "&remoteip="  & wVarIP
Set objWinHttp1 = server.CreateObject("WinHttp.WinHttpRequest.5.1")
	objWinHttp1.SetTimeouts 300000, 300000, 300000, 300000
	objWinHttp1.Open "POST", wVarURL
	objWinHttp1.Send
	wVarCaptaResultado = objWinHttp1.ResponseText
Set objWinHttp1 = Nothing
wVarCaptaPase = fregexp(wVarCaptaResultado, "true", 1)
if wVarCaptaPase <> 1 then
	response.Write("<script type='text/javascript'>")
	response.Write("var wVarJsError = 'reCAPTCHA';")
	response.Write("document.getElementById('CaptaError').innerHTML = 'El reCAPTCHA es invalido';")
	response.Write("document.getElementById('CaptaError').style.display = '';")
	response.Write("document.getElementById('CaptaError').focus();")
	response.Write("grecaptcha.reset();")
	response.Write("</script>")
	response.end
end if
response.Write("<script type='text/javascript'>")
response.Write("document.getElementById('CaptaError').innerHTML = '';")
response.Write("document.getElementById('CaptaError').style.display = 'none';")
response.Write("</script>")

Set Mail = CreateObject("CDO.Message")
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.atlansystem.net"
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sashamfiorenza@gmail.com"
	Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "A21sd5sd4s6sd87!"
	Mail.Configuration.Fields.Update
	Mail.Subject = "FORMULARIO WEB > www.gaudium.com.ar"
	Mail.From = "sashamfiorenza@gmail.com"
	Mail.To = "sasha@atlansystem.com"
	Mail.HTMLBody = CharFix("<b>FORMULARIO WEB</b><br><br>""<b>Nombre: </b>" & wVarNombre & ",<br><b>Apellido: </b>" & wVarApellido & ",<br><b>País: </b>" & wVarPais & ",<br><b>Localidad: </b>" & wVarLocalidad & ",<br><b>Status: </b>" & wVarStatus & ",<br><b>E-mail: </b>" & wVarEmail & ",<br><b>Teléfono: </b>" & wVarTelefono & ",<br><b>Comentario: </b>" & wVarComentario)
	Mail.Send
Set Mail = Nothing

response.Write("<script type='text/javascript'>")
response.Write("sendformconf(0);")
response.Write("</script>")
%>