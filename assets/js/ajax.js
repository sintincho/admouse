function creaAjax(){
  var objetoAjax=false;
  try {
   /*Para navegadores distintos a internet explorer*/
   objetoAjax = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
     /*Para explorer*/
     objetoAjax = new ActiveXObject("Microsoft.XMLHTTP");
     } 
     catch (E) {
     objetoAjax = false;
   }
  }

  if (!objetoAjax && typeof XMLHttpRequest!='undefined') {
   objetoAjax = new XMLHttpRequest();
  }
  return objetoAjax;
}
 function FAjax (url,capaResultado,capaEspere,valores,metodo)
{
   var ajax=creaAjax();
   var capaContenedora = document.getElementById(capaResultado);
   var capaMensaje = document.getElementById(capaEspere);

/*Creamos y ejecutamos la instancia si el metodo elegido es POST*/
 if(metodo.toUpperCase()=='POST'){
    ajax.open ('POST', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
		}
         else if (ajax.readyState==4){
			if(ajax.status==200)
            {
				capaMensaje.innerHTML = "Cargando!!!!!!";
			    capaContenedora.innerHTML=ajax.responseText;
				 //crear elemento div de ejecucion js
				 var scs=ajax.responseText.extractScript();
				 scs.evalScript();
			}
            else if(ajax.status==404)
                 {

                     capaMensaje.innerHTML = "La direccion existe";
                 }
            else if(ajax.status==500)
                 {
					document.getElementById('varerror').innerHTML = ajax.responseText;
					document.getElementById('popup0').innerHTML = "<div style='text-align:center;z-index:1600;'><font color='FFFFFF'>SE PRODUJO UN ERROR<br>POR FAVOR INDIQUE SU CORREO ELECTRONICO<br>NOS COMUNICAREMOS CON UD A LA BREVEDAD<br>PARA RESOLVER EL MISMO</font><br><input type='text' for='E-mail' name='wVarMailInfoError' id='wVarMailInfoError' style='width:280px;' onKeyPress='if(event.keyCode==13) senderr(0);'><br><div id='DivMailInfoError' style='display:none;'></div><a href='javascript:senderr(0);'><font color='FFFFFF'>│ENVIAR│</font></a></div>"
					document.getElementById('wVarMailInfoError').focus();
                    capaMensaje.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
             else
                 {
					document.getElementById('varerror').innerHTML = ajax.responseText;
					document.getElementById('popup0').innerHTML = "<div style='text-align:center;z-index:1600;'><font color='FFFFFF'>SE PRODUJO UN ERROR<br>POR FAVOR INDIQUE SU CORREO ELECTRONICO<br>NOS COMUNICAREMOS CON UD A LA BREVEDAD<br>PARA RESOLVER EL MISMO</font><br><input type='text' for='E-mail' name='wVarMailInfoError' id='wVarMailInfoError' style='width:280px;' onKeyPress='if(event.keyCode==13) senderr(0);'><br><div id='DivMailInfoError' style='display:none;'></div><a href='javascript:senderr(0);'><font color='FFFFFF'>│ENVIAR│</font></a></div>"
					document.getElementById('wVarMailInfoError').focus();
					capaMensaje.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-type','application/x-www-form-urlencoded');
    ajax.send(valores);
    return;
}

/*Creamos y ejecutamos la instancia si el metodo elegido es GET*/
if (metodo.toUpperCase()=='GET'){

    ajax.open ('GET', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
                 capaMensaje.innerHTML="<div align=center class=NotificacionRoja>Trabajando, por favor espere...</div>";
         }
         else if (ajax.readyState==4){
            if(ajax.status==200){ 
                 capaContenedora.innerHTML.innerHTML=ajax.responseText; 
				 capaMensaje.innerHTML="";
            }
            else if(ajax.status==404)
                 {

                     capaContenedora.innerHTML = "<div align=center class=NotificacionRoja>Se ha realizado una petición que no es posible contestarla. Este error debe informarse al Administrador.</div>";
                 }
            else if(ajax.status==500)
                 {

                     capaMensaje.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
            else
                 {
                     capaContenedora.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
    ajax.send(null);

    return
}
}


///////////////////////////////////////


 function FAjaxRefresh (url,capaResultado,capaEspere,valores,metodo)
{
   var ajax=creaAjax();
   var capaContenedora = document.getElementById(capaResultado);
   var capaMensaje = document.getElementById(capaEspere);

/*Creamos y ejecutamos la instancia si el metodo elegido es POST*/
 if(metodo.toUpperCase()=='POST'){
        document.write="hola pepe";
    ajax.open ('POST', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
                 capaMensaje.innerHTML="<div align=center class=NotificacionRoja>Actualizando...</div>";
         }
         else if (ajax.readyState==4){
            if(ajax.status==200)
            {
                 capaMensaje.innerHTML = "";
				 capaContenedora.innerHTML=ajax.responseText; 
            }
            else if(ajax.status==404)
                 {

                     capaMensaje.innerHTML = "La direccion existe";
                 }
            else if(ajax.status==500)
                 {

                     capaMensaje.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
             else
                 {
                     capaMensaje.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-type','application/x-www-form-urlencoded');
    ajax.send(valores);
    return;
}


/*Creamos y ejecutamos la instancia si el metodo elegido es GET*/
if (metodo.toUpperCase()=='GET'){

    ajax.open ('GET', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
                 capaMensaje.innerHTML="<div align=center class=NotificacionRoja>Trabajando, por favor espere...</div>";
         }
         else if (ajax.readyState==4){
            if(ajax.status==200){ 
                 capaContenedora.innerHTML.innerHTML=ajax.responseText; 
				 capaMensaje.innerHTML="";
            }
            else if(ajax.status==404)
                 {

                     capaContenedora.innerHTML = "<div align=center class=NotificacionRoja>Se ha realizado una petición que no es posible contestarla. Este error debe informarse al Administrador.</div>";
                 }
            else if(ajax.status==500)
                 {

                     capaMensaje.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
            else
                 {
                     capaContenedora.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
    ajax.send(null);
    return
}
}

///////////////////////////////////////


 function FAjaxEdit (url,capaResultado,capaEspere,valores,metodo)
{
   var ajax=creaAjax();
   var capaContenedora = document.getElementById(capaResultado);
   var capaMensaje = document.getElementById(capaEspere);

/*Creamos y ejecutamos la instancia si el metodo elegido es POST*/
 if(metodo.toUpperCase()=='POST'){
        document.write="hola pepe";
    ajax.open ('POST', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
         }
         else if (ajax.readyState==4){
            if(ajax.status==200)
            {

				 capaContenedora.innerHTML=ajax.responseText; 
            }
            else if(ajax.status==404)
                 {

                     capaContenedora.innerHTML = "La direccion existe";
                 }
            else if(ajax.status==500)
                 {
                     capaContenedora.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
             else
                 {
                     capaContenedora.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-type','application/x-www-form-urlencoded');
    ajax.send(valores);
    return;
}


/*Creamos y ejecutamos la instancia si el metodo elegido es GET*/
if (metodo.toUpperCase()=='GET'){

    ajax.open ('GET', url, true);
    ajax.onreadystatechange = function() {
         if (ajax.readyState==1) {
                 capaMensaje.innerHTML="<div align=center class=NotificacionRoja>Trabajando, por favor espere...</div>";
         }
         else if (ajax.readyState==4){
            if(ajax.status==200){ 
                 capaContenedora.innerHTML.innerHTML=ajax.responseText; 
            }
            else if(ajax.status==404)
                 {
                     capaContenedora.innerHTML = "<div align=center class=NotificacionRoja>Se ha realizado una petición que no es posible contestarla. Este error debe informarse al Administrador.</div>";
                 }
            else if(ajax.status==500)
                 {
                     capaContenedora.innerHTML = "Se produjo un error en el servidor (error 500). Contacte al administrador.";
                 }
            else
                 {
                     capaContenedora.innerHTML = "Error ".ajax.status;
                 }
        }
    }
    ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
    ajax.send(null);
    return
}
}
        var tagScript = '(?:<script.*?>)((\n|\r|.)*?)(?:<\/script>)';
		//podria hacer un remplace y crear el try
        /**
        * Eval script fragment
        * @return String
        */
        String.prototype.evalScript = function()
        {
                return (this.match(new RegExp(tagScript, 'img')) || []).evalScript();
        };
        /**
        * strip script fragment
        * @return String
        */
        String.prototype.stripScript = function()
        {
                return this.replace(new RegExp(tagScript, 'img'), '');
        };
        /**
        * extract script fragment
        * @return String
        */
        String.prototype.extractScript = function()
        {
                var matchAll = new RegExp(tagScript, 'img');
                return (this.match(matchAll) || []);
        };
        /**
        * Eval scripts
        * @return String
        */
        Array.prototype.evalScript = function(extracted)
        {
                var s=this.map(function(sr){
                         var sc=(sr.match(new RegExp(tagScript, 'im')) || ['', ''])[1];
                         if(window.execScript){
							  window.execScript(sc);
                         }
                        else
                       {
							sc = 'try {\n'+sc+'\n}\n catch(err) {\n jserror(err.message, wVarJsError);\n}'
							window.setTimeout(sc,0);
						}
                });
                return true;
        };
        /**
        * Map array elements
        * @param {Function} fun
        * @return Function
        */
//        Array.prototype.map = function(fun)
//        {
 //               if(typeof fun!=="function"){return false;}
  //              var i = 0, l = this.length;
   //             for(i=0;i<l;i++)
    //            {
     //                   fun(this[i]);
       //         }
        //        return true;
        //};  