/*!
=========================================================
* LeadMark Landing page
=========================================================

* Copyright: 2019 DevCRUD (https://devcrud.com)
* Licensed: (https://devcrud.com/licenses)
* Coded by www.devcrud.com

=========================================================

* The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
*/


$(document).ready(function() {
    setTimeout(function() {
        $(".overlay").fadeIn(1500);
    },800);
});

$(document).ready(function() {
    setTimeout(function() {
        $(".overlay2").fadeIn(1500);
    },800);
});

// protfolio filters
$(window).on("load", function() {
    var t = $(".portfolio-container");
    t.isotope({
        filter: ".actuales",
        animationOptions: {
            duration: 750,
            easing: "linear",
            queue: !1
        }
    }), $(".filters a").click(function() {
        $(".filters .active").removeClass("active"), $(this).addClass("active");
        var i = $(this).attr("data-filter");
        return t.isotope({
            filter: i,
            animationOptions: {
                duration: 750,
                easing: "linear",
                queue: false
            }
        }), false
    })
})

function sendform(){
	var wVarNombre = document.getElementsByName('wVarNombre')[0].value.replace(/ /g, "%20");
	var wVarApellido = document.getElementsByName('wVarApellido')[0].value.replace(/ /g, "%20");
	var wVarPais = document.getElementsByName('wVarPais')[0].value.replace(/ /g, "%20");
	var wVarLocalidad = document.getElementsByName('wVarLocalidad')[0].value.replace(/ /g, "%20");
	var wVarStatus = document.getElementsByName('wVarStatus')[0].value.replace(/ /g, "%20");
	var wVarEMail = document.getElementsByName('wVarMail')[0].value.replace(/ /g, "%20");
	var wVarTelefono = document.getElementsByName('wVarTelefono')[0].value.replace(/ /g, "%20");
	var wVarComentario = document.getElementsByName('wVarComentario')[0].value.replace(/ /g, "%20");
	FAjax('assets/js/modulos.asp','divproceso','divproceso','wVarNombre='+wVarNombre+'&wVarApellido='+wVarApellido+'&wVarPais='+wVarPais+'&wVarLocalidad='+wVarLocalidad+'&wVarStatus='+wVarStatus+'&wVarEmail='+wVarEMail+'&wVarTelefono='+wVarTelefono+'&wVarComentario='+wVarComentario,'POST');
}

function sendformconf(x){
	if(x == 0){
		document.getElementById('formulsend').style.display='none';
		document.getElementById('formulsendhidden').style.display='';
		document.getElementById('formulsendhidden').innerHTML = "<font color='#FFF'><b>FORMULARIO ENVIADO CON EXITO<br>SI DESEA VOLVER A ENVIAR INFORMACION HAGA CLIC <a href='javascript:sendformconf(1);'><font color='#FFF'>AQUI</font></a>.</b></font>";
	}
	if(x == 1){	
		document.getElementsByName('wVarNombre')[0].value = '';
		document.getElementsByName('wVarApellido')[0].value = '';
		document.getElementsByName('wVarPais')[0].value = '';
		document.getElementsByName('wVarLocalidad')[0].value = '';
		document.getElementsByName('wVarStatus')[0].value = '';
		document.getElementsByName('wVarMail')[0].value = '';
		document.getElementsByName('wVarTelefono')[0].value = '';
		document.getElementsByName('wVarComentario')[0].value = '';
		document.getElementById('formulsendhidden').style.display='none';
		document.getElementById('formulsend').style.display='';
		var ele = document.getElementById('contacto');   
		window.scrollTo(ele.offsetLeft,ele.offsetTop);
	}
}