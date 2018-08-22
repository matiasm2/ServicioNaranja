$(document).ready(function() {
    //Id HTML del contenedor del campo "Cuerpo del correo"
    id = '#master_DefaultContent_rts_s6441_f23876c';
    $(id)[0].innerHTML=$(id)[0].innerHTML.replace(new RegExp('http://novatickets/jpg', 'g'),'data:image/jpg;base64,');
    $(id)[0].innerHTML=$(id)[0].innerHTML.replace(new RegExp('http://novatickets/png', 'g'),'data:image/png;base64,');
})