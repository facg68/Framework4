    function swal_Popup(titulo, mensaje, icono) {
        /*
            Los iconos son: warning, error, success, info y question
        */
        swal_fire({
            title: titulo,
            text: mensaje,
            icon: icono
        })       
    }

    function swal_DraggablePopup(titulo, mensaje, icono) {
        /*
            Los iconos son: warning, error, success, info y question
        */
        swal_fire({
            title: titulo,
            text: mensaje,
            icon: icono,
            draggable: true
        })       
    }    

    function swal_Mensaje(mensaje) {
        swal_fire(mensaje)
    } 

    function swal_PopPlus(titulo, mensaje, icono, footer) {
        /*
            Los iconos son: warning, error, success, info y question
        */
        swal_fire({
            icon: icono,
            title: titulo,
            text: mensaje,
            footer: footer
        })
    }  

    function swal_imagen(href, altura, texto_alt) {
        swal_fire({
            imageUrl: href,
            imageHeight: altura,
            imageAlt: texto_alt
        })
    }

    function swal_Dialog3(titulo, nombre_boton_confirmar, nombre_boton_denegar, mensaje_exito, mensaje_error) {
        swal_fire({
            title: titulo,
            showDenyButton: true,
            showCancelButton: true,
            confirmButtonText: nombre_boton_confirmar,
            denyButtonText: nombre_boton_denegar
        }).then((result) => {
            if (result.isConfirmed) {
                swal_fire("Proceso Realizado!", "", mensaje_exito);
            } else if (result.isDenied) {
                swal_fire("El proceso no se realizó", "", mensaje_error);
            }
        })
    }    

    function swal_borrar(titulo, mensaje, href) {
        swal_fire({
            title: titulo,
            text: mensaje,
            icon: "warning",
            showCancelButton: true,
            confirmButtonColor: "#3085d6",
            cancelButtonColor: "#d33",
            confirmButtonText: "Registro borrado!"
        }).then((result) => {
            if (result.isConfirmed) {
                    /* Poner aqui la redireccion */
                    swal_fire({
                        title: "Borrado!",
                        text: "El registro ha sido eliminado.",
                        icon: "success"
                    });
            }
        })
    }   

    function swal_MensajeConImagen(titulo, mensaje, href, alt_imagen) {
        /*
            Los iconos son: warning, error, success, info y question
        */
        swal_fire({
            title: titulo,
            text: mensaje,
            imageUrl: href,
            imageWidth: 400,
            imageHeight: 200,
            imageAlt: alt_imagen
        })      
    }