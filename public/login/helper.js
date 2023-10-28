// Fetch all the forms we want to apply custom Bootstrap validation styles to
const forms = document.querySelectorAll('.needs-validation');

/** 
 * Loop over the setting forms and add the listener
 * 
 */
Array.prototype.slice.call(forms).forEach((form) => {
    form.addEventListener('submit', (event) => {
        if (!form.checkValidity()) {
            event.preventDefault();
            event.stopPropagation();
        }else{
            var data = {};
            for(var i=0; i<event.target.length; i++){
                var name = event.target[i].id;
                var value = event.target[i].value;
                data[name] = value; 
            }

            data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";


            fetch(window.location.origin+'/login', {
                method: 'POST',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            })
            .then((response) => {
                if(response){
                    window.location.href = window.location.origin+'/administrator';
                }else{
                    var x = document.getElementById("alert-error");
                    x.classList.add("show");
                    setTimeout(function(){x.classList.remove("show"); }, 3000);
                }
            })
        }
        form.classList.add('was-validated');
        event.preventDefault();
        event.stopPropagation();
    }, false);
});