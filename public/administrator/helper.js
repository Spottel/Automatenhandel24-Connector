/** 
 * Load Error Log Table
 * 
 */
$(document).ready(function () {
    var table = $('#errortable').DataTable({
        processing: true,
        serverSide: true,
        ajax: {
            url: window.location.origin+'/geterrors',
            type: 'POST',
            data: function (d) {
                d.sendkey = 'asdn9n34b374b8734vasdv7v73v324';

                d.filterAll = $('#error_log_option1').is(':checked')
                d.filterError = $('#error_log_option2').is(':checked');
                d.filterSuccess = $('#error_log_option3').is(':checked');
            }
        },
        'columnDefs': [
            {
               'targets': 0,
               'checkboxes': {
                  'selectRow': true
               }
            }
        ],
        order: [[1, 'desc']],
        columns: [
            { data: 'id' },
            { data: 'create_date' },
            { data: 'error_type' },
            { data: 'error_module' },
            { data: 'error_message' },
            { data: 'action', orderable: false}
        ]
    }).on( 'xhr.dt', function () {
        // On select checkboxes
        $("input[type='checkbox']").off('change').on( 'change', function () {
            var rows_selected = table.column(0).checkboxes.selected();
            if(rows_selected.length == 0){
                $("#error_log_delete_show_modal_button").prop("disabled", true);
            }else{
                $("#error_log_delete_show_modal_button").removeAttr('disabled');
            }
        });

        $(document).on('change', '.dt-checkboxes', function() {
            var rows_selected = table.column(0).checkboxes.selected();
            if(rows_selected.length == 0){
                $("#error_log_delete_show_modal_button").prop("disabled", true);
            }else{
                $("#error_log_delete_show_modal_button").removeAttr('disabled');
            }
        });

        // Click Actions column delete button
        $('#errortable tbody').on( 'click', 'button', function () {
            var data = {};
            data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";
            data['ids'] = [$(this).attr("data-element-id")];
            
            fetch(window.location.origin+'/deleteerror', {
                method: 'POST',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            })
            .then(response => response.json())
            .then((response) => {
                if(response){
                    table
                        .row( $(this).parents('tr') )
                        .remove()
                        .draw();
                }
            });

        });
    });

    // Set Filter for error log table
    $( "#error_log_option1" ).on( "click", function() {
        $('#errortable').DataTable().ajax.reload();
    });

    $( "#error_log_option2" ).on( "click", function() {
        $('#errortable').DataTable().ajax.reload();
    });

    $( "#error_log_option3" ).on( "click", function() {
        $('#errortable').DataTable().ajax.reload();
    });

    // On Change error log table
    table.on( 'draw', function () {
        // On select checkboxes
        $('.dt-checkboxes').prop( "checked", false );
        $("input[type='checkbox']").prop( "checked", false );
        $("#error_log_delete_show_modal_button").prop("disabled", true); 
    });


    // Show error log delete modal
    const myModalEl = document.getElementById('error_log_delete_modal')
    const modal = new mdb.Modal(myModalEl);

    $( "#error_log_delete_show_modal_button" ).on( "click", function() {
        modal.show();
    });


    // Error log table delete button
    $( "#error_log_delete_button" ).on( "click", function() {
      var rows_selected = table.column(0).checkboxes.selected();

      if(rows_selected.length != 0){
        // Iterate over all selected checkboxes
        var ids = [];

        $.each(rows_selected, function(index, rowId){
            ids.push(rowId);
        });


        var data = {};
        data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";
        data['ids'] = ids;
        
        fetch(window.location.origin+'/deleteerror', {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        })
        .then(response => response.json())
        .then((response) => {
            if(response){
                table.ajax.reload();
                modal.hide()
            }
        });
      }
    });
});


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
                        
                if(event.target[i].type == "checkbox"){
                    if(event.target[i].checked){
                        value = 'true';
                    }else{
                        value = 'false';
                    }
                }

                data[name] = value; 
            }

            data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";


            fetch(window.location.origin+'/savesettings', {
                method: 'POST',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            })
            .then(response => response.json())
            .then((response) => {
                if(response){
                    var x = document.getElementById("alert-success");
                    x.classList.add("show");
                    setTimeout(function(){ x.classList.remove("show");}, 3000);
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

/**
 * Get the setting data to load that in the setting forms 
 * 
 */
var data = {};
data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";

fetch(window.location.origin+'/getsettings', {
    method: 'POST',
    headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(data)
})
.then(response => response.json())
.then((response) => {
    for(var i=0; i<response.length; i++){
        if(response[i].setting_value == 'true'){
            document.getElementById(response[i].setting_name).checked = true;
        }else if(response[i].setting_value == 'false'){
            document.getElementById(response[i].setting_name).checked = false;
        }else{
            document.getElementById(response[i].setting_name).value = response[i].setting_value
        }

        if(response[i].setting_name == "hubspotconnectionaddress"){
            document.getElementById("connectorhubspot").href=response[i].setting_value;
        }
    }
})

/** 
 * Click on logout button
 * 
 */
document.getElementById("logoutButton").addEventListener("click", function(){ 
    document.cookie = 'token=;expires=Thu, 01 Jan 1970 00:00:01 GMT;';
    window.location.href = window.location.origin+'/login';
});


/** 
 * Click on mailer test button
 * 
 */
document.getElementById("mailertestbutton").addEventListener("click", function(){ 
    var data = {};
    data['sendkey'] = "asdn9n34b374b8734vasdv7v73v324";
    data['mailerhost'] = document.getElementById('mailerhost').value;
    data['mailerport'] = document.getElementById('mailerport').value;
    data['mailersecure'] = document.getElementById('mailersecure').checked;
    data['mailergoogle'] = document.getElementById('mailergoogle').checked;
    data['maileruser'] = document.getElementById('maileruser').value;
    data['mailerpassword'] = document.getElementById('mailerpassword').value;
    data['mailertestaddress'] = document.getElementById('mailertestaddress').value;

    if(document.getElementById('mailertestaddress').value != ""){
        document.getElementById('mailertestaddresserror').style.display = 'none';

        fetch(window.location.origin+'/testmailer', {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        })
        .then(response => response.json())
        .then((response) => {
            if(response){
                document.getElementById('mailertestaddress').value = "";
            }else{
                document.getElementById('mailertestaddresserror').style.display = 'block';
            }
        });
    }else{
        document.getElementById('mailertestaddresserror').style.display = 'block';
    }
});