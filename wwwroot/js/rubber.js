var array_data = new Array();
var array_process_cost = new Array();
var ci, ri;
jQuery(document).ready(function () {

    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };

    }

    jQuery('#table_rubber td')
        .click(function (event) {
            var table = document.getElementById("table_rubber");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);


            document.getElementById("material_name").value = table.rows[ri].cells[0].innerText;
            document.getElementById("price_kg").value = table.rows[ri].cells[1].innerText;
            document.getElementById("mixing_process_cost").value = table.rows[ri].cells[2].innerText;
            document.getElementById("weight_g").value = table.rows[ri].cells[3].innerText;
            document.getElementById("yield_rate").value = table.rows[ri].cells[4].innerText;
            document.getElementById("RubberId").value = table.rows[ri].cells[5].innerText;



        })

   


});



function jQueryAjaxPostRubber(form, page) {
    jQuery.ajax({
        type: "POST",
        url: form.action,
        data: new FormData(form),
        contentType: false,
        processData: false,
        success: function (res) {
            if (res.isValid) {
                document.getElementById("RubberMaster").click();
            }
        },
        error: function (err) {
            console.log(err)
        }
    })

    return false;
}



function delete_rubber(id) {
    var confirm_delete = confirm("Delete Process?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'Rubber/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                document.getElementById("RubberMaster").click();
            }
        });
    }

}

function refreshRubber() {
    document.getElementById("RubberMaster").click();
}