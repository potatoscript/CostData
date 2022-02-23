var array_data = new Array();
var array_process_cost = new Array();
var ci, ri;
jQuery(document).ready(function () {

    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };

    }

    jQuery('#table_tooling td')
        .click(function (event) {
            var table = document.getElementById("table_tooling");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);


            document.getElementById("description").value = table.rows[ri].cells[0].innerText;
            document.getElementById("source").value = table.rows[ri].cells[1].innerText;
            document.getElementById("qty").value = table.rows[ri].cells[2].innerText;
            document.getElementById("unit").value = table.rows[ri].cells[3].innerText;
            document.getElementById("price").value = table.rows[ri].cells[4].innerText;
            document.getElementById("od").value = table.rows[ri].cells[5].innerText;
            document.getElementById("ToolingId").value = table.rows[ri].cells[6].innerText;

            document.getElementById("statusTooling").innerText = "EDITING";

        })




});



function jQueryAjaxPostTooling(form, page) {
    jQuery.ajax({
        type: "POST",
        url: form.action,
        data: new FormData(form),
        contentType: false,
        processData: false,
        success: function (res) {
            if (res.isValid) {
                document.getElementById("ToolingMaster").click();
            }
        },
        error: function (err) {
            console.log(err)
        }
    })

    return false;
}



function delete_tooling(id) {
    var confirm_delete = confirm("Delete Tooling?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'Tooling/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                document.getElementById("ToolingMaster").click();
            }
        });
    }

}

function refreshTooling() {
    document.getElementById("ToolingMaster").click();
}