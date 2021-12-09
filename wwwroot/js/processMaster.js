var array_data = new Array();
var array_process_cost = new Array();
var ci,ri;
jQuery(document).ready(function () {

    if (document.getElementById("ProcessMasterId").value != 0)
        document.getElementById("statusMaster").innerText = "EDITING";


    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };

        s[t].onkeyup = function () {

            var overhead = parseFloat(document.getElementById("overhead_cost_master").value);
            var machine = parseFloat(document.getElementById("machine_cost_master").value);
            var labor = parseFloat(document.getElementById("labor_cost_master").value);

            
            document.getElementById("total_cost").value = (overhead + machine + labor).toFixed(4);
        };

    }


    jQuery("#process_type_master").change(function () {
        document.getElementById("processtype").value = document.getElementById("process_type_master").value;

        jQuery.ajax({
            type: "GET",
            url: _url + 'ProcessMaster/Index',
            data: jQuery.param({
                p_doc_no: "-",
                p_od: 0,
                p_type: document.getElementById("processtype").value
            }),
            success: function (res) {
                
                jQuery("#form-modal .modal-body").html(res);
                jQuery("#form-modal .modal-title").html("Process Cost Data");
                jQuery("#form-modal").modal('show');

                document.getElementById("process_type_master").value = document.getElementById("processtype").value;
            }
        })



    })

    jQuery('#table_process_master td')
        .click(function (event) {
            var table = document.getElementById("table_process_master");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);


                document.getElementById("process_name_master").value = table.rows[ri].cells[0].innerText;
                document.getElementById("process_type_master").value = table.rows[ri].cells[1].innerText;
                document.getElementById("od_min").value = table.rows[ri].cells[2].innerText;
                document.getElementById("od_max").value = table.rows[ri].cells[3].innerText;
            document.getElementById("overhead_cost_master").value = table.rows[ri].cells[4].innerText;
            document.getElementById("machine_cost_master").value = table.rows[ri].cells[5].innerText;
            document.getElementById("labor_cost_master").value = table.rows[ri].cells[6].innerText;

                document.getElementById("total_cost").value = table.rows[ri].cells[7].innerText;
                document.getElementById("ProcessMasterId").value = table.rows[ri].cells[8].innerText;

            document.getElementById("statusMaster").innerText = "EDITING";

        })


});

function delete_process_master(id) {
    var confirm_delete = confirm("Delete Process?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'ProcessMaster/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                refreshProcessMaster();
            }
        });
    }

}

function refreshProcessMaster() {
    //document.getElementById("ProcessMaster").click();
    jQuery.ajax({
        type: "GET",
        url: _url + 'ProcessMaster/Index',
        data: jQuery.param({
            p_doc_no: "-",
            p_od: 0,
            p_type: document.getElementById("processtype").value
        }),
        success: function (res) {

            jQuery("#form-modal .modal-body").html(res);
            jQuery("#form-modal .modal-title").html("Process Cost Data");
            jQuery("#form-modal").modal('show');

            document.getElementById("process_type_master").value = document.getElementById("processtype").value;
        }
    })
}


function Total_process_cost() {

    var table = document.getElementById("table_process_list");
    var total_process_cost = 0;
    for (var k = 1; k < table.rows.length; k++) {
        total_process_cost += parseFloat(table.rows[k].cells[2].innerText);
    }
    document.getElementById("process_cost").value = total_process_cost.toFixed(4);
}












