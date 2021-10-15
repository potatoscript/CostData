var array_data = new Array();
var array_process_cost = new Array();
var ci,ri;
jQuery(document).ready(function () {

    if (document.getElementById("ProcessMasterId").value != 0)
        document.getElementById("statusMaster").innerText = "EDITING";


    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };


    }
    for (var s = document.getElementsByClassName("processmaster_field"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {
            
            var overhead = parseFloat(document.getElementById("overhead_cost").value);
            var machine = parseFloat(document.getElementById("machine_cost").value);
            var labor = parseFloat(document.getElementById("labor_cost").value);
            
            document.getElementById("total_cost").value = (overhead+machine+labor).toFixed(5);
        };

    }


    jQuery('#table_process_master td')
        .click(function (event) {
            var table = document.getElementById("table_process_master");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);

            if (ci != 0 && String(table.rows[ri].cells[0].innerText).indexOf('★')==-1) {
                document.getElementById("process_name_master").value = table.rows[ri].cells[0].innerText;
                document.getElementById("process_type_master").value = table.rows[ri].cells[1].innerText;
                document.getElementById("overhead_cost").value = table.rows[ri].cells[2].innerText;
                document.getElementById("machine_cost").value = table.rows[ri].cells[3].innerText;
                document.getElementById("labor_cost").value = table.rows[ri].cells[4].innerText;
                document.getElementById("total_cost").value = table.rows[ri].cells[5].innerText;
                document.getElementById("ProcessMasterId").value = table.rows[ri].cells[6].innerText;
            } else {
                document.getElementById("process_name_master").value = "-";
                document.getElementById("process_type_master").value = "-";
                document.getElementById("overhead_cost").value = 0;
                document.getElementById("machine_cost").value = 0;
                document.getElementById("labor_cost").value = 0;
                document.getElementById("total_cost").value = 0;
                document.getElementById("ProcessMasterId").value = 0;
            }
            

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
                //set the id to zero for initial state
                var url = _url + "ProcessMaster/Index?p_doc_no=" + document.getElementById("doc_no").value + "&p_od=" + document.getElementById("od").value;
                showPopup(url, 'Process Cost');
            }
        });
    }

}

function refresh() {
    var url = _url + "ProcessMaster?p_doc_no=" + document.getElementById("doc_no").value + "&p_od=" + document.getElementById("od").value;
    showPopup(url, 'Process Cost');
}


function Total_process_cost() {

    var table = document.getElementById("table_process_list");
    var total_process_cost = 0;
    for (var k = 1; k < table.rows.length; k++) {
        total_process_cost += parseFloat(table.rows[k].cells[2].innerText);
    }
    document.getElementById("process_cost").value = total_process_cost.toFixed(4);
}












