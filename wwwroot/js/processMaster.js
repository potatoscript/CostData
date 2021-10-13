var array_data = new Array();
var array_process_cost = new Array();
var ci,ri;
jQuery(document).ready(function () {

    if (document.getElementById("ProcessMasterId").value != 0)
        document.getElementById("statusMaster").innerText = "EDITING";


    jQuery('#table_process_master td')
        .click(function (event) {

           ci = jQuery(this).parent().children().index(this);
           ri = jQuery(this).parent().parent().children().index(this.parentNode);

        })

    var table = document.getElementById("table_process_master");
    var n = 0;
    array_data = [];
    array_process_cost = [];
    table.rows[0].onclick = function () {
        var row = table.rows[0];
        var rowCost = table.rows[1]; 
        if (String(row.cells[ci].innerHTML).indexOf("★") == -1) {
            array_data[n] = row.cells[ci].innerHTML;
            array_process_cost[n] = rowCost.cells[ci].innerHTML;
            row.cells[ci].innerHTML = "★" + row.cells[ci].innerHTML;
            n++;
        } else {
            var val = String(row.cells[ci].innerHTML).split("★");
            row.cells[ci].innerHTML = val.slice(1, 2);
            for (var a = 0; a < array_data.length; a++) {
                if (array_data[a] == row.cells[ci].innerHTML) {
                    array_data[a] = ""; //clear array_data and remove ★
                    array_process_cost[a] = "";
                }
            }
        }

        var c = 0;
        for (var i = 0; i < array_process_cost.length; i++) {
            if (array_process_cost[i]!="")
            c += parseFloat(array_process_cost[i]);
        }

        document.getElementById("process_cost").value = c;

    }



});

function delete_process(id) {
    var confirm_delete = confirm("Delete Process?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'Process/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                //set the id to zero for initial state
                var url = _url + "Process?p_doc_no=" + document.getElementById("doc_no").value + "&p_id=0";
                showPopup(url, 'Process Cost');
            }
        });
    }

}

function refresh() {
    var url = _url + "Process?p_doc_no=" + document.getElementById("doc_no").value + "&p_id=0";
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












