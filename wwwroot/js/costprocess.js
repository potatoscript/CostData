var array_data = new Array();
var array_process_cost = new Array();
var ci, ri;
jQuery(document).ready(function () {

    jQuery('#table_process_master_data td')
        .mouseenter(function (event) {
            var table = document.getElementById("table_process_master_data");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);

            document.getElementById("process_name_data").value = table.rows[ri].cells[0].innerText;
            document.getElementById("process_type_data").value = table.rows[ri].cells[1].innerText;
            document.getElementById("overhead_cost_data").value = table.rows[ri].cells[2].innerText;
            document.getElementById("machine_cost_data").value = table.rows[ri].cells[3].innerText;
            document.getElementById("labor_cost_data").value = table.rows[ri].cells[4].innerText;
            document.getElementById("total_cost_data").value = table.rows[ri].cells[5].innerText;

        })
        .click(function () {
            document.getElementById("submit_costprocess").click();
        })

    /*
    var table = document.getElementById("table_process_master_data");
    var n = 0;
    array_data = [];
    array_process_cost = [];
    for (var k = 1; k < table.rows.length; k++) {
        table.rows[k].onclick = function () {

                var row = table.rows[this.rowIndex];
                if (String(row.cells[0].innerHTML).indexOf("★") == -1) {
                    array_data[n] = row.cells[0].innerHTML;
                    array_process_cost[n] = row.cells[5].innerText;
                    row.cells[0].innerHTML = "★" + row.cells[0].innerHTML;
                    n++;
                } else {
                    var val = String(row.cells[0].innerHTML).split("★");
                    row.cells[0].innerHTML = val.slice(1, 2);
                    for (var a = 0; a < array_data.length; a++) {
                        if (array_data[a] == row.cells[0].innerHTML) {
                            array_data[a] = ""; //clear array_data and remove ★
                            array_process_cost[a] = "";
                        }
                    }
                }

                var c = 0;
                for (var i = 0; i < array_process_cost.length; i++) {
                    if (array_process_cost[i] != "")
                        c += parseFloat(array_process_cost[i]);
                }

                document.getElementById("process_cost").value = c;

            jQuery.ajax({
                type: "POST",
                url: _url + 'CostProcess/Save',
                data: jQuery.param({
                    confirm: true
                }),
                success: function (res) {
                    alert(_url);
                    //set the id to zero for initial state
                    var url = _url + "CostProcess/Index?p_doc_no=" + document.getElementById("doc_no").value + "&p_od=" + document.getElementById("od").value;
                    showPopup(url, 'Process Cost');
                }
            });
            
        }
    }
    */

    //get the total process cost
    var table = document.getElementById("table_process_data");
    var totalcost = 0;
    for (var k = 1; k < table.rows.length; k++) {
        totalcost += parseFloat(table.rows[k].cells[5].innerText);
    }
    document.getElementById("process_cost").value = totalcost;
    document.getElementById("direct_process_cost").value = totalcost;// in the main page


});



function jQueryAjaxPostCostProcess(form, page) {
        jQuery.ajax({
            type: "POST",
            url: form.action,
            data: new FormData(form),
            contentType: false,
            processData: false,
            success: function (res) {
                if (res.isValid) {
                        document.getElementById("direct_process_cost").click();
                }
            },
            error: function (err) {
                console.log(err)
            }
        })

    return false;
}



function delete_process_data(id) {
    var confirm_delete = confirm("Delete Process?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'CostProcess/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                //set the id to zero for initial state
                var url = _url + "CostProcess/Index?p_doc_no=" + document.getElementById("doc_no").value + "&p_od=" + document.getElementById("od").value;
                showPopup(url, 'Process Cost Data');
            }
        });
    }

}