var array_data = new Array();
var array_process_cost = new Array();
var ci, ri;
jQuery(document).ready(function () {

    //get the total process cost
    var table = document.getElementById("table_process_data");
    var totalcost = 0;
    var total_overhead_cost = 0;
    var total_machine_cost = 0;
    var total_labor_cost = 0;
    for (var k = 1; k < table.rows.length; k++) {//-1 because the last raw was the total
        total_overhead_cost += parseFloat(table.rows[k].cells[2].innerText);
        total_machine_cost += parseFloat(table.rows[k].cells[3].innerText);
        total_labor_cost += parseFloat(table.rows[k].cells[4].innerText);
        totalcost += parseFloat(table.rows[k].cells[5].innerText);
    }
    document.getElementById("total_overhead_cost").value = total_overhead_cost.toFixed(4);
    document.getElementById("total_machine_cost").value = total_machine_cost.toFixed(4);
    document.getElementById("total_labor_cost").value = total_labor_cost.toFixed(4);
    document.getElementById("process_cost").value = totalcost.toFixed(4);
    document.getElementById("direct_process_cost").value = totalcost.toFixed(4);// in the main page
    document.getElementById("overhead_cost").value = total_overhead_cost.toFixed(4);
    document.getElementById("machine_cost").value = total_machine_cost.toFixed(4);
    document.getElementById("labor_cost").value = total_labor_cost.toFixed(4);

    var price_sgd = parseFloat(document.getElementById("target_price_sgd").value);
    document.getElementById("overhead_cost_p").value = ((total_overhead_cost / price_sgd)*100).toFixed(1);
    document.getElementById("machine_cost_p").value = ((total_machine_cost / price_sgd) * 100).toFixed(1);
    document.getElementById("labor_cost_p").value = ((total_labor_cost / price_sgd) * 100).toFixed(1);

    var process_cost_sub_total = total_overhead_cost + total_machine_cost + total_labor_cost;
    document.getElementById("process_cost_sub_total").value = process_cost_sub_total.toFixed(4);
    document.getElementById("process_cost_sub_total_p").value = ((process_cost_sub_total / price_sgd) * 100).toFixed(1);


    

    jQuery("#processtypes").change(function () {
        document.getElementById("product_type").value = document.getElementById("processtypes").value;



            jQuery.ajax({
                type: "GET",
                url: _url + 'CostProcess/Index',
                data: jQuery.param({
                    p_doc_no: document.getElementById("doc_no").value,
                    p_od: document.getElementById("od").value,
                    p_process_type: document.getElementById("product_type").value,
                    p_rubber_weight: document.getElementById("total_rubber_weight").value
                }),
                success: function (res) {
                    jQuery("#form-modal .modal-body").html(res);
                    jQuery("#form-modal .modal-title").html("Process Cost Data");
                    jQuery("#form-modal").modal('show');

                    //make sure the select value stay at the selected value after select
                    document.getElementById("processtypes").value = document.getElementById("product_type").value;
                }
            })

        

    })

    jQuery('#table_process_master_data td')
        .mouseenter(function (event) {
            var table = document.getElementById("table_process_master_data");
            ci = jQuery(this).parent().children().index(this);
            ri = jQuery(this).parent().parent().children().index(this.parentNode);

            var total_cost_data = parseFloat(table.rows[ri].cells[5].innerText);
            var rubber_weight = parseFloat(document.getElementById("total_rubber_weight").value);
            if (table.rows[ri].cells[1].innerText == "Rubber") {
                total_cost_data = parseFloat(table.rows[ri].cells[5].innerText) * rubber_weight;
            }

            document.getElementById("process_name_data").value = table.rows[ri].cells[0].innerText;
            document.getElementById("process_type_data").value = table.rows[ri].cells[1].innerText;
            document.getElementById("overhead_cost_data").value = table.rows[ri].cells[2].innerText;
            document.getElementById("machine_cost_data").value = table.rows[ri].cells[3].innerText;
            document.getElementById("labor_cost_data").value = parseFloat(table.rows[ri].cells[4].innerText).toFixed(4);
            document.getElementById("total_cost_data").value = total_cost_data.toFixed(4);

            

        })
        .click(function () {
            document.getElementById("submit_costprocess").click();
        })

    calculate_cost();

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

    if (document.getElementById("product_type").value.length>2)
    document.getElementById("processtypes").value = document.getElementById("product_type").value;


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
                    document.getElementById("product_type").value = document.getElementById("processtypes").value;
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