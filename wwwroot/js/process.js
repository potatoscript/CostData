jQuery(document).ready(function () {

    if (document.getElementById("ProcessId").value != 0)
        document.getElementById("status").innerText = "EDITING";

    Total_process_cost();
    //set the total process_cost from process page into cost page's process cost and recalclate
    document.getElementById("direct_process_cost").value = document.getElementById("process_cost").value;
    calculate_cost();

    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };


    }
    for (var s = document.getElementsByClassName("process_field"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {
            Total_process_cost();
        };

    }


    jQuery('#table_process_list td')
        .click(function (event) {

            var CI = jQuery(this).parent().children().index(this);
            var RI = jQuery(this).parent().parent().children().index(this.parentNode);
            var table = document.getElementById("table_process_list");
            var docno = document.getElementById("doc_no").value;
            if (CI != 3) {
                var url = _url + "Process?p_doc_no=" + docno + "&p_id=" + table.rows[RI].cells[0].innerText;
                showPopup(url, 'Process Cost');
            }

        })


    jQuery("#master_process_").click(function () {

        jQuery("#table_process_master")
            .show()
            .css({
                "position": "absolute",
                "top": event.clientY - 80 + "px",
                "left": event.clientX - 360 + "px"

            })
        jQuery('#table_process_master td')
            .click(function (event) {//for multiple select on process

                var RI = jQuery(this).parent().parent().children().index(this.parentNode);

                var table = document.getElementById("table_process_master");
                document.getElementById("master_process_").value =
                    table.rows[RI].cells[0].innerText;

                window.location.href = _url + "Process/Read?id=" + table.rows[RI].cells[2].innerText;


                jQuery("#table_process_master").hide();

            })
        jQuery('#table_process_master th')
            .click(function (event) {//for multiple select on process
                jQuery("#table_process_master").hide();

            })
    })



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



function Total_process_cost() {

    var workingDay = parseFloat(document.getElementById("p_working_day").value);
    var actualWorkingTime = parseFloat(document.getElementById("p_working_time_day").value);
    document.getElementById("p_working_time_month").value = workingDay * actualWorkingTime;
    var actualWorkingTimeMonth = parseFloat(document.getElementById("p_working_time_month").value);

    var shift = parseFloat(document.getElementById("p_shift").value);
    var worker = parseFloat(document.getElementById("p_worker").value);
    var directLabourUnit = parseFloat(document.getElementById("p_direct_labour_unit").value);
    document.getElementById("p_direct_labour").value = shift * worker * directLabourUnit;
    var directLabour = parseFloat(document.getElementById("p_direct_labour").value);
    document.getElementById("p_total_labour_cost").value = directLabour;
    var totalLabourCost = parseFloat(document.getElementById("p_total_labour_cost").value);

    var machineQty = parseFloat(document.getElementById("p_machine_qty").value);
    var area = parseFloat(document.getElementById("p_area").value);
    var specialMaterial = parseFloat(document.getElementById("p_special_material").value);
    var plantMaintenanceUnit = parseFloat(document.getElementById("p_plant_maintenance_unit").value);
    document.getElementById("p_plant_maintenance").value = plantMaintenanceUnit * area;
    var plantMaintenance = parseFloat(document.getElementById("p_plant_maintenance").value);


    var machineCostMonthPercentageUnit = parseFloat(document.getElementById("p_machine_cost_month_percentage_unit").value);
    var p_total_machine_cost = parseFloat(document.getElementById("p_total_machine_cost").value);
    var machineUsageDay = parseFloat(document.getElementById("p_machine_usage_day").value);
    document.getElementById("p_machine_cost_month2").value = (p_total_machine_cost / machineUsageDay).toFixed(0);
    var machineCostMonth2 = parseFloat(document.getElementById("p_machine_cost_month2").value);

    document.getElementById("p_machine_cost_month").value =
        (machineCostMonth2 * machineCostMonthPercentageUnit / 100).toFixed(2);
    var machineCostMonth = parseFloat(document.getElementById("p_machine_cost_month").value);
    

    var consumptionkwh = parseFloat(document.getElementById("p_consumption_kwh").value);
    var consumptionUnit = parseFloat(document.getElementById("p_consumption_unit").value);
    var C = (actualWorkingTimeMonth / 60);
    document.getElementById("p_consumption_sgd").value =
        (consumptionkwh * consumptionUnit * C).toFixed(0);
    var consumptionsgd = parseFloat(document.getElementById("p_consumption_sgd").value);
    var consumptionRate = parseFloat(document.getElementById("p_consumption_rate").value);
    document.getElementById("p_utility_electric").value =
        (consumptionsgd * consumptionRate/100).toFixed(0);
    var utilityElectric = parseFloat(document.getElementById("p_utility_electric").value);

    document.getElementById("p_machine_utility_cost").value =
        ((specialMaterial + plantMaintenance + machineCostMonth + machineCostMonth2 + utilityElectric) * machineQty).toFixed(0);
    var machineUtilityCost = parseFloat(document.getElementById("p_machine_utility_cost").value);

    document.getElementById("p_labour_electric_cost").value =
        (totalLabourCost + machineUtilityCost).toFixed(0);
    var labourElectricCost = parseFloat(document.getElementById("p_labour_electric_cost").value);

    if (actualWorkingTimeMonth>0)
    document.getElementById("p_charge").value =
        (labourElectricCost / actualWorkingTimeMonth).toFixed(2);
    var charge = parseFloat(document.getElementById("p_charge").value);


    var cycleTimeUnit = parseFloat(document.getElementById("p_cycle_time_unit").value);
    var time = parseFloat(document.getElementById("p_time").value);
    var capacity = parseFloat(document.getElementById("p_capacity").value);
    if (capacity>0)
    document.getElementById("p_cycle_time").value =
        (time / capacity).toFixed(2);
    var cycleTime = parseFloat(document.getElementById("p_cycle_time").value);

    if (capacity>0)
    document.getElementById("p_time_g").value =
        (time / capacity).toFixed(2);
    var p_time_g = parseFloat(document.getElementById("p_time_g").value);

    var efficiency = parseFloat(document.getElementById("p_efficiency").value);

    if (efficiency > 0 && cycleTime>0)
    document.getElementById("p_production_capacity").value =
        (actualWorkingTimeMonth * 60 / cycleTime * efficiency/100).toFixed(0);
    var productionCapacity = parseFloat(document.getElementById("p_production_capacity").value);

    if (productionCapacity>0)
    document.getElementById("p_production_cycle_time").value =
        (actualWorkingTimeMonth / productionCapacity).toFixed(4);
    var productionCycleTime = parseFloat(document.getElementById("p_production_cycle_time").value);

    var p_special_input = parseFloat(document.getElementById("p_special_input").value);

    if (productionCycleTime>0 && charge >0)
    document.getElementById("p_direct_process_cost").value =
        (charge * productionCycleTime).toFixed(4);
    var directProcessCost = parseFloat(document.getElementById("p_direct_process_cost").value);


    var table = document.getElementById("table_process_list");
    var total_process_cost = 0;
    for (var k = 1; k < table.rows.length; k++) {
        total_process_cost += parseFloat(table.rows[k].cells[2].innerText);
    }
    document.getElementById("process_cost").value = total_process_cost.toFixed(4);
}












