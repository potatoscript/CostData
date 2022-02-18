/*******
 * you need to switch over the following setting for localhost and company server doing deploring your project
 * 1. _url in site.js
 * 2. Client.BaseAddress in Helper.cs
 *******/
var _url = "/";          //if your app upload outside Default Web site - for my pc
//var _url = "/costdata/";  //if your app upload under Default Web site - for company


var tooling_no = 0;

jQuery(document).ready(function () {
    window.onresize = setWindow;
    setWindow();

    if (document.getElementById("CostId").value != 0) {
        document.getElementById("data_status").innerText = "EDITING";
        document.getElementById("Excel").style.display = "block";
    } else {
        document.getElementById("Excel").style.display = "none";
    }
        

    jQuery("#rubber_material_list").change(function () {
        var dat = String(document.getElementById("rubber_material_list").value).split(","); 
        document.getElementById("rubber_material_name").value = dat.slice(0, 1);
        document.getElementById("rubber_price_kg").value = parseFloat(dat.slice(1, 2));
        document.getElementById("rubber_mixing_process_cost").value = dat.slice(2,3);
        document.getElementById("rubber_weight_g").value = dat.slice(3,4);
        document.getElementById("rubber_yield_rate").value = dat.slice(4, 5);

        calculate_cost();

    })

    jQuery("#rubber_material_list2").change(function () {
        var dat = String(document.getElementById("rubber_material_list2").value).split(",");
        document.getElementById("rubber_material_name2").value = dat.slice(0, 1);
        document.getElementById("rubber_price_kg2").value = parseFloat(dat.slice(1, 2));
        document.getElementById("rubber_mixing_process_cost2").value = dat.slice(2, 3);
        document.getElementById("rubber_weight_g2").value = dat.slice(3, 4);
        document.getElementById("rubber_yield_rate2").value = dat.slice(4, 5);

        calculate_cost();

    })


    var obj = document.getElementById("rubber_material_list");
    obj.value = document.getElementById("rubber_material_name").value + "," +
        document.getElementById("rubber_price_kg").value + "," +
        document.getElementById("rubber_mixing_process_cost").value + "," +
        document.getElementById("rubber_weight_g").value + "," +
        document.getElementById("rubber_yield_rate").value;

    var obj2 = document.getElementById("rubber_material_list2");
    obj2.value = document.getElementById("rubber_material_name2").value + "," +
        document.getElementById("rubber_price_kg2").value + "," +
        document.getElementById("rubber_mixing_process_cost2").value + "," +
        document.getElementById("rubber_weight_g2").value + "," +
        document.getElementById("rubber_yield_rate2").value;
    /*
    for (var i = 0; i < obj.length; i++) {
        if (obj[i].text == "@ViewBag.rubber_material_name2") {
            obj[i].selected = true;
        }
    }
    */


    

    for (var i = 1; i < 16; i++) {
        jQuery("#tooling_list_data_" + i).change(function () {
            var dat = String(this.value).split(",");
            var n = String(this.id).split("_");
            var no = n.slice(3, 4);
            document.getElementById("tooling_list_description_"+no).value = dat.slice(0, 1);
            document.getElementById("tooling_list_source_"+no).value = dat.slice(1, 2);
            document.getElementById("tooling_list_qty_"+no).value = dat.slice(2, 3);
            document.getElementById("tooling_list_unit_"+no).value = dat.slice(3, 4);
            document.getElementById("tooling_list_price_"+no).value = dat.slice(4, 5);

            calculate_cost();
        })

        //load the tooling list value on read
        document.getElementById("tooling_list_data_" + i).value = document.getElementById("tooling_list_description_" + i).value + "," +
            document.getElementById("tooling_list_source_" + i).value + "," +
            document.getElementById("tooling_list_qty_" + i).value + "," +
            document.getElementById("tooling_list_unit_" + i).value + "," +
            document.getElementById("tooling_list_price_" + i).value;

    }
    
    

    jQuery("#checked_date,#issue_date,#approved_by,#expired_by")
        .css({ "cursor": "pointer" })
        .mouseover(function () {
            this.style.background = "lightyellow";
        })
        .mouseleave(function () {
            this.style.background = "white";
        })
        .datepicker({
            dateFormat: "dd-mm-yy",
            changeMonth: true,
            changeYear: true,
            showAnim: "show",
            beforeShow: function () {

            }
        })



    for (var s = document.getElementsByClassName("directInput"), t = 0; t < s.length; t++) {

        s[t].onclick = function () { this.select() };
        s[t].onkeyup = function () {
            document.getElementById("Excel").style.display = "none";
            jQuery("#table_parts_code").hide();
            calculate_cost();
        };
    }




    jQuery("#search_code").click(function () {
        document.getElementById("search_code").select();
        jQuery("#table_parts_code")
            .show()
            .css({
                "position": "absolute",
                "top": event.clientY + 20 + "px",
                "left": event.clientX - 40 + "px"

            })
        jQuery('#table_parts_code td')
            .click(function (event) {//for multiple select on process

                var RI = jQuery(this).parent().parent().children().index(this.parentNode);

                var table = document.getElementById("table_parts_code");
                document.getElementById("search_code").value =
                    table.rows[RI].cells[0].innerText;
                document.getElementById("CostId").value =
                    table.rows[RI].cells[2].innerText;

                window.location.href = _url + "Home/Read?id=" + table.rows[RI].cells[2].innerText;


                jQuery("#table_parts_code").hide();

            })
        jQuery('#table_parts_code th')
            .click(function (event) {//for multiple select on process
                jQuery("#table_parts_code").hide();

            })
    })

    calculate_cost();


    


});


function calculate_cost() {

    var qty = parseInt(document.getElementById("qty_month").value);
    var day = parseInt(document.getElementById("working_day").value);
    if (day > 0 && qty > 0)
        document.getElementById("production_qty_day").value = (qty / day).toFixed(0);


    var jpy = parseFloat(document.getElementById("exchange_rate_jpy").value);
    var n = 0;
    for (var i = 1; i < 16; i++) {
        var qty = parseFloat(document.getElementById("tooling_list_qty_" + i).value);
        var price = parseFloat(document.getElementById("tooling_list_price_" + i).value);
        document.getElementById("tooling_list_amount_sgd_" + i).value = (qty * price).toFixed(0);
        document.getElementById("tooling_list_amount_jpy_" + i).value = (qty * price * jpy).toFixed(0);
        n += (qty * price);
    }
    document.getElementById("tooling_list_total_amount_sgd").value = n.toFixed(0);
    var table = document.getElementById("table_summary");
    table.rows[1].cells[1].innerText = document.getElementById("tooling_list_total_amount_sgd").value;


    //rubber
    var rubber_mixing_process_cost = parseFloat(document.getElementById("rubber_mixing_process_cost").value);
    var rubber_weight_g = parseFloat(document.getElementById("rubber_weight_g").value);
    var rubber_yield_rate = parseFloat(document.getElementById("rubber_yield_rate").value);
    var rubber_price_kg = parseFloat(document.getElementById("rubber_price_kg").value);
    document.getElementById("rubber_weight_kg").value = (rubber_weight_g / 1000).toFixed(5);

    if (rubber_yield_rate > 0)
    document.getElementById("rubber_weight_kg_yieldrate").value = ((rubber_weight_g / 1000) / (rubber_yield_rate)).toFixed(5);

    var rubber_weight_kg_yieldrate = parseFloat(document.getElementById("rubber_weight_kg_yieldrate").value);
    document.getElementById("rubber_cost_sgd").value = ((rubber_price_kg + rubber_mixing_process_cost) * rubber_weight_kg_yieldrate).toFixed(4);
    var rubber_cost_sgd = parseFloat(document.getElementById("rubber_cost_sgd").value);

    //rubber2
    var rubber_mixing_process_cost2 = parseFloat(document.getElementById("rubber_mixing_process_cost2").value);
    var rubber_weight_g2 = parseFloat(document.getElementById("rubber_weight_g2").value);
    var rubber_yield_rate2 = parseFloat(document.getElementById("rubber_yield_rate2").value);
    var rubber_price_kg2 = parseFloat(document.getElementById("rubber_price_kg2").value);
    document.getElementById("rubber_weight_kg2").value = (rubber_weight_g2 / 1000).toFixed(5);
    if (rubber_yield_rate2 > 0)
    document.getElementById("rubber_weight_kg_yieldrate2").value = ((rubber_weight_g2 / 1000) / (rubber_yield_rate2)).toFixed(5);
    var rubber_weight_kg_yieldrate2 = parseFloat(document.getElementById("rubber_weight_kg_yieldrate2").value);
    document.getElementById("rubber_cost_sgd2").value = ((rubber_price_kg2 + rubber_mixing_process_cost2) * rubber_weight_kg_yieldrate2).toFixed(4);
    var rubber_cost_sgd2 = parseFloat(document.getElementById("rubber_cost_sgd2").value);

    //material
    var material_inhouse_value_1 = parseFloat(document.getElementById("material_inhouse_value_1").value);
    var material_inhouse_value_1b = parseFloat(document.getElementById("material_inhouse_value_1b").value);
    var material_inhouse_value_2 = parseFloat(document.getElementById("material_inhouse_value_2").value);
    var material_inhouse_value_2b = parseFloat(document.getElementById("material_inhouse_value_2b").value);
    var material_inhouse_value_3 = parseFloat(document.getElementById("material_inhouse_value_3").value);
    var material_inhouse_value_3b = parseFloat(document.getElementById("material_inhouse_value_3b").value);

    document.getElementById("material_inhouse_cost_sgd_1").value = ((material_inhouse_value_1 / 1000) * material_inhouse_value_1b).toFixed(4);
    document.getElementById("material_inhouse_cost_sgd_2").value = ((material_inhouse_value_2 / 1000) * material_inhouse_value_2b).toFixed(4);
    document.getElementById("material_inhouse_cost_sgd_3").value = ((material_inhouse_value_3 / 1000) * material_inhouse_value_3b).toFixed(4);

    var material_outside_value_1 = parseFloat(document.getElementById("material_outside_value_1").value);
    var material_outside_value_1b = parseFloat(document.getElementById("material_outside_value_1b").value);
    var material_outside_value_2 = parseFloat(document.getElementById("material_outside_value_2").value);
    var material_outside_value_2b = parseFloat(document.getElementById("material_outside_value_2b").value);
    var material_outside_value_3 = parseFloat(document.getElementById("material_outside_value_3").value);
    var material_outside_value_3b = parseFloat(document.getElementById("material_outside_value_3b").value);

    document.getElementById("material_outside_cost_sgd_1").value = ((material_outside_value_1 / 1000) * material_outside_value_1b).toFixed(4);
    document.getElementById("material_outside_cost_sgd_2").value = ((material_outside_value_2 / 1000) * material_outside_value_2b).toFixed(4);
    document.getElementById("material_outside_cost_sgd_3").value = ((material_outside_value_3 / 1000) * material_outside_value_3b).toFixed(4);

    var direct_material_cost =
        parseFloat(document.getElementById("rubber_cost_sgd").value) +
        parseFloat(document.getElementById("rubber_cost_sgd2").value) +
        parseFloat(document.getElementById("material_inhouse_cost_sgd_1").value) +
        parseFloat(document.getElementById("material_inhouse_cost_sgd_2").value) +
        parseFloat(document.getElementById("material_inhouse_cost_sgd_3").value) +
        parseFloat(document.getElementById("material_outside_cost_sgd_1").value) +
        parseFloat(document.getElementById("material_outside_cost_sgd_2").value) +
        parseFloat(document.getElementById("material_outside_cost_sgd_3").value);

    document.getElementById("direct_material_cost").value = direct_material_cost.toFixed(4);

    var sub_material_percentage = parseFloat(document.getElementById("sub_material_percentage").value);
    document.getElementById("sub_material_cost").value = (direct_material_cost * sub_material_percentage / 100).toFixed(4);


    var direct_process_cost = parseFloat(document.getElementById("direct_process_cost").value);


    document.getElementById("total_direct_cost").value = (direct_material_cost +
        direct_material_cost * sub_material_percentage / 100 +
        direct_process_cost).toFixed(4);


    //Defective
    var defective_percentage = parseFloat(document.getElementById("defective_percentage").value);
    document.getElementById("defective_cost").value = (parseFloat(document.getElementById("total_direct_cost").value) * defective_percentage / 100).toFixed(4);


    //Indirect cost
    var indirect_percentage = parseFloat(document.getElementById("indirect_percentage").value);
    document.getElementById("indirect_cost").value = ((
        parseFloat(document.getElementById("total_direct_cost").value) +
        parseFloat(document.getElementById("defective_cost").value)
    ) * indirect_percentage / 100).toFixed(4);


    //packing
    var packing_material_percentage = parseFloat(document.getElementById("packing_material_percentage").value);
    var special_package_cost = parseFloat(document.getElementById("special_package_cost").value);
    document.getElementById("packing_material_cost").value = (((
        parseFloat(document.getElementById("direct_material_cost").value)
    ) * packing_material_percentage / 100) + special_package_cost).toFixed(4);




    //Administration
    var administration_percentage = parseFloat(document.getElementById("administration_percentage").value);
    document.getElementById("administration_cost").value = ((
        parseFloat(document.getElementById("total_direct_cost").value) +
        parseFloat(document.getElementById("defective_cost").value) +
        parseFloat(document.getElementById("indirect_cost").value)
    ) * administration_percentage / 100).toFixed(4);


    //Tooling Maintenance was plant_retail_percentage
    var plant_retail_percentage = parseFloat(document.getElementById("plant_retail_percentage").value);
    document.getElementById("plant_retail_cost").value = ((
        parseFloat(document.getElementById("total_direct_cost").value) +
        parseFloat(document.getElementById("defective_cost").value) +
        parseFloat(document.getElementById("indirect_cost").value) +
        parseFloat(document.getElementById("packing_material_cost").value) +
        parseFloat(document.getElementById("administration_cost").value)
    ) * plant_retail_percentage / 100).toFixed(4);


    //Tool Mold and Jig
    var moldjig_percentage = parseFloat(document.getElementById("moldjig_percentage").value);
    if (parseFloat(document.getElementById("qty_month").value)>0)
    document.getElementById("moldjig_cost").value =
        (moldjig_percentage / 24 / parseFloat(document.getElementById("qty_month").value)).toFixed(4);


    //Tool Die
    var die_percentage = parseFloat(document.getElementById("die_percentage").value);
    if (parseFloat(document.getElementById("qty_month").value)>0)
    document.getElementById("die_cost").value =
        (die_percentage / 24 / parseFloat(document.getElementById("qty_month").value)).toFixed(4);


    //Incuded toling cost
    document.getElementById("net_included_tooling_cost").value =
        (
            parseFloat(document.getElementById("total_direct_cost").value) +
            parseFloat(document.getElementById("defective_cost").value) +
            parseFloat(document.getElementById("indirect_cost").value) +
            parseFloat(document.getElementById("packing_material_cost").value) +
            parseFloat(document.getElementById("administration_cost").value) +
            parseFloat(document.getElementById("plant_retail_cost").value) +
            parseFloat(document.getElementById("moldjig_cost").value) +
            parseFloat(document.getElementById("die_cost").value)
        ).toFixed(4);


    //Exclude toling cost
    document.getElementById("net_exclude_tooling_cost").value =
        (
            parseFloat(document.getElementById("net_included_tooling_cost").value) -
            parseFloat(document.getElementById("moldjig_cost").value) -
            parseFloat(document.getElementById("die_cost").value)
        ).toFixed(4);


    //calculate the target price(SGD)
    document.getElementById("target_price_sgd").value = (parseFloat(document.getElementById("net_exclude_tooling_cost").value) * 1.2).toFixed(4);
    var target_price_sgd = parseFloat(document.getElementById("target_price_sgd").value);
    document.getElementById("target_price_usd").value = (parseFloat(document.getElementById("exchange_rate_usd").value) * target_price_sgd).toFixed(4);
    document.getElementById("target_price_eud").value = (parseFloat(document.getElementById("exchange_rate_eud").value) * target_price_sgd).toFixed(4);

    var target_price_wr_sgd = parseFloat(document.getElementById("target_price_wr_sgd").value);
    document.getElementById("target_price_wr_usd").value = (parseFloat(document.getElementById("exchange_rate_usd").value) * target_price_wr_sgd).toFixed(4);
    document.getElementById("target_price_wr_eud").value = (parseFloat(document.getElementById("exchange_rate_eud").value) * target_price_wr_sgd).toFixed(4);


    var table = document.getElementById("table_summary");
    table.rows[0].cells[1].innerText = document.getElementById("target_price_sgd").value;



    document.getElementById("defectives").value = document.getElementById("defective_cost").value;
    document.getElementById("defectives_p").value = ((parseFloat(document.getElementById("defective_cost").value) / target_price_sgd) * 100).toFixed(1);
    var admin_engin_qc = parseFloat(document.getElementById("indirect_cost").value) +
        parseFloat(document.getElementById("administration_cost").value) +
        parseFloat(document.getElementById("plant_retail_cost").value);
    document.getElementById("admin_engin_qc").value = admin_engin_qc;
    document.getElementById("admin_engin_qc_p").value = ((admin_engin_qc / target_price_sgd) * 100).toFixed(1);

    var tooling_cost = parseFloat(document.getElementById("moldjig_cost").value) +
        parseFloat(document.getElementById("die_cost").value);
    document.getElementById("tooling_cost").value = tooling_cost;
    document.getElementById("tooling_cost_p").value = ((tooling_cost / target_price_sgd) * 100).toFixed(1);

    if (target_price_sgd > 0) {
        document.getElementById("rubber_percentage_target_price").value = (rubber_cost_sgd / target_price_sgd * 100).toFixed(0);
        document.getElementById("rubber_percentage_target_price2").value = (rubber_cost_sgd2 / target_price_sgd * 100).toFixed(0);
        document.getElementById("material_inhouse_percentage_target_price_1").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_1").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("material_inhouse_percentage_target_price_2").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_2").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("material_inhouse_percentage_target_price_3").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_3").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("material_outside_percentage_target_price_1").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_1").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("material_outside_percentage_target_price_2").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_2").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("material_outside_percentage_target_price_3").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_3").value) / target_price_sgd) * 100).toFixed(1);
        document.getElementById("direct_material_cost_percentage").value = ((direct_material_cost / target_price_sgd) * 100).toFixed(2);
        
        document.getElementById("sub_material_cost_percentage").value = (((direct_material_cost * sub_material_percentage / 100) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("direct_process_cost_percentage").value = ((direct_process_cost / target_price_sgd) * 100).toFixed(2);
        document.getElementById("total_direct_cost_percentage").value =
            ((parseFloat(document.getElementById("total_direct_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("defective_cost_percentage").value = ((parseFloat(document.getElementById("defective_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("indirect_cost_percentage").value = ((parseFloat(document.getElementById("indirect_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("packing_material_cost_percentage").value = ((parseFloat(document.getElementById("packing_material_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("administration_cost_percentage").value = ((parseFloat(document.getElementById("administration_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("plant_retail_cost_percentage").value = ((parseFloat(document.getElementById("plant_retail_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("moldjig_cost_percentage").value = ((parseFloat(document.getElementById("moldjig_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("die_cost_percentage").value = ((parseFloat(document.getElementById("die_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("net_included_tooling_cost_percentage").value = ((parseFloat(document.getElementById("net_included_tooling_cost").value) / target_price_sgd) * 100).toFixed(2);
        document.getElementById("net_exclude_tooling_cost_percentage").value = ((parseFloat(document.getElementById("net_exclude_tooling_cost").value) / target_price_sgd) * 100).toFixed(2);

        document.getElementById("direct_raw_material").value = direct_material_cost.toFixed(4);
        document.getElementById("direct_raw_material_p").value = ((direct_material_cost / target_price_sgd) * 100).toFixed(2);
        document.getElementById("sub_material").value =
            parseFloat(document.getElementById("sub_material_cost").value) +
            parseFloat(document.getElementById("packing_material_cost").value);

        document.getElementById("sub_material_p").value =
            ((parseFloat(document.getElementById("sub_material").value) / target_price_sgd) * 100).toFixed(2);
    }


    document.getElementById("raw_material_cost_sub_total").value =
        (parseFloat(document.getElementById("direct_raw_material").value) +
            parseFloat(document.getElementById("sub_material").value)).toFixed(4);

    document.getElementById("raw_material_cost_sub_total_p").value =
        (parseFloat(document.getElementById("direct_raw_material_p").value) +
            parseFloat(document.getElementById("sub_material_p").value)).toFixed(1);


    document.getElementById("other_fixed_cost_sub_total").value =
        (parseFloat(document.getElementById("defectives").value) +
            parseFloat(document.getElementById("admin_engin_qc").value)+
            parseFloat(document.getElementById("tooling_cost").value)+
            parseFloat(document.getElementById("process_margin_adjust").value)
        ).toFixed(4);

    document.getElementById("other_fixed_cost_sub_total_p").value =
        (parseFloat(document.getElementById("defectives_p").value) +
            parseFloat(document.getElementById("admin_engin_qc_p").value) +
            parseFloat(document.getElementById("tooling_cost_p").value) +
            parseFloat(document.getElementById("process_margin_adjust_p").value)
        ).toFixed(1);


    document.getElementById("grand_total_cost").value =
        (parseFloat(document.getElementById("sub_material").value) +
        parseFloat(document.getElementById("process_cost_sub_total").value) +
            parseFloat(document.getElementById("other_fixed_cost_sub_total").value)
        ).toFixed(4);

    document.getElementById("grand_total_cost_p").value =
        (parseFloat(document.getElementById("sub_material_p").value) +
            parseFloat(document.getElementById("process_cost_sub_total_p").value) +
            parseFloat(document.getElementById("other_fixed_cost_sub_total_p").value)
        ).toFixed(1);

    if (parseFloat(document.getElementById("cycle_time").value) > 0) {
        var prod_capa = parseFloat(document.getElementById("actual_working_time").value) * 60 / parseFloat(document.getElementById("cycle_time").value);
        var prod_capa_e = prod_capa * (parseFloat(document.getElementById("efficiency").value) / 100);
        document.getElementById("production_capacity").value = (prod_capa_e).toFixed(0);
    }
    
    var daily_qty_days = parseFloat(document.getElementById("production_capacity").value) / 20;
    if (daily_qty_days > 0) {
        document.getElementById("daily_qty_days").value = daily_qty_days.toFixed(0);

        var qty_month = parseFloat(document.getElementById("qty_month").value);
        document.getElementById("daily_qty_days_p").value = (qty_month / daily_qty_days).toFixed(2);



        var dailyamount = 1.15 * parseFloat(document.getElementById("net_exclude_tooling_cost").value) * daily_qty_days;
        document.getElementById("daily_amount").value = (dailyamount).toFixed(0);
    }
    

}


function jQueryAjaxPost(form, page) {
    try {
        jQuery.ajax({
            type: "POST",
            url: form.action,
            data: new FormData(form),
            contentType: false,
            processData: false,
            success: function (res) {
                if (res.isValid) {
                    var p = _url + page;
                    if (page == "ProcessMaster/Index") {
                        //document.getElementById("RefreshProcessMaster").click();
                        refreshProcessMaster();
                    } 
                    if (page == "Home/Index") {
                      
                        RefreshData();
                    }
                    
                }
            },
            error: function (err) {
                console.log(err)
            }
        })
    } catch (e) {
        console.log(e);
    }
    // to prevent default form submit event
    return false;
}

function setWindow() {

    document.getElementById('div_body').style.width = (window.innerWidth - 80) + 'px';
    document.getElementById('div_body').style.height = (window.innerHeight - 105) + 'px';

    document.getElementById('table_summary').style.width = (window.innerWidth - 110) + 'px';
}

// keep for reference not working
function ExcelData() {
    jQuery.ajax({
        type: "GET",
        url: _url + 'Home/ExcelData',
        data: jQuery.param({
            customer: "abc",
            confirm: true
        }),
        success: function (res) {

        }
    });


}

function delete_data(id) {
    
    var confirm_delete = confirm("Delete Data?");
    if (confirm_delete) {
        jQuery.ajax({
            type: "POST",
            url: _url + 'Home/Delete',
            data: jQuery.param({
                Id: id,
                confirm: true
            }),
            success: function (res) {
                window.location.href = _url + "Home/Index";
            }
        });
    }

}

function MasterData() {

}


function RefreshData() {
    
    var url = _url + "Home";
    window.location.href = url;
}


function SubmitData(page) {
    document.getElementById("SubmitButton").value=page;
    document.getElementById("SubmitButton").click();
   
}


function showPopup(url, title) {
    /*
    jQuery.ajax({
        type: "GET",
        url: url,
        success: function (res) {
            jQuery("#form-modal .modal-body").html(res);
            jQuery("#form-modal .modal-title").html(title);
            jQuery("#form-modal").modal('show');
        }
    })
    */
    var total_rubber_weight = parseFloat(document.getElementById("rubber_weight_g").value) +
        parseFloat(document.getElementById("rubber_weight_g2").value);

    jQuery.ajax({
        type: "GET",
        url: _url + 'CostProcess/Index',
        data: jQuery.param({
            p_doc_no: document.getElementById("doc_no").value,
            p_od: document.getElementById("item_od").value,
            p_process_type: document.getElementById("product_type").value,
            p_rubber_weight: total_rubber_weight
        }),
        success: function (res) {
            jQuery("#form-modal .modal-body").html(res);
            jQuery("#form-modal .modal-title").html("Process Cost Data");
            jQuery("#form-modal").modal('show');

        }
    })


}


function showPopupMaster(url, title) {

    jQuery.ajax({
        type: "GET",
        url: url,
        success: function (res) {
            jQuery("#form-modal .modal-body").html(res);
            jQuery("#form-modal .modal-title").html(title);
            jQuery("#form-modal").modal('show');
        }
    })


}








