var _url = "/";          //if your app upload outside Default Web site - for my pc
//var _url = "/costnag/";  //if your app upload under Default Web site - for company

jQuery(document).ready(function () {
    window.onresize = setWindow;
    setWindow();

    if (document.getElementById("CostId").value != 0)
        document.getElementById("data_status").innerText = "EDITING";

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
            jQuery("#table_parts_code").hide();
        };

    }

    for (var s = document.getElementsByClassName("production_qty"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {
            var qty = parseInt(document.getElementById("qty_month").value);
            var day = parseInt(document.getElementById("working_day").value);
            if (day > 0 && qty > 0)
                document.getElementById("production_qty_day").value = (qty / day).toFixed(0);
        };

    }


    for (var s = document.getElementsByClassName("material"), t = 0; t < s.length; t++) {
        s[t].onkeyup = function () {
            calculate_cost();
        };
    }

    for (var s = document.getElementsByClassName("toolinglist"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {
            var jpy = parseFloat(document.getElementById("exchange_rate").value);
            var n = 0;
            for (var i = 1; i < 16; i++) {
                var qty = parseFloat(document.getElementById("tooling_list_qty_"+i).value);
                var price = parseFloat(document.getElementById("tooling_list_price_"+i).value);
                document.getElementById("tooling_list_amount_usd_"+i).value = (qty * price).toFixed(0);
                document.getElementById("tooling_list_amount_jpy_"+i).value = (qty * price * jpy).toFixed(0);
                n += (qty * price);
            }
            document.getElementById("tooling_list_total_amount_usd").value = n.toFixed(0);
            var table = document.getElementById("table_summary");
            table.rows[1].cells[1].innerText = document.getElementById("tooling_list_total_amount_usd").value;

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


    

});


function calculate_cost() {
    //rubber
    var rubber_database_price_new = parseFloat(document.getElementById("rubber_database_price_new").value);
    var rubber_mixing_process_cost = parseFloat(document.getElementById("rubber_mixing_process_cost").value);
    var rubber_weight_g = parseFloat(document.getElementById("rubber_weight_g").value);
    var rubber_yield_rate = parseFloat(document.getElementById("rubber_yield_rate").value);

    document.getElementById("rubber_price_kg").value = (rubber_database_price_new * 105 / 100).toFixed(2);
    var rubber_price_kg = parseFloat(document.getElementById("rubber_price_kg").value);

    document.getElementById("rubber_weight_kg").value = (rubber_weight_g / 1000).toFixed(5);

    document.getElementById("rubber_weight_kg_yieldrate").value = ((rubber_weight_g / 1000) / (rubber_yield_rate / 100)).toFixed(5);
    var rubber_weight_kg_yieldrate = parseFloat(document.getElementById("rubber_weight_kg_yieldrate").value);

    document.getElementById("rubber_cost_sgd").value = ((rubber_price_kg + rubber_mixing_process_cost) * rubber_weight_kg_yieldrate).toFixed(4);
    var rubber_cost_sgd = parseFloat(document.getElementById("rubber_cost_sgd").value);

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
    document.getElementById("moldjig_cost").value =
        (moldjig_percentage / 24 / parseFloat(document.getElementById("qty_month").value)).toFixed(4);


    //Tool Die
    var die_percentage = parseFloat(document.getElementById("die_percentage").value);
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
    document.getElementById("target_price_bht").value = (parseFloat(document.getElementById("net_exclude_tooling_cost").value) * 1.2).toFixed(4);
    var table = document.getElementById("table_summary");
    table.rows[0].cells[1].innerText = document.getElementById("target_price_bht").value;

    var target_price_sgd = parseFloat(document.getElementById("target_price_bht").value);

    document.getElementById("rubber_percentage_target_price").value = (rubber_cost_sgd / target_price_sgd * 100).toFixed(0);
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

}


function jQueryAjaxPost(form,page) {
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
                    if (page == "Process/Index") {
                        document.getElementById("direct_process_cost").click();

                    } else {
                        window.location.href = p;
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

    document.getElementById('div_body').style.width = (window.innerWidth - 60) + 'px';
    document.getElementById('div_body').style.height = (window.innerHeight - 105) + 'px';

    document.getElementById('table_summary').style.width = (window.innerWidth - 90) + 'px';
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

function RefreshData() {
    window.location.href = _url + "Home/Index";
}

function SubmitData() {
    document.getElementById("SubmitButton").click();
   
}


function showPopup(url, title) {
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









