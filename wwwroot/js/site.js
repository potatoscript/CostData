var _url = "/";          //if your app upload outside Default Web site - for my pc
//var _url = "/costnag/";  //if your app upload under Default Web site - for company

jQuery(document).ready(function () {
    jQuery("#checked_date,#issue_date")
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


    for (var s = document.getElementsByClassName("rubber"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {
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
            var target_price_bht = parseFloat(document.getElementById("target_price_bht").value);

            document.getElementById("rubber_percentage_target_price").value = (rubber_cost_sgd / target_price_bht * 100).toFixed(0);
        };
    }

    for (var s = document.getElementsByClassName("material"), t = 0; t < s.length; t++) {

        s[t].onkeyup = function () {

            var target_price_sgd = parseFloat(document.getElementById("target_price_bht").value);

            var material_inhouse_value_1 = parseFloat(document.getElementById("material_inhouse_value_1").value);
            var material_inhouse_value_1b = parseFloat(document.getElementById("material_inhouse_value_1b").value);
            var material_inhouse_value_2 = parseFloat(document.getElementById("material_inhouse_value_2").value);
            var material_inhouse_value_2b = parseFloat(document.getElementById("material_inhouse_value_2b").value);
            var material_inhouse_value_3 = parseFloat(document.getElementById("material_inhouse_value_3").value);
            var material_inhouse_value_3b = parseFloat(document.getElementById("material_inhouse_value_3b").value);

            document.getElementById("material_inhouse_cost_sgd_1").value = ((material_inhouse_value_1 / 1000) * material_inhouse_value_1b).toFixed(4);
            document.getElementById("material_inhouse_cost_sgd_2").value = ((material_inhouse_value_2 / 1000) * material_inhouse_value_2b).toFixed(4);
            document.getElementById("material_inhouse_cost_sgd_3").value = ((material_inhouse_value_3 / 1000) * material_inhouse_value_3b).toFixed(4);

            document.getElementById("material_inhouse_percentage_target_price_1").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_1").value) / target_price_sgd) * 100).toFixed(1);
            document.getElementById("material_inhouse_percentage_target_price_2").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_2").value) / target_price_sgd) * 100).toFixed(1);
            document.getElementById("material_inhouse_percentage_target_price_3").value = ((parseFloat(document.getElementById("material_inhouse_cost_sgd_3").value) / target_price_sgd) * 100).toFixed(1);


            var material_outside_value_1 = parseFloat(document.getElementById("material_outside_value_1").value);
            var material_outside_value_1b = parseFloat(document.getElementById("material_outside_value_1b").value);
            var material_outside_value_2 = parseFloat(document.getElementById("material_outside_value_2").value);
            var material_outside_value_2b = parseFloat(document.getElementById("material_outside_value_2b").value);
            var material_outside_value_3 = parseFloat(document.getElementById("material_outsdie_value_3").value);
            var material_outside_value_3b = parseFloat(document.getElementById("material_outside_value_3b").value);

            document.getElementById("material_outside_cost_sgd_1").value = ((material_outside_value_1 / 1000) * material_outside_value_1b).toFixed(4);
            document.getElementById("material_outside_cost_sgd_2").value = ((material_outside_value_2 / 1000) * material_outside_value_2b).toFixed(4);
            document.getElementById("material_outside_cost_sgd_3").value = ((material_outside_value_3 / 1000) * material_outside_value_3b).toFixed(4);

            document.getElementById("material_outside_percentage_target_price_1").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_1").value) / target_price_sgd) * 100).toFixed(1);
            document.getElementById("material_outside_percentage_target_price_2").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_2").value) / target_price_sgd) * 100).toFixed(1);
            document.getElementById("material_outside_percentage_target_price_3").value = ((parseFloat(document.getElementById("material_outside_cost_sgd_3").value) / target_price_sgd) * 100).toFixed(1);


        };

    }


    jQuery("#parts_code").click(function () {
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
                document.getElementById("parts_code").value =
                    table.rows[RI].cells[0].innerText;
                document.getElementById("CostId").value =
                    table.rows[RI].cells[1].innerText;

                window.location.href = _url + "Home/Read?id=" + table.rows[RI].cells[1].innerText;

                jQuery("#table_parts_code").hide();

            })
        jQuery('#table_parts_code th')
            .click(function (event) {//for multiple select on process
                jQuery("#table_parts_code").hide();

            })
    });

});

function jQueryAjaxPost(form) {
    try {
        jQuery.ajax({
            type: "POST",
            url: form.action,
            data: new FormData(form),
            contentType: false,
            processData: false,
            success: function (res) {
                if (res.isValid) {
                    window.location.href = _url + "Home/Index";
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