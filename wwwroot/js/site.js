//var _url = "/";          //if your app upload outside Default Web site - for my pc
var _url = "/costnag/";  //if your app upload under Default Web site - for company

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