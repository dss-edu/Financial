$(document).ready(function() {
    // Initialize DataTables
    let table = new DataTable("#dataTable", {
        // scrollX: true,
        ordering: false,
        // dom: "lrtip",
        dom: "Bfrtlip",
        buttons: [
            {
                text: "add",
                action: function(e, dt, node, config) {
                    $("#add-modal").modal("show");
                },
            },
            {
                text: "delete",
                extend: "selected",
                action: function(e, dt, node, config) {
                    $("#page-load-spinner").modal("show");
                    deleteAdjustments();
                },
            },
            {
                text: "update",
                extend: "selectedSingle",
                action: function(e, dt, node, config) {
                    updateAdjustments();
                },
            },
        ],
        select: true,
    });

    $("#dataTable thead tr:nth-child(2) th").each(function(index) {
        var column = table.column(index);
        var select = $('<select><option value=""></option></select>')
            .appendTo($(this).empty())
            .on("change", function() {
                var val = $.fn.dataTable.util.escapeRegex($(this).val());
                column.search(val ? val : "", true, false).draw(); // Updated search method
            });

        // Populate the select dropdown with unique values from the column
        column
            .data()
            .unique()
            .each(function(d) {
                select.append('<option value="' + d + '">' + d + "</option>");
            });
    });
    function updateAdjustments() {
        $("#update-modal").modal("show");
        const csrfToken = $("input[name='csrfmiddlewaretoken']").val();
        const school = $("input[name='school']").val();
        const rows = table.rows({ selected: true }).data().toArray();
        const headers = Array.from($("thead tr:first-child th")).map(function(th) {
            return $(th).text();
        });

        // Map the data array to an object using headers as keys
        let rowObject = {};
        for (let i = 0; i < rows[0].length; i++) {
            rowObject[headers[i]] = rows[0][i];
        }
        rowObject["school"] = school;

        console.log(rowObject);
        $("#update-modal select[name='func']").val(rowObject["func"]);
        console.log($("#update-modal input[name='func']"));
        $("#update-modal select[name='fund']").val(rowObject["fund"]);
        $("#update-modal select[name='obj']").val(rowObject["obj"]);
        $("#update-modal select[name='org']").val(rowObject["org"]);
        $("#update-modal select[name='fscl_yr']").val(rowObject["fscl_yr"]);
        $("#update-modal select[name='pgm']").val(rowObject["pgm"]);
        $("#update-modal select[name='edSpan']").val(rowObject["edSpan"]);
        $("#update-modal select[name='projDtl']").val(rowObject["projDtl"]);
        $("#update-modal select[name='acctDescr']").val(rowObject["AcctDescr"]);
        $("#update-modal select[name='number']").val(rowObject["Number"]);
        const date = new Date(rowObject["Date"]);
        $("#update-modal input[name='date']").val(date);
        $("#update-modal select[name='acctPer']").val(rowObject["AcctPer"]);
        $("#update-modal input[name='acct-real']").val(rowObject["Real"]);
        $("#update-modal input[name='expend']").val(rowObject["Expend"]);
        $("#update-modal input[name='bal']").val(rowObject["Bal"]);
        $("#update-modal input[name='work_descr']").val(rowObject["WorkDescr"]);
        $("#update-modal input[name='type']").val(rowObject["Type"]);
    }

    function deleteAdjustments() {
        const rowCount = table.rows({ selected: true }).count();
        const csrfToken = $("input[name='csrfmiddlewaretoken']").val();
        const school = $("input[name='school']").val();
        const rows = table.rows({ selected: true }).data().toArray();
        const headers = Array.from($("thead tr:first-child th")).map(function(th) {
            return $(th).text();
        });

        // Map the data array to an object using headers as keys
        let rowObject = [];
        for (let i = 0; i < rows.length; i++) {
            let rowData = {};
            for (let j = 0; j < rows[i].length; j++) {
                rowData[headers[j]] = rows[i][j];
            }
            rowData["school"] = school;
            rowObject.push(rowData);
        }
        data = {
            csrfmiddlewaretoken: csrfToken,
            adjustments: JSON.stringify(rowObject),
            school: school,
        };
        $("#page-load-spinner").modal("show");
        $.ajax({
            url: "/delete-adjustments/",
            async: false,
            type: "POST",
            // dataType: "json", // Set the expected data type of the response
            // contentType: "application/json", // Set the content type of the request
            data: data,
            success: function(response) {
                $("#page-load-spinner").modal("hide");
                alert("Successfully deleted " + rowCount + " row(s)");
                window.location.href = "/manual-adjustments/" + school;
            },
            error: function(error) {
                // Handle errors here
                $("#page-load-spinner").modal("hide");
                alert("Delete unsuccessful");
            },
            complete: function() {
                $("#page-load-spinner").modal("hide");
            },
        });
    }
});
