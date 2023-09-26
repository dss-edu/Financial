$(document).ready(function() {
    const schoolSelect = $("#school-select").select2({
        placeholder: "Select a school",
        allowClear: true,
    });

    schoolSelect.on("select2:select", function(e) {
        const selectedOption = e.params.data.id;
        const anchorYear = $("#hidden-year").val();

        $("#page-load-spinner").modal("show");

        if (selectedOption) {
            window.location.href = "/dashboard/" + selectedOption;
        }
        // if (selectedOption && anchorYear) {
        //     window.location.href = "/dashboard/" + selectedOption + "/" + anchorYear;
        // } else {
        //     window.location.href = "/dashboard/" + selectedOption;
        // }
    });

    $("#year-select").on("change", function() {
        const year = $(this).val();
        const school = $("#hidden-school").val();
        if (year === "default") {
            window.location.href = "/dashboard/" + school;
        } else {
            $("#page-load-spinner").modal("show");
            window.location.href = "/dashboard/" + school + "/" + year;
        }
    });
});
