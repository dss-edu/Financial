$(document).ready(function() {
    const schoolSelect = $("#school-select").select2({
        placeholder: "Select a school",
        allowClear: true,
        dropdownCssClass: "side-search",
    });

    schoolSelect.on("select2:select", function(e) {
        const selectedOption = e.params.data.id;
        const anchorYear = $("#hidden-year").val();

      
        $("#page-load-spinner").modal("show");

        if (selectedOption) {
            if(currentPath.includes("/dashboard/")){
                window.location.href = "/dashboard/" + selectedOption;
            }else if(currentPath.includes("/profit-loss/")){
                window.location.href = "/profit-loss/" + selectedOption;
            }else if (currentPath.includes("/balance-sheet/")) {
                window.location.href = "/balance-sheet/" + selectedOption;
            } else if (currentPath.includes("/cashflow-statement/")) {
                window.location.href = "/cashflow-statement/" + selectedOption;
            }else if (currentPath.includes("/charter-first/")) {
                window.location.href = "/charter-first/" + selectedOption;
            }
            else if (currentPath.includes("/general-ledger/")) {
                window.location.href = "/charter-first/" + selectedOption;
            }else{
                window.location.href = "/dashboard/" + selectedOption;
            }
            
        }
        // if (selectedOption && anchorYear) {
        //     window.location.href = "/dashboard/" + selectedOption + "/" + anchorYear;
        // } else {
        //     window.location.href = "/dashboard/" + selectedOption;
        // }
    });
    const currentPath = window.location.pathname;
 
    $("#year-select").on("change", function() {
        const year = $(this).val();
        const school = $("#hidden-school").val();
        if (year === "default") {
            if (currentPath.includes("/dashboard/" + school)) {
                window.location.href = "/dashboard/" + school;
            } else if (currentPath.includes("/profit-loss/" + school)) {
                window.location.href = "/profit-loss/" + school;
            }else if (currentPath.includes("/balance-sheet/" + school)) {
                window.location.href = "/balance-sheet/" + school;
            }
            else if (currentPath.includes("/cashflow-statement/" + school)) {
                window.location.href = "/cashflow-statement/" + school;
            }else if (currentPath.includes("/charter-first/" + school)) {
                window.location.href = "/charter-first/" + school;
            }else{
                window.location.href = "/general-ledger/" + school;
            }
           
        } else {
            $("#page-load-spinner").modal("show");
            //window.location.href = "/dashboard/" + school + "/" + year ;
            if (currentPath.includes("/dashboard/" + school)) {
                window.location.href = "/dashboard/" + school + "/" + year;
            } else if (currentPath.includes("/profit-loss/" + school)) {
                window.location.href = "/profit-loss/" + school + "/" + year;
            }else if (currentPath.includes("/balance-sheet/" + school)) {
                window.location.href = "/balance-sheet/" + school + "/" + year;
            }
            else if (currentPath.includes("/cashflow-statement/" + school)) {
                window.location.href = "/cashflow-statement/" + school + "/" + year;
            }else if (currentPath.includes("/charter-first/" + school)) {
                window.location.href = "/charter-first/" + school + "/" + year;
            }else{
                window.location.href = "/general-ledger//" + school + "/" + year;
            }
        }
    });
});
