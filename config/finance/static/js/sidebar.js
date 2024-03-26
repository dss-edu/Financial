document.addEventListener("DOMContentLoaded", function() {
    const school = $("#hidden-school").val();
    const year = $("#hidden-year").val();
    const currentPath = window.location.pathname;
    // /charter-first/{{ school }}
    console.log(year)
    function highlightActiveLink(linkId) {
        $(linkId).addClass("activebg"); // Add a class for styling
    }

    const parts = currentPath.split('/');
    const month = parts[parts.length - 1];
  

    $("#first-link").on("click", function(event) {
        event.preventDefault();

        if (currentPath.includes("monthly")) {
          
            window.location.href = "/charter-first-monthly/" + school + "/" + month;

    }
        else{
            if (year) {
                window.location.href = "/charter-first/" + school + "/" + year;
            } else {
                window.location.href = "/charter-first/" + school;
            }
        }
  
    });

    $("#pl-link").on("click", function(event) {
        event.preventDefault();

        if (currentPath.includes("monthly")) {
          
            window.location.href = "/profit-loss-monthly/" + school + "/" + month;

    }else{
        if (year) {
            window.location.href = "/profit-loss/" + school + "/" + year;
        } else {
            window.location.href = "/profit-loss/" + school;
        }
    }

    });

    $("#bs-link").on("click", function(event) {
        event.preventDefault();

        if (currentPath.includes("monthly")) {
          
            window.location.href = "/balance-sheet-monthly/" + school + "/" + month;

    }else{
        if (year) {
            window.location.href = "/balance-sheet/" + school + "/" + year;
        } else {
            window.location.href = "/balance-sheet/" + school;
        }

    }

    });

    $("#cs-link").on("click", function(event) {
        event.preventDefault();

        if (currentPath.includes("monthly")) {
          
                window.location.href = "/cashflow-statement-monthly/" + school + "/" + month;
  
        }
        else{
            if (year) {
                window.location.href = "/cashflow-statement/" + school + "/" + year;
            } else {
                window.location.href = "/cashflow-statement/" + school;
            }
        }

    });

    $("#gl-link").on("click", function(event) {
        event.preventDefault();

        // $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/general-ledger/" + school;
        } else {
            window.location.href = "/general-ledger/" + school;
        }
    });

    $("#expend-link").on("click", function(event) {
        event.preventDefault();

        // $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/ytd-expend/" + school + "/" + year;
        } else {
            window.location.href = "/ytd-expend/" + school;
        }
    });


    if (currentPath === "/charter-first/" + school || currentPath === "/charter-first/" + school + "/" + year) {
        highlightActiveLink("#first-link");
        highlightActiveLink("#first-link2");
    
    } else if (currentPath === "/profit-loss/" + school || currentPath === "/profit-loss/" + school + "/" + year) {
        highlightActiveLink("#pl-link");
        highlightActiveLink("#pl-link2");
    } else if (currentPath === "/balance-sheet/" + school || currentPath === "/balance-sheet/" + school + "/" + year) {
        highlightActiveLink("#bs-link");
        highlightActiveLink("#bs-link2");
    } else if (currentPath === "/cashflow-statement/" + school || currentPath === "/cashflow-statement/" + school + "/" + year) {
        highlightActiveLink("#cs-link");
        highlightActiveLink("#cs-link2");
    } else if (currentPath === "/general-ledger/" + school || currentPath === "/general-ledger/" + school + "/" + year) {
        highlightActiveLink("#gl-link");
        highlightActiveLink("#gl-link2");
    }
    else if (currentPath === "/ytd-expend/" + school || currentPath === "/ytd-expend/" + school + "/" + year) {
        highlightActiveLink("#expend-link");
        highlightActiveLink("#expend-link2");
    }
    else if (currentPath === "/dashboard/" + school || currentPath === "/dashboard/" + school + "/" + year) {
        highlightActiveLink("#dashboard-link");
        highlightActiveLink("#dashboard-link2");
    }

});
