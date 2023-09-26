document.addEventListener("DOMContentLoaded", function() {
    const school = $("#hidden-school").val();
    const year = $("#hidden-year").val();
    // /charter-first/{{ school }}
    $("#first-link").on("click", function(event) {
        event.preventDefault();

        $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/charter-first/" + school + "/" + year;
        } else {
            window.location.href = "/charter-first/" + school;
        }
    });

    $("#pl-link").on("click", function(event) {
        event.preventDefault();

        $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/profit-loss/" + school + "/" + year;
        } else {
            window.location.href = "/profit-loss/" + school;
        }
    });

    $("#bs-link").on("click", function(event) {
        event.preventDefault();

        $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/balance-sheet/" + school + "/" + year;
        } else {
            window.location.href = "/balance-sheet/" + school;
        }
    });

    $("#cs-link").on("click", function(event) {
        event.preventDefault();

        $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/cashflow-statement/" + school + "/" + year;
        } else {
            window.location.href = "/cashflow-statement/" + school;
        }
    });

    $("#gl-link").on("click", function(event) {
        event.preventDefault();

        $("#page-load-spinner").modal("show");
        if (year) {
            window.location.href = "/general-ledger/" + school + "/" + year;
        } else {
            window.location.href = "/general-ledger/" + school;
        }
    });
});
