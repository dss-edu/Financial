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
    const month = parts[parts.length - 1].toString().padStart(2, '0');
    let September 
    if (typeof sept !== 'undefined' && sept !== null) {
        September = sept;
    } else {
        September = "True";
    }



        $("#dashboard-link").on("click", function(event) {
            event.preventDefault();
    
            if (currentPath.includes("monthly")) {
                let  currentYearcharter = new Date().getFullYear().toString();
                let monthValue = parseInt(month, 10); // Convert to integer
       
                
                if (September == 'True'){
                    
                    if (monthValue > 9 ){
                        currentYearcharter = (parseInt(currentYearcharter, 10) - 1).toString();
                       
                    }
           
                }else{
                    if (monthValue > 7 ){
                        currentYearcharter = (parseInt(currentYearcharter, 10) - 1).toString();
                    }
                }
              
                if (monthValue < 10) {
                    monthValue = monthValue.toString(); // Convert back to string
                }
              window.location.href = "/dashboard-monthly/" + school + "/" + currentYearcharter + "/" + monthValue;
    
        }
            else{
                if (year) {
                    window.location.href = "/dashboard/" + school + "/" + year;
                } else {
                    window.location.href = "/dashboard/" + school;
                }
            }
      
        });
    $("#first-link").on("click", function(event) {
        event.preventDefault();

        if (currentPath.includes("monthly")) {
            let  currentYearcharter = new Date().getFullYear().toString();
            let monthValue = parseInt(month, 10); // Convert to integer
   
            
            if (September == 'True'){
                
                if (monthValue > 9 ){
                    currentYearcharter = (parseInt(currentYearcharter, 10) - 1).toString();
                   
                }
       
            }else{
                if (monthValue > 7 ){
                    currentYearcharter = (parseInt(currentYearcharter, 10) - 1).toString();
                }
            }
          
            if (monthValue < 10) {
                monthValue = monthValue.toString(); // Convert back to string
            }
          window.location.href = "/charter-first-monthly/" + school + "/" + currentYearcharter + "/" + monthValue;

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
