document.addEventListener("DOMContentLoaded", function() {
    const circles = {
        "green-circle": "On Track",
        "yellow-circle": "Of Concern",
        "red-circle": "At Risk",
    };
    StatusChecker();
    commaSeparator();



    function totalAssets(){

    }

    function StatusChecker() {
        test("indicator", criteriaPassFail);
        test("net-assets", projectionCriteria);
        test("estimated-actual-ada", projectionCriteria);
        test("budget-vs-revenue", projectionCriteria);
        test("reporting-peims", projectionCriteria);
        test("annual-audit", projectionCriteria);
        test("post-financial-info", projectionCriteria);
        test("estimated-first-rating", ratingCriteria);
        test("ratio-student-teacher", measureCriteria);
        test("approved-geo-boundaries", measureCriteria);
        test("days-coh", cohCriteria);
        test("current-assets", currAssetsCriteria);
    
        test("num11", num11Criteria);
        test("num12", num12Criteria);
        test("num13", num13Criteria);
        
    }

    function test(rowId, criteriaFunc) {
        const row = document.getElementById(rowId);
        
        const td = row.getElementsByTagName("td");
        let itemVal = td[1].textContent.toLowerCase();
       
        statusTD = td[3];
        pointsTD = td[2];
        
        
        if (rowId === "estimated-first-rating") {
            statusTD = td[2];
            pointsTD = "";
        }

        criteriaFunc(itemVal, statusTD, pointsTD);
    }
    function num13Criteria(value, statusTD, pointsTD) {
   
        
      
        value = parseFloat(value.replace('%', ''));
        console.log(value)
        const p = statusTD.querySelector("p");
        let status = "";
        if (value < 95) {
            points = 5;
            status = "green-circle";
         
        }
        else if (value >= 95) {
            points = 0;
            status = "red-circle";
        }
     

        pointsTD.textContent = points
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function num12Criteria(value, statusTD, pointsTD) {
   
        value = parseFloat(value)
        const p = statusTD.querySelector("p");
        let status = "";
        if (value >= 1.20) {
            points = 10;
            status = "green-circle";
         
        } 
        else if (value < 1.20  && value >= 1.15) {
            points = 8;
            status = "green-circle";
        }
        else if (value < 1.15 && value >= 1.10) {
            points = 6;
            status = "green-circle";
        }
        else if (value < 1.10 && value >= 1.05) {
            points = 4;
            status = "red-circle";
        }
        else if (value < 1.05 && value >= 1.00) {
            points = 2;
            status = "red-circle";
        }
        else if (value < 1.00 ) {
            points = 0;
            status = "red-circle";
            
        }

        pointsTD.textContent = points
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function num11Criteria(value, statusTD, pointsTD) {
   
        value = parseFloat(value)
        const p = statusTD.querySelector("p");
        let status = "";
        if (value <= .60) {
            points = 10;
            status = "green-circle";
         
        } 
        else if (value > .60 && value <= .70) {
            points = 8;
            status = "green-circle";
        }
        else if (value > .70 && value <= .80) {
            points = 6;
            status = "green-circle";
        }
        else if (value > .80 && value <= .90) {
            points = 4;
            status = "red-circle";
        }
        else if (value > .90 && value <= 1.00) {
            points = 2;
            status = "red-circle";
        }
        else if (value > 1.00 ) {
            points = 0;
            status = "red-circle";
            
        }

        pointsTD.textContent = points
        p.classList.toggle(status);
        p.textContent = circles[status];
    }


    function criteriaPassFail(value, statusTD,pointsTD) {
        
        const p = statusTD.querySelector("p");
        let status = "";
        if (value === "pass") {
            status = "green-circle";
        } else if (value === "fail") {
            status = "red-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }

    function projectionCriteria(value, statusTD,pointsTD) {
        const p = statusTD.querySelector("p");
        let status = "";
        if (value === "projected") {
            status = "green-circle";
        } else {
            status = "red-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function ratingCriteria(value, statusTD,pointsTD) {
        const rating = parseInt(value);
        const p = statusTD.querySelector("p");
        const grade = document.getElementById("rating-grade");
        let status = "";
        if (rating < 69) {
            grade.innerText = "F - Fail";
            status = "red-circle";
        } else if (rating < 80) {
            grade.innerText = "C - Meets Standard";
            status = "yellow-circle";
        } else if (rating < 90) {
            grade.innerText = "B - Above Standard";
            status = "green-circle";
        } else {
            grade.innerText = "A - Superior";
            status = "green-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }

    function measureCriteria(value, statusTD,pointsTD) {
        const rating = parseInt(value);
        const p = statusTD.querySelector("p");
        let status = "";
        if (value === "not measured by dss") {
            status = "green-circle";
        } else {
            status = "red-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function cohCriteria(value, statusTD,pointsTD) {
        const coh = parseInt(value);
        const p = statusTD.querySelector("p");
        let status = "";

        if (coh > 60) {
            status = "green-circle";
        } else if (coh < 20) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function currAssetsCriteria(value, statusTD, pointsTD) {
        const currAssets = parseFloat(value);
        const p = statusTD.querySelector("p");
        let status = "";

        if (currAssets > 2) {
            status = "green-circle";
        } else if (currAssets < 1) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function netEarningsCriteria(value, statusTD,pointsTD) {
        const currAssets = parseFloat(value);
        const p = statusTD.querySelector("p");
        let status = "";

        if (currAssets > 2) {
            status = "green-circle";
        } else if (currAssets < 1) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }
    function assetsCriteria(value, statusTD,pointsTD) {
        const currAssets = parseFloat(value);
        const p = statusTD.querySelector("p");
        let status = "";

        if (currAssets > 2) {
            status = "green-circle";
        } else if (currAssets < 1) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        p.classList.toggle(status);
        p.textContent = circles[status];
    }

    function commaSeparator() {
        // ytd_income
        const netIncomeYTD = $("#net-income-ytd").text().trim();
        // net_earnings
        const netEarnings = $("#net-earnings-td").text().trim();
        // initialize comma separated variables
        let commaYTD, commaEarnings;
        commaYTD = numFormatter(netIncomeYTD);
        commaEarnings = numFormatter(netEarnings);
        $("#net-income-ytd").text(commaYTD);
        $("#net-earnings-td").text(commaEarnings);

        function numFormatter(str) {
            const matchNumber = str.match(/\d+/);
            let commaStr;
            if (matchNumber) {
                const number = parseInt(matchNumber[0], 10); // Convert the extracted string to a number
                const formattedNumber = number.toLocaleString(); // Add commas
                if (str.includes("(")) {
                    commaStr = `$(${formattedNumber})`; // Format the result
                } else {
                    commaStr = `$${formattedNumber}`; // Format the result
                }
            }
            return commaStr;
        }
    }
});
