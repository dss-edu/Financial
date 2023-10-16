document.addEventListener("DOMContentLoaded", function() {
    const circles = {
        "green-circle": "On Track",
        "yellow-circle": "Of Concern",
        "red-circle": "At Risk",
    };
    StatusChecker();
    commaSeparator();

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
    }

    function test(rowId, criteriaFunc) {
        const row = document.getElementById(rowId);
        const td = row.getElementsByTagName("td");
        let itemVal = td[1].innerText.toLowerCase();
        statusTD = td[3];
        if (rowId === "estimated-first-rating") {
            statusTD = td[2];
        }

        criteriaFunc(itemVal, statusTD);
    }
    function criteriaPassFail(value, statusTD) {
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

    function projectionCriteria(value, statusTD) {
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
    function ratingCriteria(value, statusTD) {
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

    function measureCriteria(value, statusTD) {
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
    function cohCriteria(value, statusTD) {
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
    function currAssetsCriteria(value, statusTD) {
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
    function netEarningsCriteria(value, statusTD) {
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
