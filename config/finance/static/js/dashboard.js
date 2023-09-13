document.addEventListener("DOMContentLoaded", function() {
    const circles = {
        "green-circle": "G",
        "yellow-circle": "Y",
        "red-circle": "R",
    };
    StatusChecker();

    function StatusChecker() {
        test("days-coh", cohCriteria);
        test("student-staff", studentStaffCriteria);
        test("debt-service", debtServiceCriteria);
    }

    function test(rowId, criteriaFunc) {
        const row = document.getElementById(rowId);
        const td = row.getElementsByTagName("td");
        let itemVal = td[1].innerText.toLowerCase();
        let statusTD = row.querySelector(".status");

        let status = criteriaFunc(itemVal);

        statusTD.classList.toggle(status);
        statusTD.classList.toggle("legend");
        statusTD.innerText = circles[status];
    }

    function cohCriteria(value) {
        // from 93 days to 93(int)
        const coh = parseInt(value.trim().split(" "));
        let status = "";

        if (coh > 60) {
            status = "green-circle";
        } else if (coh < 20) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        return status;
    }

    function studentStaffCriteria(value) {
        let status = "";

        if (value === "no") {
            status = "green-circle";
        } else {
            status = "red-circle";
        }
        return status;
    }
    function debtServiceCriteria(value) {
        const debt = parseFloat(removeStringsInNum(value));
        let status = "";
        if (debt >= 1.1) {
            status = "green-circle";
        } else if (debt < 1.0) {
            status = "red-circle";
        } else {
            status = "yellow-circle";
        }
        return status;
    }
    function removeStringsInNum(str) {
        return str.replace(/[^0-9.\-]/g, "");
    }
});
