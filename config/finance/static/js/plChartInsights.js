document.addEventListener("DOMContentLoaded", function() {
    test("rating-value", "rating-grade", ratingCriteria);
    test("dtcRatio", "dtcCircle", dtcCriteria);
    test("liabRatio", "liabCircle", liabCriteria);
    test("adminRatio", "adminCircle", adminCriteria);
    test("dscr", "dscrCircle", dscrCriteria);
    test("currentRatio", "currentCircle", currentCriteria);
    test("daysCOH", "cohCircle", cohCriteria);

    function test(elementId, statusId, criteriaFunc) {
        const el = document.getElementById(elementId);
        const val = el.textContent;
        const status = document.getElementById(statusId);
        // const td = row.getElementsByTagName("td");
        // let itemVal = td[1].innerText.toLowerCase();
        // statusTD = td[3];
        // if (rowId === "estimated-first-rating") {
        //     statusTD = td[2];
        // }

        criteriaFunc(val, status);
    }

    function dtcCriteria(value, statusCircle) {
        const percent = parseFloat(value.trim());
        if (percent < 95) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }

    function liabCriteria(value, statusCircle) {
        const liabVal = parseFloat(value.trim());
        if (liabVal < 0.6) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }

    function adminCriteria(value, statusCircle) {
        const adminVal = parseFloat(value.trim().split("/")[0]);
        if (adminVal < 15.61) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }

    function dscrCriteria(value, statusCircle) {
        const dscrVal = parseFloat(value.trim());
        if (dscrVal > 1.2) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }

    function currentCriteria(value, statusCircle) {
        const val = parseFloat(value.trim());
        if (val > 2) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }

    function cohCriteria(value, statusCircle) {
        const val = parseInt(value.trim());
        if (val > 60) {
            statusCircle.classList.toggle("check-circle");
            statusCircle.textContent = "check_circle";
        } else {
            statusCircle.classList.toggle("error-circle");
            statusCircle.textContent = "error";
        }
    }
    function ratingCriteria(value, status) {
        const val = parseInt(value);
        if (val >= 90) {
            status.textContent = "A - Superior";
        } else if (val >= 80) {
            status.textContent = "B - Above Standard";
        } else if (val >= 70) {
            status.textContent = "C - Meets Standard";
        } else {
            status.textContent = "F - Fail";
        }
    }
});
