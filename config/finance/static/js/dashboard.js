document.addEventListener("DOMContentLoaded", function() {
    const circles = {
        "green-circle": "On Track",
        "yellow-circle": "Of Concern",
        "red-circle": "At Risk",
    };
    StatusChecker();
    commaSeparator();

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

    const editNotesBtn = document.getElementById("edit-notes");
    const tableButtonsDiv = document.getElementById("table-save");
    const saveBtn = tableButtonsDiv.querySelector(".save");
    const cancelBtn = tableButtonsDiv.querySelector(".cancel");

    editNotesBtn.addEventListener("click", function() {
        tableButtonsDiv.classList.toggle("d-none");
        tableButtonsDiv.classList.toggle("d-flex");
        editNotesBtn.disabled = true;
        changeEditableProp("true");
    });

    function resetEditNotes() {
        editNotesBtn.disabled = false;
        tableButtonsDiv.classList.toggle("d-none");
        tableButtonsDiv.classList.toggle("d-flex");
    }

    let defaults = [];
    const table = document.querySelector("#financial-update tbody");
    const rows = table.getElementsByTagName("tr");
    function changeEditableProp(state) {
        rows.forEach((row) => {
            const td = row.getElementsByTagName("td");
            const cellToRevert = td[3];
            const input = cellToRevert.textContent;
            cellToRevert.contentEditable = state;
            if (state === "true") {
                defaults.push(input);
            } else {
                cellToRevert.textContent = defaults.shift();
            }
        });
    }
    saveBtn.addEventListener("click", function(event) {
        const csrfToken = $("input[name='csrfmiddlewaretoken']").val();
        const school = this.dataset.school;
        const fy = this.dataset.fy;
        const anch_month = this.dataset.anch_month;
  
        var url;
        if(anch_month != ""){
            url ="/dashboard/notes/" + school + "/" + fy + "/" + anch_month;
        }else{
            url ="/dashboard/notes/" + school ;
        }
        const noteElements = document.querySelectorAll("td[name='notes']");
        function getNotes(noteElements) {
            const notes = [];
            noteElements.forEach((note) => {
                notes.push(note.textContent);
            });
            return notes;
        }
        const notes = getNotes(noteElements);
        data = {
            csrfmiddlewaretoken: csrfToken,
            notesList: notes,
        }; // Convert data to JSON format
        $("#spinner-modal").modal("show");
        $.ajax({
            url: url,
            async: false,
            type: "POST",
            // dataType: "json", // Set the expected data type of the response
            // contentType: "application/json", // Set the content type of the request
            data: data,
            success: function(response) {
                // Handle the successful response here
                console.log(response);
            },
            error: function(error) {
                // Handle errors here
                console.error(error);
            },
            complete: function() {
                $("#spinner-modal").modal("hide");
            },
        });

        defaults = notes;
        changeEditableProp("false"); // Define the URL of the endpoint
        resetEditNotes();
    });
    cancelBtn.addEventListener("click", function() {
        changeEditableProp("false");
        resetEditNotes();
    });

    function commaSeparator() {
        // ytd_income
        const netIncomeYTD = $("#ytd td:eq(1) p:eq(0)").text().trim();
        // net_earnings
        const netEarnings = $("#ytd td:eq(1) p:eq(1)").text().trim();
        // initialize comma separated variables
        let commaYTD, commaEarnings;
        commaYTD = numFormatter(netIncomeYTD);
        commaEarnings = numFormatter(netEarnings);
        $("#ytd td:eq(1) p:eq(0)").text(commaYTD);
        $("#ytd td:eq(1) p:eq(1)").text(commaEarnings);

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
