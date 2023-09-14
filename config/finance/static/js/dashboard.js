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
    // document
    //     .getElementById("financial-form")
    //     .addEventListener("submit", function(event) {
    //         event.preventDefault();
    //         const csrfToken = $("#financial-form")
    //             .find('input[name="csrfmiddlewaretoken"]')
    //             .val();
    //         const school = this.dataset.school;
    //         const noteElements = document.querySelectorAll("td[name='notes']");
    //         function getNotes(noteElements) {
    //             const notes = [];
    //             noteElements.forEach((note) => {
    //                 notes.push(note.textContent);
    //             });
    //             return notes;
    //         }
    //         const notes = getNotes(noteElements);
    //         const formData = {
    //             csrfmiddlewaretoken: csrfToken,
    //             notesList: notes,
    //         };
    //
    //         fetch("/dashboard/notes/" + school, {
    //             method: "POST",
    //             body: formData,
    //         })
    //             .then((response) => response.json())
    //             .then((data) => {
    //                 console.log(data); // Handle the response from the server
    //             })
    //             .catch((error) => {
    //                 console.error(error); // Handle errors
    //             });
    //     });
    // edit saveBtn
    saveBtn.addEventListener("click", function(event) {
        const csrfToken = $("input[name='csrfmiddlewaretoken']").val();
        const school = this.dataset.school;
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
            url: "/dashboard/notes/" + school,
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

    function getCookie(name) {
        let cookie = document.cookie.match("(^|;) ?" + name + "=([^;]*)(;|$)");
        return cookie ? cookie[2] : null;
    }
});
