document.addEventListener("DOMContentLoaded", function() {
    const addAdjustmentForm = document.getElementById("add-adjustment-form");
    const saveAdjustmentBtn = document.getElementById("save-adjustment");
    saveAdjustmentBtn.addEventListener("click", async function(event) {
        event.preventDefault();
        const school = $("input[name='school']").val();

        // You might want to do some client-side validation here before submitting the form.

        $("#add-modal").modal("hide");
        $("#page-load-spinner").modal("show");

        try {
            // Submit the form asynchronously
            const response = await fetch("/add-adjustments/", {
                method: "POST", // or "PUT" or "DELETE" depending on your server-side implementation
                body: new FormData(addAdjustmentForm), // Assuming you want to send form data
            });

            // Check the response status code (you may need to adjust this condition)
            if (response.ok) {
                $("page-load-spinner").modal("hide");
                window.location.href = "/manual-adjustments/" + school;
                // const responseData = await response.json(); // Parse the JSON response if applicable

                // Depending on the condition, you can hide the modal
                // if (responseData.status === "success") {
                //     $("#add-modal").modal("hide");
                // } else {
                //     $("#add-modal").modal("hide");
                //     alert("Failed to add a new manual adjustment");
                //     // Handle the case where the response indicates failure
                // }
            } else {
                $("page-load-spinner").modal("hide");
                throw new Error("Failed to add manual adjustment.");
            }
        } catch (error) {
            // Handle any errors that occurred during the fetch request
            $("page-load-spinner").modal("hide");
            alert("An error occurred while processing your request.");
        }
    });

    $("#workdescr-input").select2({
        tags: true, // Allow user to enter new tags
        tokenSeparators: [",", " "], // Define how tags are separated
        placeholder: "", // Placeholder text
        maximumSelectionLength: 1, // Allow only one selection
    });
});
