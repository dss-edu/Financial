document.addEventListener("DOMContentLoaded", function() {
    const addAdjustmentBtn = document.getElementById("add-adjustment");
    addAdjustmentBtn.addEventListener("click", function() {
        $("#add-modal").modal("show");
    });

    const addAdjustmentForm = document.getElementById("add-adjustment-form");
    const saveAdjustmentBtn = document.getElementById("save-adjustment");
    saveAdjustmentBtn.addEventListener("click", function(event) {
        event.preventDefault();
        addAdjustmentForm.submit();
        $("#add-modal").modal("hide");
    });
});
