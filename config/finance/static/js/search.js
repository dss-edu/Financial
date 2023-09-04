$(document).ready(function(){
    const schoolSelect = $('#school-select').select2({
        placeholder: 'Select a school',
        allowClear: true,
    });


    schoolSelect.on('select2:select', function(e){
        var selectedOption = e.params.data.id;
        if (selectedOption) {
            $('#page-load-spinner').modal('show');
            window.location.href = "/dashboard/"+selectedOption;
        }
    });
});
