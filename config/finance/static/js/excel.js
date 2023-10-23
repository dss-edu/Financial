$(document).ready(function(){
    document.getElementById('export-excel-button').addEventListener('click', function() {
        generateExcel(school);
        console.log(year)
    });

   
    
    function generateExcel(school) {
        $('#spinner-modal').modal('show');
        fetch('/generate_excel/' + school + '/' + year )
            .then(function(response) {
                
                return response.blob();
            })
            .then(function(blob) {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = 'GeneratedExcel.xlsx'; 
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
            })
            .catch(function(error) {
                console.error('Error generating Excel:', error);
            }).finally(function (){
                $('#spinner-modal').modal('hide');
            });
    }
});