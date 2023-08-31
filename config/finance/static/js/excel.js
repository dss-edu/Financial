$(document).ready(function(){
      document.getElementById('export-excel-button').addEventListener('click', function() {
          generateExcel();
      });

      function generateExcel() {
          $('#spinner-modal').modal('show');
          fetch('/generate_excel/')
              .then(function(response) {
                  return response.blob();
              })
              .then(function(blob) {
                  const url = window.URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.style.display = 'none';
                  a.href = url;
                  a.download = 'GeneratedExcel.xlsx'; // Change the filename as needed
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
