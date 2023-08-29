$(document).ready(function() {

  document.getElementById('export-pdf-button').addEventListener('click', function() {
    const element = document.getElementById('table-container');
    $('#spinner-modal').modal('show');

    let opt = {
      margin: 1,
      filename: 'cashflow.pdf',
      image: { type: 'png', quality: 0.98 },
      html2canvas: { scale: 2 },
      //[w, h]
      // jsPDF: { unit: 'mm', format: [297, 420], orientation: 'landscape' },
      jsPDF: { unit: 'mm', format: 'a1', orientation: 'portrait' },
    };
    html2pdf().from(element).set(opt).outputPdf().then(function(pdf) {
      // Hide the spinner modal
      $('#spinner-modal').modal('hide');
    }).save();
  });
});
