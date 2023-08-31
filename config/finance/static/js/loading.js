$('#page-load-spinner').modal('show');  // Close the modal

window.addEventListener('load', function() {
  $('#page-load-spinner').modal('hide');  // Close the modal
});

$(document).ready(function() {
  const sidebarWrapper = document.querySelector('.sidebar');
  const linkEl = sidebarWrapper.querySelectorAll('li a');


  linkEl.forEach(link => {
    link.addEventListener('click', function() {
      $('#page-load-spinner').modal('show');  // Close the modal
    });
  });
  // $('#page-load-spinner').modal('show');  // Close the modal
});

