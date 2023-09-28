// Function to check if values in two rows match
function checkValuesMatch() {
    const liabilitiesAndNetRow = document.getElementById("liabilitiesandnet-total-row");
    const totalAssetsRow = document.getElementById("total-assets-total-row");
    
    if (!liabilitiesAndNetRow || !totalAssetsRow) {
      return; 
    }
    
    const liabilitiesAndNetCells = liabilitiesAndNetRow.cells;
    const totalAssetsCells = totalAssetsRow.cells;
    
    let match = true;
    
    for (let i = 3; i <= 17; i++) {
      const liabilitiesAndNetValue = liabilitiesAndNetCells[i].textContent.trim();
      const totalAssetsValue = totalAssetsCells[i].textContent.trim();
    
      if (liabilitiesAndNetValue !== totalAssetsValue) {
        match = false;
        break;
      }
    }
    
    return match;
  }

  function checkforMissingActivities(){
    const table = document.getElementById('settings-table');
    const tBody = document.getElementsByTagName('tbody')[0];
    // const tRows = tBody.getElementsByTagName('tr');
    const tRows = $('#settings-table .no-act')

    if (tRows.length > 0){
      return true;

    }
    return false;
  }
 
  function showModal(match) {
    const modal = document.getElementById("myModal2");
    const modalText = document.getElementById("modal-text");
    const missingActivites = checkforMissingActivities();
    
    if (!modal || !modalText) {
      return; 
    }

    let template = ``
    if (match) {
      template = `Total Assets and Total Liabilities and Net Assets are Balanced`;
    } else {
      template  = `Total Assets and Total Liabilities and Net Assets are not Balanced`;
    }

    if (missingActivites){
      template += `\n\nMissing tags for activities. Click <a id="settings-link" href="#">here</a> to set tags.`

    }

    modalText.innerHTML  = template

    if (missingActivites){
      const settingsLink = document.getElementById('settings-link')
      settingsLink.addEventListener('click', function(event){
        event.preventDefault()
        $('#myModal2').modal('hide')

        $('#settings-modal').modal('show')
      })
    }

    // if (match) {
    //   modalText.innerHTML  = "Total Assets and Total Liabilities and Net Assets are Balanced ";
    // } else {
    //   modalText.innerHTML  = "Total Assets and Total Liabilities and Net Assets are not Balanced";
    // }

    modal.style.display = "block";
  }


  const closeButton = document.querySelector(".close");
  if (closeButton) {
    closeButton.addEventListener("click", function() {
      const modal = document.getElementById("myModal2");
      modal.style.display = "none";
    });
  }


  window.addEventListener("load", function() {
    const match = checkValuesMatch();
    showModal(match);
  });

  function toggleRow(rowId) {
    for (var i = 0; i <=100; i++) {
      var row = document.getElementById("row" + rowId + "-" + i);
      if (row) {
        row.style.display = row.style.display === "none" ? "table-row" : "none";
      }
    }
    
    var rowElement = document.getElementById("row" + rowId);
    if (rowElement) {
      rowElement.innerHTML = rowElement.innerHTML === "-" ? "+" : "-";
    }
  }
