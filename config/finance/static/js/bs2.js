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

    
    for (let i = 3; i <= last_month_number-2; i++) {
      const liabilitiesAndNetValue = liabilitiesAndNetCells[i].textContent.trim();
      const totalAssetsValue = totalAssetsCells[i].textContent.trim();
      

      if (liabilitiesAndNetValue !== totalAssetsValue) {
        liabilitiesAndNetCells[i].style.backgroundColor = "yellow";
        totalAssetsCells[i].style.backgroundColor = "yellow";
   
        match = false;
       
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
   
      return;
    } else {
      template  = `Total Assets and Total Liabilities and Net Assets are not Balanced`;
    }


    modalText.innerHTML  = template
   

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

  // function toggleRow(rowId) {
  //   for (var i = 0; i <=100; i++) {
  //     var row = document.getElementById("row" + rowId + "-" + i);
  //     if (row) {
  //       row.style.display = row.style.display === "none" ? "table-row" : "none";
  //     }
  //   }
    
  //   var rowElement = document.getElementById("row" + rowId);
  //   if (rowElement) {
  //     rowElement.innerHTML = rowElement.innerHTML === "-" ? "+" : "-";
  //   }
  // }

  function toggleRow(rowId) {
    for (var i = 0; i <= 100; i++) {
      var row = document.getElementById("row" + rowId + "-" + i);
      if (row) {
        row.style.display = row.style.display === "none" ? "table-row" : "none";
      }
    }
    
    var iconElement = document.getElementById("row" + rowId + "-icon");
    if (iconElement) {
        iconElement.className = iconElement.className === "fa-solid fa-chevron-down" ? "fa-solid fa-chevron-up" : "fa-solid fa-chevron-down";
    }
  }
  