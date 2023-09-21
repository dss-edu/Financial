

  

  document.addEventListener("DOMContentLoaded", function () {
 
    var modal = document.getElementById("myModal");

      
      var modalTableBody = document.getElementById("modal-table-body");

   
      function populateModal(data) {
        
        modalTableBody.innerHTML = "";

        
        data.glbs_data.forEach(function (row) {
          var newRow = document.createElement("tr");
          newRow.innerHTML = `
            <td class="fund-td">${row.fund}</td>
            <td class="fund-td">${row.func}</td>
            <td>${row.obj}</td>
            <td>${row.org}</td>
            <td>${row.fscl_yr}</td>
            <td>${row.pgm}</td>
           
            <td>${row.projDtl}</td>
            <td>${row.AcctDescr}</td>
            <td>${row.Number}</td>
            <td style="white-space: nowrap;">${row.Date}</td>
            <td>${row.AcctPer}</td>
            
            <td>${row.Real}</td>
            
            
            <td>${row.Expend}</td>
            <td>${row.Bal}</td>
            <td style="white-space: nowrap;">${row.WorkDescr}</td>
            <td>${row.Type}</td>
            
          `;
          modalTableBody.appendChild(newRow);
        });

        var totalRow = document.createElement("tr");
        totalRow.innerHTML = `
          <td colspan="13" style="text-align: right;">Total Balance</td>
          <td>$${data.total_bal}</td>
          <td colspan="2"></td>
        `;
        modalTableBody.appendChild(totalRow);

      }

    function fetchDataAndPopulateModal(obj, yr) {
      fetch(`/viewgl_activitybs/${obj}/${yr}/`)
        .then(function (response) {
          
          return response.json();
        })
        .then(function (data) {
     
          if (data.status === "success") {
            
            populateModal(data.data , data.total_bal);
            
            modal.style.display = "block";
          } else {

          }
        })
        .catch(function (error) {
  
        });
    }

   
    var viewGLLinks = document.querySelectorAll(".viewgl_activitybs-link");
    viewGLLinks.forEach(function (link) {
      link.addEventListener("click", function (event) {
        event.preventDefault();
        
        var obj = link.dataset.obj;
        var yr = link.dataset.yr;
        fetchDataAndPopulateModal(obj , yr);
      });
    });

    
    var closeButton = modal.querySelector(".modal-footer button");
    closeButton.addEventListener("click", function () {
      modal.style.display = "none";
    });
  });


  var columnsCollapsed = false;
  function toggleColumns() {
    var toggleButton = document.getElementById("toggle-button");
    var collapsedColumns = document.getElementsByClassName("collapsed");
    
    for (var i = 0; i < collapsedColumns.length; i++) {
      if (columnsCollapsed) {
        collapsedColumns[i].style.display = "table-cell";
        toggleButton.innerHTML = "-";
      } else {
        collapsedColumns[i].style.display = "none";
        toggleButton.innerHTML = "+";
      }
    }
    
    // Toggle the state
    columnsCollapsed = !columnsCollapsed;
  }



    
  function validateBudgetInput(event) {
    const input = event.target;
    const value = input.value;

    
    const regex = /^[0-9.]*$/;

    
    if (!regex.test(value)) {
        
        input.value = value.replace(/[^0-9.]/g, '');
    }
}
  
   
    
function extractNumericValue(content) {
  const match = content.match(/\(([^)]+)\)/);
  if (match) {
    const numericValue = parseFloat(match[1].replace(/[$,]/g, '').trim());
    return isNaN(numericValue) ? 0 : -numericValue;
  } else {
    const numericValue = parseFloat(content.replace(/[$,]/g, '').trim());
    return isNaN(numericValue) ? 0 : numericValue;
  }
}

  /* --------------current assets total function ------------------*/
  function calculateCurrentAssets() {
    const rows = document.querySelectorAll('.current-assets-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`current-assets-total-${cellId}`);
      
      totalCells.push(totalCell);
      
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
      const columnTotal = parseInt(columnTotals[i]);
      totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';

    }
  }

 /* --------------capital assets total function ------------------*/
  function calculateCapitalAssets() {
    const rows = document.querySelectorAll('.capital-assets-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`capital-assets-total-${cellId}`);
      
      totalCells.push(totalCell);
      
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
      const columnTotal = parseInt(columnTotals[i]);
      totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';
    }
  }

  /* -------------- total  assets function ------------------*/
  function calculateTotalAssets() {
    const rows = document.querySelectorAll('.capital-assets-row,.current-assets-row');
    const totalCells = [];
  
    const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,16, 17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`total-assets-total-${cellId}`);

      totalCells.push(totalCell);
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
      const columnTotal = columnTotals[i]; // Don't parse the integer value here
      totalCell.textContent = columnTotal !== 0 ? '$' + columnTotal.toLocaleString() : '';
    }
  }

  /* --------------total current Liabilities function ------------------*/
  function calculateCurrentLiabilities() {
    const rows = document.querySelectorAll('.current-liabilities-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`current-liabilities-total-${cellId}`);
      
      totalCells.push(totalCell);
      
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
    const columnTotal = parseInt(columnTotals[i]);
    const formattedTotal = columnTotal < 0 ? `(${Math.abs(columnTotal).toLocaleString()})` : columnTotal.toLocaleString();
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
   }
  }


  /* --------------total  Liabilities function ------------------*/
  function calculateTotalLiabilities() {
    const rows = document.querySelectorAll('.total-liabilities-row,.current-liabilities-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`total-liabilities-total-${cellId}`);
      
      totalCells.push(totalCell);
      
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
    const columnTotal = parseInt(columnTotals[i]);
    const formattedTotal = columnTotal < 0 ? `(${Math.abs(columnTotal).toLocaleString()})` : columnTotal.toLocaleString();
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
  }
  }


  /* --------------total  Liabilities and nets assets function ------------------*/
  function calculateTotalLiabilitiesAndNetAssets() {
    const rows = document.querySelectorAll('.net-assets-row, .total-liabilities-row, .current-liabilities-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`liabilitiesandnet-total-${cellId}`);
      
      totalCells.push(totalCell); document.addEventListener("DOMContentLoaded", function () {
 
        var modal = document.getElementById("myModal");
    
          
          var modalTableBody = document.getElementById("modal-table-body");
    
       
          function populateModal(data) {
            
            modalTableBody.innerHTML = "";
    
            
            data.glbs_data.forEach(function (row) {
              var newRow = document.createElement("tr");
              newRow.innerHTML = `
                <td class="fund-td">${row.fund}</td>
                <td class="fund-td">${row.func}</td>
                <td>${row.obj}</td>
                <td>${row.org}</td>
                <td>${row.fscl_yr}</td>
                <td>${row.pgm}</td>
               
                <td>${row.projDtl}</td>
                <td>${row.AcctDescr}</td>
                <td>${row.Number}</td>
                <td style="white-space: nowrap;">${row.Date}</td>
                <td>${row.AcctPer}</td>
                
                <td>${row.Real}</td>
                
                
                <td>${row.Expend}</td>
                <td>${row.Bal}</td>
                <td style="white-space: nowrap;">${row.WorkDescr}</td>
                <td>${row.Type}</td>
                
              `;
              modalTableBody.appendChild(newRow);
            });
    
            var totalRow = document.createElement("tr");
            totalRow.innerHTML = `
              <td colspan="13" style="text-align: right;">Total Balance</td>
              <td>$${data.total_bal}</td>
              <td colspan="2"></td>
            `;
            modalTableBody.appendChild(totalRow);
    
          }
    
        function fetchDataAndPopulateModal(obj, yr) {
          fetch(`/viewgl_activitybs/${obj}/${yr}/`)
            .then(function (response) {
              
              return response.json();
            })
            .then(function (data) {
         
              if (data.status === "success") {
                
                populateModal(data.data , data.total_bal);
                
                modal.style.display = "block";
              } else {
    
              }
            })
            .catch(function (error) {
      
            });
        }
    
       
        var viewGLLinks = document.querySelectorAll(".viewgl_activitybs-link");
        viewGLLinks.forEach(function (link) {
          link.addEventListener("click", function (event) {
            event.preventDefault();
            
            var obj = link.dataset.obj;
            var yr = link.dataset.yr;
            fetchDataAndPopulateModal(obj , yr);
          });
        });
    
        
        var closeButton = modal.querySelector(".modal-footer button");
        closeButton.addEventListener("click", function () {
          modal.style.display = "none";
        });
      });
    

    
    
    
    
        
      function validateBudgetInput(event) {
        const input = event.target;
        const value = input.value;
    
        
        const regex = /^[0-9.]*$/;
    
        
        if (!regex.test(value)) {
            
            input.value = value.replace(/[^0-9.]/g, '');
        }
    }
      
       
        
      function extractNumericValue(content) {
        const match = content.match(/\(([^)]+)\)/);
        if (match) {
          const numericValue = parseFloat(match[1].replace(/,/g, '').trim());
          return isNaN(numericValue) ? 0 : -numericValue;
        } else {
          const numericValue = parseFloat(content.replace(/,/g, '').trim());
          return isNaN(numericValue) ? 0 : numericValue;
        }
      }
    
      /* --------------current assets total function ------------------*/
      function calculateCurrentAssets() {
        const rows = document.querySelectorAll('.current-assets-row');
        const totalCells = [];
      
        
        const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`current-assets-total-${cellId}`);
          
          totalCells.push(totalCell);
          
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
          const columnTotal = parseInt(columnTotals[i]);
          totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';
    
        }
      }
    
     /* --------------capital assets total function ------------------*/
      function calculateCapitalAssets() {
        const rows = document.querySelectorAll('.capital-assets-row');
        const totalCells = [];
      
        
        const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`capital-assets-total-${cellId}`);
          
          totalCells.push(totalCell);
          
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
          const columnTotal = parseInt(columnTotals[i]);
          totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';
        }
      }
    
      /* -------------- total  assets function ------------------*/
      function calculateTotalAssets() {
        const rows = document.querySelectorAll('.capital-assets-row,.current-assets-row');
        const totalCells = [];
      
        const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,16, 17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`total-assets-total-${cellId}`);
    
          totalCells.push(totalCell);
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
          const columnTotal = columnTotals[i]; // Don't parse the integer value here
          totalCell.textContent = columnTotal !== 0 ? '$' + columnTotal.toLocaleString() : '';
        }
      }
    
      /* --------------total current Liabilities function ------------------*/
      function calculateCurrentLiabilities() {
        const rows = document.querySelectorAll('.current-liabilities-row');
        const totalCells = [];
      
        
        const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`current-liabilities-total-${cellId}`);
          
          totalCells.push(totalCell);
          
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
        const columnTotal = parseInt(columnTotals[i]);
        const formattedTotal = columnTotal < 0 ? `(${Math.abs(columnTotal).toLocaleString()})` : columnTotal.toLocaleString();
        totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
       }
      }
    
    
      /* --------------total  Liabilities function ------------------*/
      function calculateTotalLiabilities() {
        const rows = document.querySelectorAll('.total-liabilities-row,.current-liabilities-row');
        const totalCells = [];
      
        
        const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`total-liabilities-total-${cellId}`);
          
          totalCells.push(totalCell);
          
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
        const columnTotal = parseInt(columnTotals[i]);
        const formattedTotal = columnTotal < 0 ? `(${Math.abs(columnTotal).toLocaleString()})` : columnTotal.toLocaleString();
        totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
      }
      }
    
    
      /* --------------total  Liabilities and nets assets function ------------------*/
      function calculateTotalLiabilitiesAndNetAssets() {
        const rows = document.querySelectorAll('.net-assets-row, .total-liabilities-row, .current-liabilities-row');
        const totalCells = [];
      
        
        const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];
      
        for (let i = 0; i < totalCellIds.length; i++) {
          const cellId = totalCellIds[i];
          const totalCell = document.getElementById(`liabilitiesandnet-total-${cellId}`);
          
          totalCells.push(totalCell);
          
        }
      
        let columnTotals = new Array(totalCells.length).fill(0);
      
        rows.forEach(row => {
          const cells = row.cells;
          for (let i = 3; i < cells.length; i++) {
            const cell = cells[i];
            const value = extractNumericValue(cell.textContent);
            if (!isNaN(value)) {
              for (let j = 0; j < totalCells.length; j++) {
                if (i === totalCellIds[j]) {
                  columnTotals[j] += value;
                  break;
                }
              }
            }
          }
        });
      
        for (let i = 0; i < totalCells.length; i++) {
          const totalCell = totalCells[i];
          const columnTotal = parseInt(columnTotals[i]);
          totalCell.textContent = columnTotal !== 0 ? '$' + columnTotal.toLocaleString() : '';
        }
      }
    
      
      function calculateFYTD() {
        const rows = document.querySelectorAll('.balancesheet-activity-row');
      
        rows.forEach(row => {
          const cells = row.cells;
          let rowTotal = 0;
          for (let i = 4; i <= 15; i++) {
            const cellValue = extractNumericValue(cells[i].textContent);
            if (!isNaN(cellValue)) {
              rowTotal += cellValue;
            }
          }
      
          
          
          const formattedTotal = rowTotal < 0 ? `(${Math.abs(rowTotal).toLocaleString()})` : rowTotal.toLocaleString();
          cells[16].textContent = rowTotal !== 0 ? formattedTotal : ''; //fytd
          cells[17].textContent = rowTotal !== 0 ? formattedTotal : ''; // as of last month
        });
      }
    
      function calculateFYTD2() {
        const rows = document.querySelectorAll('.liabilities-activity-row');
      
        rows.forEach(row => {
          const cells = row.cells;
          let rowTotal = 0;
          for (let i = 4; i <= 15; i++) {
            const cellValue = extractNumericValue(cells[i].textContent);
            if (!isNaN(cellValue)) {
              rowTotal -= cellValue;
            }
          }
      
          
          
          const formattedTotal = rowTotal < 0 ? `(${Math.abs(rowTotal).toLocaleString()})` : rowTotal.toLocaleString();
          cells[16].textContent = rowTotal !== 0 ? formattedTotal : ''; //fytd
          cells[17].textContent = rowTotal !== 0 ? formattedTotal : ''; // as of last month
        });
      }
    
    
      
      
      
            
            
    
    //   var updateValuesArray = [];
    
    //   //---------------- update for local revenue, spr , fpr ------------------
    //   function updateBS(event) {
    //    var tdElement = event.target;
    //    var rowNumber = tdElement.id.split("-")[2];
    //    var updateValueInput = document.getElementById("updatefye-" + rowNumber);
     
    //    var cellContent = tdElement.innerText.trim();
    
    
    //    var cleanedValue = cellContent.replace(/[^0-9.]/g, '');
     
    //    if (cleanedValue === "" || isNaN(parseFloat(cleanedValue))) {
    //      if (updateValueInput) {
    //        updateValueInput.value = 0; // Set to an empty string when cell is empty or not a number
    //      }
    //    } else {
    //      var value = parseFloat(cleanedValue);
         
     
    //      if (!isNaN(value) || value === 0) {
    //        if (updateValueInput) {
           
     
    //          updateValueInput.value = value;
    //        }
    //      } else {
    //        if (updateValueInput) {
    
     
    //          updateValueInput.value = 0;
    //        }
    //      }
     
    //      // Update the value in the updateValuesArray based on the rowNumber
    //      updateValuesArray[rowNumber - 1] = !isNaN(value) ? value : "";
    //    }
    //  }
    
    //   var tbodyElements = document.querySelectorAll('.current-assets-row, .capital-assets-row, .current-liabilities-row, .total-liabilities-row, .net-assets-row ');
    // tbodyElements.forEach(function(tbodyElement) {
    //   tbodyElement.addEventListener("input", updateBS);
    // });
    // var tdElements = document.querySelectorAll('.current-assets-row td[contenteditable], .capital-assets-row td[contenteditable], .current-liabilities-row td[contenteditable], .total-liabilities-row td[contenteditable], .net-assets-row td[contenteditable] ');
    // tdElements.forEach(function(tdElement) {
    //   var rowNumber = tdElement.id.split("-")[2];
    //   updateBS({ target: tdElement });
    
    // });
    
    function hideRowsByClass(className) {
      const rows = document.querySelectorAll(className);
      rows.forEach(row => {
        row.style.display = 'none';
      });
    }
    
    function hideRowsOnLoad() {
      const collapsedRows = document.querySelectorAll('.collapsed');
      collapsedRows.forEach(row => {
        row.style.display = 'none';
      });
    
    
      hideRowsByClass('.balancesheet-activity-row');
      hideRowsByClass('.liabilities-activity-row');
      
    }
    
    
      window.addEventListener('DOMContentLoaded', () => {
        hideRowsOnLoad();
      
          const editableCells = document.querySelectorAll('.current-assets-row td[contenteditable]');
          
          editableCells.forEach(cell => {
            cell.addEventListener('input', () => {
              calculateCurrentAssets();
              calculateCapitalAssets();
              calculateTotalAssets();
              calculateCurrentLiabilities();
              calculateTotalLiabilities();
              calculateTotalLiabilitiesAndNetAssets();
              calculateFYTD();
              
            
              
              
            });
          });
          
          
          calculateCurrentAssets();
          calculateCapitalAssets();
          calculateTotalAssets();
          calculateCurrentLiabilities();
          calculateTotalLiabilities();
          calculateTotalLiabilitiesAndNetAssets();
          calculateFYTD();
          calculateFYTD2();
          
      });
    
      
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
  
    rows.forEach(row => {
      const cells = row.cells;
      for (let i = 3; i < cells.length; i++) {
        const cell = cells[i];
        const value = extractNumericValue(cell.textContent);
        if (!isNaN(value)) {
          for (let j = 0; j < totalCells.length; j++) {
            if (i === totalCellIds[j]) {
              columnTotals[j] += value;
              break;
            }
          }
        }
      }
    });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
      const columnTotal = parseInt(columnTotals[i]);
      totalCell.textContent = columnTotal !== 0 ? '$' + columnTotal.toLocaleString() : '';
    }
  }

  
  function calculateFYTD() {
    const rows = document.querySelectorAll('.balancesheet-activity-row');
  
    rows.forEach(row => {
      const cells = row.cells;
      let rowTotal = 0;
      for (let i = 4; i <= 15; i++) {
        const cellValue = extractNumericValue(cells[i].textContent);
        if (!isNaN(cellValue)) {
          rowTotal += cellValue;
        }
      }
  
      
      
      const formattedTotal = rowTotal < 0 ? `(${Math.abs(rowTotal).toLocaleString()})` : rowTotal.toLocaleString();
      cells[16].textContent = rowTotal !== 0 ? formattedTotal : '';
      cells[17].textContent = rowTotal !== 0 ? formattedTotal : '';
    });
  }


  
  
  
        
        

//   var updateValuesArray = [];

//   //---------------- update for local revenue, spr , fpr ------------------
//   function updateBS(event) {
//    var tdElement = event.target;
//    var rowNumber = tdElement.id.split("-")[2];
//    var updateValueInput = document.getElementById("updatefye-" + rowNumber);
 
//    var cellContent = tdElement.innerText.trim();


//    var cleanedValue = cellContent.replace(/[^0-9.]/g, '');
 
//    if (cleanedValue === "" || isNaN(parseFloat(cleanedValue))) {
//      if (updateValueInput) {
//        updateValueInput.value = 0; // Set to an empty string when cell is empty or not a number
//      }
//    } else {
//      var value = parseFloat(cleanedValue);
     
 
//      if (!isNaN(value) || value === 0) {
//        if (updateValueInput) {
       
 
//          updateValueInput.value = value;
//        }
//      } else {
//        if (updateValueInput) {

 
//          updateValueInput.value = 0;
//        }
//      }
 
     
//      updateValuesArray[rowNumber - 1] = !isNaN(value) ? value : "";
//    }
//  }

//   var tbodyElements = document.querySelectorAll('.current-assets-row, .capital-assets-row, .current-liabilities-row, .total-liabilities-row, .net-assets-row ');
// tbodyElements.forEach(function(tbodyElement) {
//   tbodyElement.addEventListener("input", updateBS);
// });
// var tdElements = document.querySelectorAll('.current-assets-row td[contenteditable], .capital-assets-row td[contenteditable], .current-liabilities-row td[contenteditable], .total-liabilities-row td[contenteditable], .net-assets-row td[contenteditable] ');
// tdElements.forEach(function(tdElement) {
//   var rowNumber = tdElement.id.split("-")[2];
//   updateBS({ target: tdElement });

// });

function hideRowsByClass(className) {
  const rows = document.querySelectorAll(className);
  rows.forEach(row => {
    row.style.display = 'none';
  });
}

function hideRowsOnLoad() {
  const collapsedRows = document.querySelectorAll('.collapsed');
  collapsedRows.forEach(row => {
    row.style.display = 'none';
  });

  // Hide local revenue rows
  hideRowsByClass('.balancesheet-activity-row');
  hideRowsByClass('.liabilities-activity-row');
  // Hide state program revenue rows

}


  window.addEventListener('DOMContentLoaded', () => {
    hideRowsOnLoad();
  
      // const editableCells = document.querySelectorAll('.current-assets-row td[contenteditable]');
      
      // editableCells.forEach(cell => {
      //   cell.addEventListener('input', () => {
      //     calculateCurrentAssets();
      //     calculateCapitalAssets();
      //     calculateTotalAssets();
      //     calculateCurrentLiabilities();
      //     calculateTotalLiabilities();
      //     calculateTotalLiabilitiesAndNetAssets();
          // calculateFYTD();
          
        
          
          
      //   });
      // });
      
      
      // calculateCurrentAssets();
      // calculateCapitalAssets();
      // calculateTotalAssets();
      // calculateCurrentLiabilities();
      // calculateTotalLiabilities();
      // calculateTotalLiabilitiesAndNetAssets();
      // calculateFYTD();
      
  });

  


