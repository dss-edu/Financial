$(document).ready(function() {
///////////////////////////////////   PDF   /////////////////////////////////////////////
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


///////////////////////////   TABLE   //////////////////////////////////
  calculateNetCashTotal();

  function validateBudgetInput(event) {
    const input = event.target;
    const value = input.value;
    const regex = /^[0-9.]*$/;
    if (!regex.test(value)) {
        
        input.value = value.replace(/[^0-9.]/g, '');
    }
}

function toggleRow(rowId) {
  for (var i = 0; i <= 50; i++) {
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

function toggleColumns() {
    var toggleButton = document.getElementById("toggle-button");
    var collapsedColumns = document.getElementsByClassName("collapsed");
    for (var i = 0; i < collapsedColumns.length; i++) {
      if (collapsedColumns[i].style.display === "none") {
        collapsedColumns[i].style.display = "table-cell";
        toggleButton.innerHTML = "-";
      } else {
        collapsedColumns[i].style.display = "none";
        toggleButton.innerHTML = "+";
      }
    }
  }

  function hideRowsByClass(className) {
    const rows = document.querySelectorAll(className);
    rows.forEach(row => {
      row.style.display = 'none';
    });
  }
  function extractNumericValue(content) {
    const match = content.match(/-?\$?\(?([\d,.]+)\)?/);
    if (match) {
      const numericValue = parseFloat(match[1].replace(/,/g, '').trim());
      return isNaN(numericValue) ? 0 : (content.includes('(') ? -numericValue : numericValue);
    } else {
      return 0;
    }
  }
  

  /* NET CASH FLOWS FROM OPEARTING ACTIVITIES TOTAL*/
  function calculateNetCashTotal() {
    const rows = document.querySelectorAll('.netcash-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17,18 ];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`netcash-total-${cellId}`);
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
      const columnTotal = parseFloat(columnTotals[i]);
      if (columnTotal !== 0) {
        const formattedTotal = columnTotal.toLocaleString();
        totalCell.textContent = columnTotal < 0 ? '($' + formattedTotal.replace('-', '') + ')' : '$' + formattedTotal;
      } else {
        totalCell.textContent = '';
      }
    }
  }
});
