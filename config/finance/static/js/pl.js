

  

  function validateBudgetInput(event) {
    const input = event.target;
    const value = input.value;

    
    const regex = /^[0-9.]*$/;

    
    if (!regex.test(value)) {
        
        input.value = value.replace(/[^0-9.]/g, '');
    }
}

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

var columnsCollapsed = true;


function toggleColumns() {
  console.log(showcurr)
    const dataTable = document.getElementById('data-table');
    const toggleButton = document.getElementById("toggle-button");
    const toggleIcon = document.getElementById("toggle-icon");
    const collapsedColumns = dataTable.getElementsByClassName("collapsed");
    const show = document.getElementById("showCurrentMonth");
    const hide = document.getElementById("hideCurrentMonth");


    for (var i = 0; i < collapsedColumns.length; i++) {
        if (columnsCollapsed) {
            collapsedColumns[i].style.display = "table-cell";
            toggleIcon.className = "fa-solid fa-chevron-left";

            // hide.style.display = "flex";
            // show.style.display = "none";

        } else {
            collapsedColumns[i].style.display = "none";
            toggleIcon.className = "fa-solid fa-chevron-right";
            // hide.style.display = "none";
            // show.style.display = "flex";
            hidecurrentMonth();
        }
    }


    // Toggle the state
    columnsCollapsed = !columnsCollapsed;
}

var showcurr = "";
function showcurrentMonth() {
  showcurr = "True";

  const show = document.getElementById("showCurrentMonth");
  const hide = document.getElementById("hideCurrentMonth");

  
  hide.style.display = "flex";
  show.style.display = "none";
  var lm = last_month_number ;

  var columnsToHide2 = document.querySelectorAll(
    "#data-table th:nth-child(" +
    lm +
    "), #data-table td:nth-child(" +
    lm +
    ")"
  );

  for (var k = 0; k < columnsToHide2.length; k++) {

    columnsToHide2[k].style.display = "table-cell";
    
  }

  var colorColumn = document.querySelectorAll(
    "#data-table td:nth-child(" + lm + ")"
  );

  for (var k = 0; k < colorColumn.length; k++) {
    colorColumn[k].classList.add("hidden-column");
  
    
  }


}

             
function hidecurrentMonth() {

  const show = document.getElementById("showCurrentMonth");
  const hide = document.getElementById("hideCurrentMonth");
  hide.style.display = "none";
  show.style.display = "flex";

  showcurr = "False";
  var lm = last_month_number ;

  var columnsToHide2 = document.querySelectorAll(
    "#data-table th:nth-child(" +
    lm +
    "), #data-table td:nth-child(" +
    lm +
    ")"
  );

  for (var k = 0; k < columnsToHide2.length; k++) {
    
    columnsToHide2[k].style.display = "none";
    
  }


}

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
    //hideRowsByClass('.local-revenue-row');
    // Hide state program revenue rows
    //hideRowsByClass('.spr-row');
    // Hide federal program revenue rows
    //hideRowsByClass('.fpr-row');
    //hideRowsByClass('.PCS-row');
    //hideRowsByClass('.sm-row');
    //hideRowsByClass('.ooe-row');
    //hideRowsByClass('.expense-row');
    //hideRowsByClass('.payroll-row');
    // hideRowsByClass('.total-row1');
    //hideRowsByClass('.DnA-row');
   
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
  
  
 


window.addEventListener('DOMContentLoaded', () => {
  hideRowsOnLoad();




 
  

/* --------------revenue function ------------------*/
function calculateLocalRevenueTotal() {
  const rows = document.querySelectorAll('.local-revenue-row');
  const totalCells = [];

  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`local-revenue-total-${cellId}`);
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


      
      
      

      /* --------------spr function ------------------*/
function calculateSPRTotal() {
  const rows = document.querySelectorAll('.spr-row');
  const totalCells = [];

  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`spr-total-${cellId}`);
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

      
    
     /* --------------spr function ------------------*/
 function calculateFPRTotal() {
  const rows = document.querySelectorAll('.fpr-row');
  const totalCells = [];

   
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16];

   for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`fpr-total-${cellId}`);
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
    const formattedTotal = columnTotal < 0 ? `$(${Math.abs(columnTotal).toLocaleString()})` : `$${columnTotal.toLocaleString()}`;
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
  }
 }


/* ------row 1 total --------- */
function calculaterow1Total() {
  const rows = document.querySelectorAll('.total-row1');
  const totalCells = [];

  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`total1-${cellId}`);
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
/*----------- calculate the local, state program and federal program revenue ----*/
function calculateLSFTotals() {
  const rows = document.querySelectorAll('.fpr-row, .spr-row, .local-revenue-row');
  const totalCells = [];

  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`total-${cellId}`);
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



/*---- year to date totals ------------*/
function calculateYTD() {
  const rows = document.querySelectorAll('.local-revenue-row, .spr-row, .fpr-row, .total-row1, .DnA-row, .ttl-lr-row, .ttl-fpr-row, .ttl-spr-row, .ttl-1-row, .ttl-0-row');

  rows.forEach(row => {
    const cells = row.cells;
    let rowTotal = 0;
    for (let i = 5; i <= 18; i++) {
      const cellValue = extractNumericValue(cells[i].textContent);
      if (!isNaN(cellValue)) {
        rowTotal += cellValue;
      }
    }

    
    if (row.classList.contains('ttl-lr-row') || row.classList.contains('ttl-fpr-row') || row.classList.contains('ttl-spr-row') || row.classList.contains('ttl-1-row') || row.classList.contains('ttl-0-row')) {
      cells[17].textContent = rowTotal !== 0 ? '$' + rowTotal.toLocaleString() : '';
    } else {
      cells[17].textContent = rowTotal !== 0 ? rowTotal.toLocaleString() : '';
    }
  });
}

function calculateYTD2() {
  const rows = document.querySelectorAll('.EOC-row, .EOC-total-row');

  rows.forEach(row => {
    const cells = row.cells;
    let rowTotal = 0;
    for (let i = 5; i <= 18; i++) {
      const cellValue = extractNumericValue(cells[i].textContent);
      if (!isNaN(cellValue)) {
        rowTotal += cellValue;
      }
    }

    
    if (row.classList.contains('EOC-total-row')) {
      cells[17].textContent = rowTotal !== 0 ? '$' + rowTotal.toLocaleString() : '';
    } else {
      cells[17].textContent = rowTotal !== 0 ? rowTotal.toLocaleString() : '';
    }
  });
}



/*----- Calculate Variances for spr  to 1st total------------ */
function CalculateVariances1() {
  const rows = document.querySelectorAll(' .spr-row, .fpr-row,  .ttl-lr-row, .ttl-fpr-row, .ttl-spr-row, .ttl-0-row');

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let budgetCell = cells[4];
    
    
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, ''));
    let budgetValue = parseFloat(budgetCell.textContent.trim().replace('$', '').replace(/,/g, ''));

    if(isNaN(ytdValue)){
      ytdValue=0;

    }
    if(isNaN(budgetValue)){
      budgetValue=0;

    }
    
    
    const difference = ytdValue - budgetValue ;

   
    let formattedDifference = '';
    if (!isNaN(difference)) {
      if (difference < 0) {
        formattedDifference = `(${Math.abs(difference).toLocaleString()})`;
      } else {
        formattedDifference = `${difference.toLocaleString()}`;
      }
    }
    
    
    if (row.classList.contains('ttl-lr-row') || row.classList.contains('ttl-fpr-row') || row.classList.contains('ttl-spr-row') || row.classList.contains('ttl-1-row') || row.classList.contains('ttl-0-row')) {
      formattedDifference = `$${formattedDifference}`;
    }
    
    
    cells[18].textContent = difference !== 0 ? formattedDifference : '';
  });
}

/*----- Calculate Variances for total row 1 ex. data 2 from fund------------ */
function CalculateVariances2() {
  const rows = document.querySelectorAll('.total-row1, .ttl-1-row, .DnA-row ' );

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let budgetCell = cells[4];
    
    
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, ''));
    let budgetValue = parseFloat(budgetCell.textContent.trim().replace('$', '').replace(/,/g, ''));

    if(isNaN(ytdValue)){
      ytdValue=0;

    }
    if(isNaN(budgetValue)){
      budgetValue=0;

    }
    
    
    const difference = budgetValue- ytdValue ;

    
    let formattedDifference = '';
    if (!isNaN(difference)) {
      if (difference < 0) {
        formattedDifference = `(${Math.abs(difference).toLocaleString()})`;
      } else {
        formattedDifference = `${difference.toLocaleString()}`;
      }
    }
    
  
    if (row.classList.contains('ttl-lr-row') || row.classList.contains('ttl-fpr-row') || row.classList.contains('ttl-spr-row') || row.classList.contains('ttl-1-row') || row.classList.contains('ttl-0-row')) {
      formattedDifference = `$${formattedDifference}`;
    }
    
   
    cells[18].textContent = difference !== 0 ? formattedDifference : '';
  });
}


/*----- Calculate VAR. for dna row total------------ */
function CalculateDnAVar42() {
  const rows = document.querySelectorAll(' .DnA-row-total');

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let ammendedCell = cells[4];

    // Parse the values from the cells and remove any non-numeric characters
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, ''));
    let ammendedValue = parseFloat(ammendedCell.textContent.trim().replace('$', '').replace(/,/g, ''));



    if (isNaN(ytdValue)) {
      ytdValue = 0;
    }
    if (isNaN(ammendedValue)) {
      ammendedValue = 1; // To avoid division by zero
    }

    // Calculate the percentage
    const percentage = (ytdValue / ammendedValue) * 100;

    // Format the percentage with parentheses for negative values and append '%'
    let formattedPercentage = '';
    if (!isNaN(percentage)) {
      if (percentage < 0) {
        formattedPercentage = `(${Math.abs(percentage.toFixed(0))}%)`;
      } else {
        formattedPercentage = `${percentage.toFixed(0)}%`;
      }
    }

   

    // Display the result in cell[19]
    cells[19].textContent = percentage !== 0 ? formattedPercentage : '';
  });
}

function CalculateVar42() {
  const rows = document.querySelectorAll('.total-row1, .ttl-1-row,  .ttl-lr-row, .ttl-fpr-row, .ttl-spr-row, .ttl-0-row , .DnA-row');

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let ammendedCell = cells[3];

    // Parse the values from the cells and remove any non-numeric characters
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, ''));
    let ammendedValue = parseFloat(ammendedCell.textContent.trim().replace('$', '').replace(/,/g, ''));

    if (isNaN(ytdValue)) {
      ytdValue = 0;
    }
    if (isNaN(ammendedValue)) {
      ammendedValue = 1; // To avoid division by zero
    }

    // Calculate the percentage
    const percentage = (ytdValue / ammendedValue) * 100;

    // Format the percentage with parentheses for negative values and append '%'
    let formattedPercentage = '';
    if (!isNaN(percentage)) {
      if (percentage < 0) {
        formattedPercentage = `(${Math.abs(percentage.toFixed(0))}%)`;
      } else {
        formattedPercentage = `${percentage.toFixed(0)}%`;
      }
    }

   

    // Display the result in cell[19]
    cells[19].textContent = percentage !== 0 ? formattedPercentage : '';
  });
}

var updateValuesArray = [];

 //---------------- update for local revenue, spr , fpr ------------------
 function updateLR(event) {
  var tdElement = event.target;
  var rowNumber = tdElement.id.split("-")[2];
  var updateValueInput = document.getElementById("updatevalue-" + rowNumber);

  var cellContent = tdElement.innerText.trim();
 

  var cleanedValue = cellContent.replace(/[^0-9.]/g, '');

  if (cleanedValue === "" || isNaN(parseFloat(cleanedValue))) {
    if (updateValueInput) {
      updateValueInput.value = 0; // Set to an empty string when cell is empty or not a number
    }
  } else {
    var value = parseFloat(cleanedValue);
    

    if (!isNaN(value) || value === 0) {
      if (updateValueInput) {
       

        updateValueInput.value = value;
      }
    } else {
      if (updateValueInput) {
       

        updateValueInput.value = 0;
      }
    }

  
    updateValuesArray[rowNumber - 1] = !isNaN(value) ? value : "";
  }
}


var tbodyElements = document.querySelectorAll('.local-revenue-row, .spr-row, .fpr-row ');
tbodyElements.forEach(function(tbodyElement) {
  tbodyElement.addEventListener("input", updateLR);
});
var tdElements = document.querySelectorAll('.local-revenue-row td[contenteditable], .spr-row td[contenteditable], .fpr-row td[contenteditable]');
tdElements.forEach(function(tdElement) {
  var rowNumber = tdElement.id.split("-")[2];
  updateLR({ target: tdElement });

});


var updateValuesArray2 = [];
//------------- update for func -----------
function updateFUNC(event) {
  var tdElement = event.target;
  var rowNumber = tdElement.id.split("-")[2];
  var updateValueInput = document.getElementById("updatevalue-func-" + rowNumber);

  var cellContent = tdElement.innerText.trim();


  var cleanedValue = cellContent.replace(/[^0-9.]/g, ''); // Remove non-numeric characters
  cleanedValue = cleanedValue.replace(/,/g, '');

  if (cleanedValue === "" || isNaN(parseFloat(cleanedValue))) {
    if (updateValueInput) {
      updateValueInput.value = 0; // Set to an empty string when cell is empty or not a number
    }
  }
   else {
    var value = parseFloat(cleanedValue);
    

    if (!isNaN(value) || value === 0) {
      if (updateValueInput) {


        updateValueInput.value = value;
      }
    } else {
      if (updateValueInput) {
        

        updateValueInput.value = 0;
      }
    }

    // Update the value in the updateValuesArray based on the rowNumber
    updateValuesArray2[rowNumber - 1] = !isNaN(value) ? value : "";
  }
}


var totalRow1Elements = document.querySelectorAll('.total-row1');
totalRow1Elements.forEach(function(totalRow1Element) {
  totalRow1Element.addEventListener("input", updateFUNC);
});

var funcElements = document.querySelectorAll('.total-row1 td[contenteditable]');
funcElements.forEach(function(tdElement) {
  var rowNumber = tdElement.id.split("-")[1];
  updateFUNC({ target: tdElement });
});

function SurplusDeficitTotal() {
  const rows = document.querySelectorAll('.fpr-row, .spr-row, .local-revenue-row');
  const totalCells = [];
  const rows2 = document.querySelectorAll('.total-row1');

  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`surplus-total-${cellId}`);
    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);
  let columnTotals2 = new Array(totalCells.length).fill(0);

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

  rows2.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = extractNumericValue(cell.textContent);
      if (!isNaN(value)) {
        for (let j = 0; j < totalCells.length; j++) {
          if (i === totalCellIds[j]) {
            columnTotals2[j] += value;
            break;
          }
        }
      }
    }
  });

  for (let i = 0; i < totalCells.length; i++) {
    const totalCell = totalCells[i];
    const columnTotal = columnTotals[i] - columnTotals2[i];
    const formattedTotal = columnTotal < 0 ? `$(${Math.abs(columnTotal).toLocaleString()})` : `$${columnTotal.toLocaleString()}`;
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
  }
 }

 
function CalculateVar42forSurplus() {
  const rows = document.querySelectorAll('.surplus-ttl-1-row' );

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let ammendedCell = cells[3];

    // Parse the values from the cells and remove any non-numeric characters
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));
    let ammendedValue = parseFloat(ammendedCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));

    if (isNaN(ytdValue)) {
      ytdValue = 0;
    }
    if (isNaN(ammendedValue)) {
      ammendedValue = 1; // 
    }



    // Calculate the percentage
    const percentage = (ytdValue / ammendedValue) * 100;

    
    
    let formattedPercentage = `${percentage.toFixed(0)}%`;
     

   

    
    cells[19].textContent = percentage !== 0 ? formattedPercentage : '';
  });
}

function CalculateVar42forNetSurplus() {
  const rows = document.querySelectorAll('.net-surplus-ttl-1-row, .EOC-row, .EOC-total-row, .netincome-total-row' );

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let ammendedCell = cells[3];

    // Parse the values from the cells and remove any non-numeric characters
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));
    let ammendedValue = parseFloat(ammendedCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));

    if (isNaN(ytdValue)) {
      ytdValue = 0;
    }
    if (isNaN(ammendedValue)) {
      ammendedValue = 1; // 
    }



    // Calculate the percentage
    const percentage = (ytdValue / ammendedValue) * 100;

    
    
    let formattedPercentage = `${percentage.toFixed(0)}%`;
     

   

    
    cells[19].textContent = percentage !== 0 ? formattedPercentage : '';
  });
}


//variance for surplus deficits before depreciation and Expense by object
function CalculateVariances3() {
  const rows = document.querySelectorAll('.surplus-ttl-1-row, .net-surplus-ttl-1-row ');

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let budgetCell = cells[4];
    
    
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));
    let budgetValue = parseFloat(budgetCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));

    if(isNaN(ytdValue)){
      ytdValue=0;

    }
    if(isNaN(budgetValue)){
      budgetValue=0;

    }
 
    
    const difference = ytdValue - budgetValue ;

    
    let formattedDifference = '';
    if (!isNaN(difference)) {
      if (difference < 0) {
        formattedDifference = `(${Math.abs(difference).toLocaleString()})`;
      } else {
        formattedDifference = `${difference.toLocaleString()}`;
      }
    }
    
  
    if (row.classList.contains('surplus-ttl-1-row')) {
      formattedDifference = `$${formattedDifference}`;
    }
    
    
   
    cells[18].textContent = difference !== 0 ? formattedDifference : '';

  });
}
function CalculateVariances4() {
  const rows = document.querySelectorAll('.EOC-row  ');

  rows.forEach(row => {
    const cells = row.cells;
    let ytdCell = cells[17];
    let budgetCell = cells[4];
    
    
    let ytdValue = parseFloat(ytdCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));
    let budgetValue = parseFloat(budgetCell.textContent.trim().replace('$', '').replace(/,/g, '').replace('(', '-').replace(')', ''));

    if(isNaN(ytdValue)){
      ytdValue=0;

    }
    if(isNaN(budgetValue)){
      budgetValue=0;

    }
 
    
    const difference = ytdValue - budgetValue ;

    
    let formattedDifference = '';
    if (!isNaN(difference)) {
      if (difference < 0) {
        formattedDifference = `(${Math.abs(difference).toLocaleString()})`;
      } else {
        formattedDifference = `${difference.toLocaleString()}`;
      }
    }
    
  
    if (row.classList.contains('surplus-ttl-1-row')) {
      formattedDifference = `$${formattedDifference}`;
    }
    
    
   
    cells[18].textContent = difference !== 0 ? formattedDifference : '';

  });
}









/* --------------PayrollCostTotal function ------------------*/
function PayrollCostTotal() {
  const rows = document.querySelectorAll('.payroll-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`payroll-row-total-${cellId}`);
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
    totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';
  }
}

//----------------PROFESSIONAL AND CONT SVCS
function PCSTotal() {
  const rows = document.querySelectorAll('.PCS-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`pcs-total-${cellId}`);
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

function SMTotal() {
  const rows = document.querySelectorAll('.sm-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`sm-total-${cellId}`);
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
    totalCell.textContent = columnTotal !== 0 ?  columnTotal.toLocaleString() : '';
  }
}


//---------- OTHER OPERATING EXPENSES TOTAL 
function OOETotal() {
  const rows = document.querySelectorAll('.ooe-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`ooe-total-${cellId}`);
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

function OperatingExpenseTotal() {
  const rows = document.querySelectorAll('.expense-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`oe-total-${cellId}`);
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
    totalCell.textContent = columnTotal !== 0 ? '$' +  columnTotal.toLocaleString() : '';
  }
}

// -------------- TOTAL EXPENSE FUNCTION
/*function ExpenseTotal() {
  const rows = document.querySelectorAll('.ooe-row, .sm-row, .PCS-row , .payroll-row , expense-row');
  const totalCells = [];

  
  const totalCellIds = [ 5,6,7,8,9,10,11,12,13,14,15,16,17,18,19];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`expense-total-${cellId}`);
    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = parseFloat(cell.textContent.replace(/,/g, '').trim());
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
} */



//----------------Total Expense by object function
function totalEOC() { 
  const rows = document.querySelectorAll('.EOC-row');
  const totalCells = [];

  
  const totalCellIds = [ 3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`EOC-total-${cellId}`);
    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = parseFloat(cell.textContent.replace(/,/g, '').trim());
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
    
    totalCell.textContent = columnTotal !== 0 ?  '$' + columnTotal.toLocaleString() : '';
  }
}

//----------------Total Expense by object function
function DandATotal() {
  const rows = document.querySelectorAll('.DnA-row');
  const totalCells = [];

  
  const totalCellIds = [ 3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`DnA-total-${cellId}`);
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

function NetSurplusTotal() {
  const rows = document.querySelectorAll('.fpr-row, .spr-row, .local-revenue-row');
  const totalCells = [];
  const rows2 = document.querySelectorAll('.total-row1');
  const rows3 = document.querySelectorAll('.DnA-row');
  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`net-total-${cellId}`);
    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);
  let columnTotals2 = new Array(totalCells.length).fill(0);
  let columnTotals3 = new Array(totalCells.length).fill(0);

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

  rows2.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = extractNumericValue(cell.textContent);
      if (!isNaN(value)) {
        for (let j = 0; j < totalCells.length; j++) {
          if (i === totalCellIds[j]) {
            columnTotals2[j] += value;
            break;
          }
        }
      }
    }
  });

  rows3.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = extractNumericValue(cell.textContent);
      if (!isNaN(value)) {
        for (let j = 0; j < totalCells.length; j++) {
          if (i === totalCellIds[j]) {
            columnTotals3[j] += value;
       
            break;
          }
        }
      }
    }
  });

  for (let i = 0; i < totalCells.length; i++) {
    const totalCell = totalCells[i];
 
    const columnTotal = (columnTotals[i] - columnTotals2[i])-columnTotals3[i];
    const formattedTotal = columnTotal < 0 ? `$(${Math.abs(columnTotal).toLocaleString()})` : `$${columnTotal.toLocaleString()}`;
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
  }
 }

 

 

 function NetIncomeVariance() {
  const rows = document.querySelectorAll('.ttl-0-row');
  const totalCells = [];
  const rows2 = document.querySelectorAll('.EOC-row');
 
  
  const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15,16,17,18];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`netincome-total-${cellId}`);
    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);
  let columnTotals2 = new Array(totalCells.length).fill(0);


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

  rows2.forEach(row => {
    const cells = row.cells;
    for (let i = 3; i < cells.length; i++) {
      const cell = cells[i];
      const value = extractNumericValue(cell.textContent);
      
      if (!isNaN(value)) {
        for (let j = 0; j < totalCells.length; j++) {
          if (i === totalCellIds[j]) {
            columnTotals2[j] += value;
            break;
          }
        }
      }
    }
  });



  for (let i = 0; i < totalCells.length; i++) {
    const totalCell = totalCells[i];
 
    const columnTotal = columnTotals[i] - columnTotals2[i];
   
    const formattedTotal = columnTotal < 0 ? `$(${Math.abs(columnTotal).toLocaleString()})` : `$${columnTotal.toLocaleString()}`;
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
  }
 }

 


 
    // const editableCells = document.querySelectorAll('.local-revenue-row td[contenteditable], .spr-row td[contenteditable], .fpr-row td[contenteditable], .total-row1 td[contenteditable]');
    // editableCells.forEach(cell => {
    //   cell.addEventListener('input', () => {
        
    //     // calculateLocalRevenueTotal();
    //     calculateSPRTotal();
    //     calculateFPRTotal();
    //     calculaterow1Total();
    //     calculateLSFTotals();
        
    //     CalculateVariances1();
    //     CalculateVariances2();
    //     CalculateVar42();
    //     SurplusDeficitTotal();
    //     CalculateVariances3();
        
    //     CalculateVar42forSurplus();
    //     PayrollCostTotal();
    //     PCSTotal();
    //     SMTotal();
    //     OperatingExpenseTotal();
    //     OOETotal();
    //     //ExpenseTotal();
    //     totalEOC();
    //     DandATotal();
    //     NetSurplusTotal();
    //     NetIncomeTotal();
    //     CalculateVar42forNetSurplus();
        
        
    //   });
    // });
   
    // // calculateLocalRevenueTotal();
    // calculateSPRTotal();
    // calculateFPRTotal();
    // calculaterow1Total();
    // calculateLSFTotals();
    // calculateYTD();
    // CalculateVariances1();
    // CalculateVariances2();
    // CalculateVar42();
    // SurplusDeficitTotal();
   
   
    // PayrollCostTotal();
    // PCSTotal();
    // SMTotal();
    // OperatingExpenseTotal();
    // OOETotal();
    // //ExpenseTotal();
 
    // CalculateVar42forSurplus();
    
    // DandATotal();
    // NetSurplusTotal();
    
    // CalculateVariances3();
    // calculateYTD2();
    // CalculateVariances4();
    // totalEOC();
    // NetIncomeVariance();
    // CalculateVar42forNetSurplus();
    // CalculateDnAVar42();
  
});
