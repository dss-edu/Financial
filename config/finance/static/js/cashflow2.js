window.addEventListener('DOMContentLoaded', () => {
    // calculateYTD(); //1
    // calculateOperatingTotal(); //2
    // calculateNetCashTotal(); //2
    // NICTotal(); //3
    // calculateYTD2();
  });

  
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
  

    /* NET CASH FLOWS FROM OPEARTING ACTIVITIES TOTAL*/
  function calculateNetCashTotal() {
    const rows = document.querySelectorAll('.netcash-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15];
  
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
  

   /* NET CASH FLOWS FROM INVESTING ACTIVITIES TOTAL*/
  function calculateOperatingTotal() {
    const rows = document.querySelectorAll('.investing-row');
    const totalCells = [];
  
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14,15];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`investing-total-${cellId}`);
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


       /* NET INCREASE (DECREASE) IN CASH*/
  function NICTotal() {
    const rows = document.querySelectorAll('.netcash-row-total, .investing-row-total');
    const totalCells = [];
    // const rows2 = document.querySelectorAll('.total-row1');
    
    
    const totalCellIds = [3,4, 5,6,7,8,9,10,11,12,13,14];
  
    for (let i = 0; i < totalCellIds.length; i++) {
      const cellId = totalCellIds[i];
      const totalCell = document.getElementById(`NIC-total-${cellId}`);
      totalCells.push(totalCell);
    }
  
    let columnTotals = new Array(totalCells.length).fill(0);
    // let columnTotals2 = new Array(totalCells.length).fill(0);
    // let columnTotals3 = new Array(totalCells.length).fill(0);
  
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
  
    // rows2.forEach(row => {
    //   const cells = row.cells;
    //   for (let i = 3; i < cells.length; i++) {
    //     const cell = cells[i];
    //     const value = parseFloat(cell.textContent.replace(/,/g, '').trim());
    //     if (!isNaN(value)) {
    //       for (let j = 0; j < totalCells.length; j++) {
    //         if (i === totalCellIds[j]) {
    //           columnTotals2[j] += value;
    //           break;
    //         }
    //       }
    //     }
    //   }
    // });
  
    // rows3.forEach(row => {
    //   const cells = row.cells;
    //   for (let i = 3; i < cells.length; i++) {
    //     const cell = cells[i];
    //     const value = extractNumericValue(cell.textContent);
    //     if (!isNaN(value)) {
    //       for (let j = 0; j < totalCells.length; j++) {
    //         if (i === totalCellIds[j]) {
    //           columnTotals3[j] += value;
         
    //           break;
    //         }
    //       }
    //     }
    //   }
    // });
  
    for (let i = 0; i < totalCells.length; i++) {
      const totalCell = totalCells[i];
   
      const columnTotal = parseFloat(columnTotals[i]);
      const formattedTotal = columnTotal < 0 ? `$(${Math.abs(columnTotal).toLocaleString()})` : `$${columnTotal.toLocaleString()}`;
      totalCell.textContent = columnTotal !== 0 ? formattedTotal : '';
    }
   }

  function calculateYTD() {
    const rows = document.querySelectorAll('.netcash-row, .investing-row');
  
    rows.forEach(row => {
      const cells = row.cells;
      let rowTotal = 0;
      for (let i = 5; i <= 15; i++) {
        const cellValue = extractNumericValue(cells[i].textContent);
        if (!isNaN(cellValue)) {
          rowTotal += cellValue;
        }
      }
  
    //   if (row.classList.contains('EOC-total-row')) {
    //     cells[17].textContent = rowTotal !== 0 ? '$' + rowTotal.toLocaleString() : '';
    //   } else {
      if (rowTotal < 0) {
        cells[15].textContent = `(${Math.abs(rowTotal).toLocaleString()})`;
      } else {
        cells[15].textContent = rowTotal !== 0 ? rowTotal.toLocaleString() : '';
      }
   
    //   }
    });
  }
  
//   function calculateYTD() {
//     const rows = document.querySelectorAll('.CBP-row, .CEP-row');
  
//     rows.forEach(row => {
//       const cells = row.cells;
//       let rowTotal = 0;
//       for (let i = 5; i <= 15; i++) {
//         const cellValue = parseFloat(cells[i].textContent.trim().replace('$', '').replace(/,/g, ''));
//         if (!isNaN(cellValue)) {
//           rowTotal += cellValue;
//         }
//       }
  
//     console.log(rowTotal);
//     //   if (row.classList.contains('EOC-total-row')) {
//     //     cells[17].textContent = rowTotal !== 0 ? '$' + rowTotal.toLocaleString() : '';
//     //   } else {
//     cells[15].textContent = rowTotal !== 0 ? rowTotal.toLocaleString() : '';
   
//     //   }
//     });
//   }
  