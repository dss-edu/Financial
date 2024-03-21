document.addEventListener("DOMContentLoaded", function () {
  var modal = document.getElementById("myModal");

  var modalTableBody = document.getElementById("modal-table-body");
  var mdfooter = document.getElementById("myModalFooter");

  function formatDateToYYYYMMDD(dateString) {
      // Create a new Date object
      if (!dateString){
        return ''
      }

      const date = new Date(dateString);

      // Get the year, month, and day
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0'); // +1 because months are 0-indexed
      const day = String(date.getDate()).padStart(2, '0');

      // Format the date as YYYY-MM-DD
      return `${year}-${month}-${day}`;
  }

  function formatNumberToComma(numString) {
  return numString.toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
          })
  }

  // <td class="text-end">${row.org}</td>
  // <td class="text-end">${row.fscl_yr}</td>
  // <td class="text-end">${row.pgm}</td>
  // <td class="text-end">${row.projDtl}</td>

  // <td class="text-end px-3">${row.sobj}</td>
  // <td class="text-end px-3">${row.org}</td>

  // <td class="text-start px-3">${row.PI}</td>
  // <td class="text-start px-3">${row.LOC}</td>

  function populateModal(data) {
    console.log(data)
    const existingDataTable = $('#balancesheet-data-table').DataTable()

    if (existingDataTable) {
      existingDataTable.destroy();
    }

    modalTableBody.innerHTML = "";
    mdfooter.innerHTML = "";

    data.glbs_data.forEach(function (row) {
      var newRow = document.createElement("tr");

      if (ascender == "True") {

        newRow.innerHTML = `
                    <td class="text-end  px-3">${row.fund}</td>
                    <td class="text-end  px-3">${row.func}</td>
                    <td class="text-end  px-3">${row.obj}</td>


                      <td class="text-start text-nowrap  px-3"  style=" padding-left:30px !important;">${row.AcctDescr}</td>
                    <td class="text-center  px-3">${row.Number}</td>
                    <td class="text-center  px-3" style="white-space: nowrap;">${formatDateToYYYYMMDD(row.Date)}</td>

                    <td class="text-end  px-3">${ row.AcctPer}</td>
                    <td class="text-end  px-3">${row.Real}</td>
                    <td class="text-end  px-3">${row.Expend}</td>
                    <td class="text-end  px-3">${row.Bal}</td>

                    <td class="text-start  px-3" style="white-space: nowrap;  padding-left:30px !important;">${row.WorkDescr}</td>
                    <td class="text-start  px-3">${row.Type}</td>
                  `;
      } else {
        newRow.innerHTML = `
                <td class="text-end px-3">${row.fund}</td>

                <td class="text-start px-3">${row.T}</td>

                <td class="text-end px-3">${row.func}</td>
                <td class="text-end px-3">${row.obj}</td>

                <td class="text-start text-nowrap px-3">${formatDateToYYYYMMDD(row.Date)}</td>
                <td class="text-start px-3">${row.Source}</td>
                <td class="text-start px-3">${row.Subsource}</td>
                <td class="text-start text-nowrap px-3">${row.Batch}</td>
                <td class="text-start text-nowrap px-3">${row.Vendor}</td>
                <td class="text-start text-nowrap px-3">${row.TransactionDescr}</td>
                <td class="text-start text-nowrap px-3">${formatDateToYYYYMMDD(row.InvoiceDate)}</td>

                <td class="text-end px-3">${row.CheckNumber}</td>

                <td class="text-start text-nowrap px-3">${formatDateToYYYYMMDD(row.CheckDate)}</td>

                <td class="text-end px-3">${formatNumberToComma(row.Amount)}</td>
              `;
      }

      modalTableBody.appendChild(newRow);
    });

    var totalRow = document.createElement("tr");
    totalRow.innerHTML = `
        <td colspan="11"><div style="width:800px"></div></td>
        <td style="text-align: right; font-size:25px"><strong>Total:</strong></td>
        <td id="modal-total-balance" style="font-size:25px"> $ ${data.total_bal}</td>
        <td colspan="4"></td>
        `;
    mdfooter.appendChild(totalRow);

   $("#balancesheet-data-table").DataTable({
      paging: false,
      searching: true,
    });
  }

  // function fetchDataAndPopulateModal(obj, yr, school, year, url) {
  //   fetch(`/viewgl_activitybs/${obj}/${yr}/${school}/${year}/${url}/`)
  //     .then(function (response) {
  //       return response.json();
  //     })
  //     .then(function (data) {
  //       if (data.status === "success") {
  //         $("#spinner-modal").modal("hide");

  //         populateModal(data.data, data.total_bal);

  //         modal.style.display = "block";
  //       } else {
  //       }
  //     })
  //     .catch(function (error) {});
  // }


  function fetchDataAndPopulateModal(obj, yr, school, year, url) {
    data = {
      obj: obj
    }

    const csrftoken = document.querySelector('input[name=csrfmiddlewaretoken]').getAttribute('value')
    fetch(`/viewgl_activitybs/${yr}/${school}/${year}/${url}/`, {
      method: 'POST', // Specify the method
      headers: {
          'Content-Type': 'application/json', // Specify the content type
          "X-CSRFToken": csrftoken,
      },
      body: JSON.stringify(data) // Convert the JavaScript object to a JSON string
    })
    .then(function (response) {
      return response.json();
    })
    .then(function (data) {
      if (data.status === "success") {
        $("#spinner-modal").modal("hide");

        populateModal(data.data, data.total_bal);
        const gltable = $('#balancesheet-data-table').DataTable()
        if(ascender == 'True'){
          gltable.column(7).visible(false);
          gltable.column(8).visible(false);
          
        }


        // modal.style.display = "block";

        $("#myModal").modal("show");
      } else {
      }
    })
    .catch(function (error) {
      console.log(error)
    });
  }

  var viewGLLinks = document.querySelectorAll(".viewgl_activitybs-link");
  viewGLLinks.forEach(function (link) {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();

      var obj = link.dataset.obj ? [link.dataset.obj] : [];
      var yr = link.dataset.yr;

      // code for totals
      if (!obj[0]){
        const section =  Array.from(link.classList).filter(cls => /-total$/.test(cls));
        console.log(section)
        const sectionName = section[0].split("-")[0]
        const sectionElements = document.querySelectorAll(`.${sectionName}-section`)
        console.log(sectionElements)
        sectionElements.forEach((element) => {
          obj.push(element.querySelector('.viewgl_activitybs-link').dataset.obj)
        })
      }
      console.log("this runs")
      fetchDataAndPopulateModal(obj, yr, school, year, url);
    });
  });

  var closeButton = modal.querySelector(".modal-footer button");
  closeButton.addEventListener("click", function () {
    modal.style.display = "none";
  });
});

var columnsCollapsed = true;
function toggleColumns() {
  const dataTable = document.getElementById("data-table");
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
  var lm = last_month_number;

  console.log(lm);
  var columnsToHide2 = document.querySelectorAll(
    "#data-table th:nth-child(" + lm + "), #data-table td:nth-child(" + lm + ")"
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
  var lm = last_month_number;

  var columnsToHide2 = document.querySelectorAll(
    "#data-table th:nth-child(" + lm + "), #data-table td:nth-child(" + lm + ")"
  );

  for (var k = 0; k < columnsToHide2.length; k++) {
    columnsToHide2[k].style.display = "none";
  }
}

function validateBudgetInput(event) {
  const input = event.target;
  const value = input.value;

  const regex = /^[0-9.]*$/;

  if (!regex.test(value)) {
    input.value = value.replace(/[^0-9.]/g, "");
  }
}

function extractNumericValue(content) {
  const match = content.match(/\(([^)]+)\)/);
  if (match) {
    const numericValue = parseFloat(match[1].replace(/[$,]/g, "").trim());
    return isNaN(numericValue) ? 0 : -numericValue;
  } else {
    const numericValue = parseFloat(content.replace(/[$,]/g, "").trim());
    return isNaN(numericValue) ? 0 : numericValue;
  }
}

/* --------------current assets total function ------------------*/
function calculateCurrentAssets() {
  const rows = document.querySelectorAll(".current-assets-row");
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`current-assets-total-${cellId}`);

    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach((row) => {
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
    totalCell.textContent =
      columnTotal !== 0 ? columnTotal.toLocaleString() : "";
  }
}

/* --------------capital assets total function ------------------*/
function calculateCapitalAssets() {
  const rows = document.querySelectorAll(".capital-assets-row");
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`capital-assets-total-${cellId}`);

    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach((row) => {
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
    totalCell.textContent =
      columnTotal !== 0 ? columnTotal.toLocaleString() : "";
  }
}

/* -------------- total  assets function ------------------*/
function calculateTotalAssets() {
  const rows = document.querySelectorAll(
    ".capital-assets-row,.current-assets-row"
  );
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(`total-assets-total-${cellId}`);

    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach((row) => {
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
    totalCell.textContent =
      columnTotal !== 0 ? "$" + columnTotal.toLocaleString() : "";
  }
}

/* --------------total current Liabilities function ------------------*/
function calculateCurrentLiabilities() {
  const rows = document.querySelectorAll(".current-liabilities-row");
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(
      `current-liabilities-total-${cellId}`
    );

    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach((row) => {
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
    const formattedTotal =
      columnTotal < 0
        ? `(${Math.abs(columnTotal).toLocaleString()})`
        : columnTotal.toLocaleString();
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : "";
  }
}

/* --------------total  Liabilities function ------------------*/
function calculateTotalLiabilities() {
  const rows = document.querySelectorAll(
    ".total-liabilities-row,.current-liabilities-row"
  );
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(
      `total-liabilities-total-${cellId}`
    );

    totalCells.push(totalCell);
  }

  let columnTotals = new Array(totalCells.length).fill(0);

  rows.forEach((row) => {
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
    const formattedTotal =
      columnTotal < 0
        ? `(${Math.abs(columnTotal).toLocaleString()})`
        : columnTotal.toLocaleString();
    totalCell.textContent = columnTotal !== 0 ? formattedTotal : "";
  }
}

/* --------------total  Liabilities and nets assets function ------------------*/
function calculateTotalLiabilitiesAndNetAssets() {
  const rows = document.querySelectorAll(
    ".net-assets-row, .total-liabilities-row, .current-liabilities-row"
  );
  const totalCells = [];

  const totalCellIds = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (let i = 0; i < totalCellIds.length; i++) {
    const cellId = totalCellIds[i];
    const totalCell = document.getElementById(
      `liabilitiesandnet-total-${cellId}`
    );

    totalCells.push(totalCell);
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
              populateModal(data.data, data.total_bal);

              modal.style.display = "block";
            } else {
            }
          })
          .catch(function (error) {});
      }

      var viewGLLinks = document.querySelectorAll(".viewgl_activitybs-link");
      viewGLLinks.forEach(function (link) {
        link.addEventListener("click", function (event) {
          event.preventDefault();

          var obj = link.dataset.obj;
          var yr = link.dataset.yr;
          fetchDataAndPopulateModal(obj, yr);
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
        input.value = value.replace(/[^0-9.]/g, "");
      }
    }

    function extractNumericValue(content) {
      const match = content.match(/\(([^)]+)\)/);
      if (match) {
        const numericValue = parseFloat(match[1].replace(/,/g, "").trim());
        return isNaN(numericValue) ? 0 : -numericValue;
      } else {
        const numericValue = parseFloat(content.replace(/,/g, "").trim());
        return isNaN(numericValue) ? 0 : numericValue;
      }
    }

    /* --------------current assets total function ------------------*/
    function calculateCurrentAssets() {
      const rows = document.querySelectorAll(".current-assets-row");
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `current-assets-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        totalCell.textContent =
          columnTotal !== 0 ? columnTotal.toLocaleString() : "";
      }
    }

    /* --------------capital assets total function ------------------*/
    function calculateCapitalAssets() {
      const rows = document.querySelectorAll(".capital-assets-row");
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `capital-assets-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        totalCell.textContent =
          columnTotal !== 0 ? columnTotal.toLocaleString() : "";
      }
    }

    /* -------------- total  assets function ------------------*/
    function calculateTotalAssets() {
      const rows = document.querySelectorAll(
        ".capital-assets-row,.current-assets-row"
      );
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `total-assets-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        totalCell.textContent =
          columnTotal !== 0 ? "$" + columnTotal.toLocaleString() : "";
      }
    }

    /* --------------total current Liabilities function ------------------*/
    function calculateCurrentLiabilities() {
      const rows = document.querySelectorAll(".current-liabilities-row");
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `current-liabilities-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        const formattedTotal =
          columnTotal < 0
            ? `(${Math.abs(columnTotal).toLocaleString()})`
            : columnTotal.toLocaleString();
        totalCell.textContent = columnTotal !== 0 ? formattedTotal : "";
      }
    }

    /* --------------total  Liabilities function ------------------*/
    function calculateTotalLiabilities() {
      const rows = document.querySelectorAll(
        ".total-liabilities-row,.current-liabilities-row"
      );
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `total-liabilities-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        const formattedTotal =
          columnTotal < 0
            ? `(${Math.abs(columnTotal).toLocaleString()})`
            : columnTotal.toLocaleString();
        totalCell.textContent = columnTotal !== 0 ? formattedTotal : "";
      }
    }

    /* --------------total  Liabilities and nets assets function ------------------*/
    function calculateTotalLiabilitiesAndNetAssets() {
      const rows = document.querySelectorAll(
        ".net-assets-row, .total-liabilities-row, .current-liabilities-row"
      );
      const totalCells = [];

      const totalCellIds = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
      ];

      for (let i = 0; i < totalCellIds.length; i++) {
        const cellId = totalCellIds[i];
        const totalCell = document.getElementById(
          `liabilitiesandnet-total-${cellId}`
        );

        totalCells.push(totalCell);
      }

      let columnTotals = new Array(totalCells.length).fill(0);

      rows.forEach((row) => {
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
        totalCell.textContent =
          columnTotal !== 0 ? "$" + columnTotal.toLocaleString() : "";
      }
    }

    function calculateFYTD() {
      const rows = document.querySelectorAll(".balancesheet-activity-row");

      rows.forEach((row) => {
        const cells = row.cells;
        let rowTotal = 0;
        for (let i = 4; i <= 15; i++) {
          const cellValue = extractNumericValue(cells[i].textContent);
          if (!isNaN(cellValue)) {
            rowTotal += cellValue;
          }
        }

        const formattedTotal =
          rowTotal < 0
            ? `(${Math.abs(rowTotal).toLocaleString()})`
            : rowTotal.toLocaleString();
        cells[16].textContent = rowTotal !== 0 ? formattedTotal : ""; //fytd
        cells[17].textContent = rowTotal !== 0 ? formattedTotal : ""; // as of last month
      });
    }

    function calculateFYTD2() {
      const rows = document.querySelectorAll(".liabilities-activity-row");

      rows.forEach((row) => {
        const cells = row.cells;
        let rowTotal = 0;
        for (let i = 4; i <= 15; i++) {
          const cellValue = extractNumericValue(cells[i].textContent);
          if (!isNaN(cellValue)) {
            rowTotal -= cellValue;
          }
        }

        const formattedTotal =
          rowTotal < 0
            ? `(${Math.abs(rowTotal).toLocaleString()})`
            : rowTotal.toLocaleString();
        cells[16].textContent = rowTotal !== 0 ? formattedTotal : ""; //fytd
        cells[17].textContent = rowTotal !== 0 ? formattedTotal : ""; // as of last month
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
      rows.forEach((row) => {
        row.style.display = "none";
      });
    }

    function hideRowsOnLoad() {
      const collapsedRows = document.querySelectorAll(".collapsed");
      collapsedRows.forEach((row) => {
        row.style.display = "none";
      });

      hideRowsByClass(".balancesheet-activity-row");
      hideRowsByClass(".liabilities-activity-row");
    }

    window.addEventListener("DOMContentLoaded", () => {
      hideRowsOnLoad();

      const editableCells = document.querySelectorAll(
        ".current-assets-row td[contenteditable]"
      );

      editableCells.forEach((cell) => {
        cell.addEventListener("input", () => {
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

  rows.forEach((row) => {
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
    totalCell.textContent =
      columnTotal !== 0 ? "$" + columnTotal.toLocaleString() : "";
  }
}

function calculateFYTD() {
  const rows = document.querySelectorAll(".balancesheet-activity-row");

  rows.forEach((row) => {
    const cells = row.cells;
    let rowTotal = 0;
    for (let i = 4; i <= 15; i++) {
      const cellValue = extractNumericValue(cells[i].textContent);
      if (!isNaN(cellValue)) {
        rowTotal += cellValue;
      }
    }

    const formattedTotal =
      rowTotal < 0
        ? `(${Math.abs(rowTotal).toLocaleString()})`
        : rowTotal.toLocaleString();
    cells[16].textContent = rowTotal !== 0 ? formattedTotal : "";
    cells[17].textContent = rowTotal !== 0 ? formattedTotal : "";
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
  rows.forEach((row) => {
    row.style.display = "none";
  });
}

function hideRowsOnLoad() {
  const collapsedRows = document.querySelectorAll(".collapsed");
  collapsedRows.forEach((row) => {
    row.style.display = "none";
  });

  // Hide local revenue rows
  // hideRowsByClass(".balancesheet-activity-row");
  // hideRowsByClass(".liabilities-activity-row");
  // Hide state program revenue rows
}

window.addEventListener("DOMContentLoaded", () => {
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
