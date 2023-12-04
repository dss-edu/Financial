document.addEventListener("DOMContentLoaded", function () {
  var modal = document.getElementById("myModal");

  var modalTableBody = document.getElementById("modal-table-body");
  var mdfooter = document.getElementById("myModalFooter");

  function formatDateToYYYYMMDD(dateString) {
      // Create a new Date object
      const date = new Date(dateString);

      // Get the year, month, and day
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0'); // +1 because months are 0-indexed
      const day = String(date.getDate()).padStart(2, '0');

      // Format the date as YYYY-MM-DD
      return `${year}-${month}-${day}`;
  }

  function populateModal(data) {
    var existingDataTable = $("#balancesheet-data-table1").DataTable();

    if (existingDataTable) {
      existingDataTable.destroy();
    }

    modalTableBody.innerHTML = "";
    mdfooter.innerHTML = "";
    

    data.gl_data.forEach(function (row) {
      var newRow = document.createElement("tr");


      if (ascender == 'True'){
      newRow.innerHTML = `
            <td class="text-end">${row.fund}</td>
            <td class="text-end">${row.func}</td>
            <td class="text-end">${row.obj}</td>
            <td class="text-end">${row.org}</td>
            <td class="text-end">${row.fscl_yr}</td>
            <td class="text-end">${row.pgm}</td>
            <td class="text-end">${row.projDtl}</td>
            <td class="text-end text-nowrap">${row.AcctDescr}</td>
            <td class="text-end">${row.Number}</td>
            <td class="text-end" style="white-space: nowrap;">${formatDateToYYYYMMDD(row.Date)}</td>
            <td class="text-end">${row.AcctPer}</td>
            <td class="text-end">${row.Real}</td>
            <td class="text-end">${row.Expend}</td>
            <td class="text-end">${row.Bal}</td>
            <td class="text-end" style="white-space: nowrap;">${row.WorkDescr}</td>
            <td class="text-end">${row.Type}</td>
          `;
      }      else{
        newRow.innerHTML = `
<td class="text-end">${row.fund}</td>
<td class="text-end">${row.T}</td>
<td class="text-end">${row.func}</td>
<td class="text-end">${row.obj}</td>
<td class="text-end">${row.sobj}</td>
<td class="text-end">${row.org}</td>
<td class="text-end">${row.fscl_yr}</td>
<td class="text-end">${row.PI}</td>
<td class="text-end">${row.LOC}</td>
<td class="text-end">${formatDateToYYYYMMDD(row.Date)}</td>
<td class="text-end">${row.AcctPer}</td>
<td class="text-end">${row.Source}</td>
<td class="text-end">${row.Subsource}</td>
<td class="text-end">${row.Batch}</td>
<td class="text-end">${row.Vendor}</td>
<td class="text-end">${row.TransactionDescr}</td>
<td class="text-end">${row.InvoiceDate}</td>
<td class="text-end">${row.CheckNumber}</td>
<td class="text-end">${row.CheckDate}</td>
<td class="text-end">${row.Amount}</td>

      `;
  }
      // newRow.innerHTML = `
      //   <td class="right-align-td">${row.fund}</td>
      //   <td class="right-align-td">${row.func}</td>
      //   <td class="right-align-td">${row.obj}</td>
      //   <td class="right-align-td">${row.org}</td>
      //   <td class="right-align-td">${row.fscl_yr}</td>
      //   <td class="right-align-td">${row.pgm}</td>
      //   <td class="right-align-td">${row.projDtl}</td>
      //   <td class="right-align-td">${row.AcctDescr}</td>
      //   <td class="right-align-td">${row.Number}</td>
      //   <td class="right-align-td" style="white-space: nowrap;">${row.Date}</td>
      //   <td class="right-align-td">${row.AcctPer}</td>
      //   <td class="right-align-td">${row.Real}</td>
      //   <td class="right-align-td">${row.Expend}</td>
      //   <td class="right-align-td">${row.Bal}</td>
      //   <td class="right-align-td" style="white-space: nowrap;">${row.WorkDescr}</td>
      //   <td class="right-align-td">${row.Type}</td>
      // `;
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

    $("#balancesheet-data-table1").DataTable({
      paging: false,
      searching: true,
    });
  }
  // REVENUE TOTAL

  function fetchDataAndPopulateModal(fund, obj, yr, school, year,url) {
    fetch(`/viewgl/${fund}/${obj}/${yr}/${school}/${year}/${url}`)
      .then(function (response) {
        return response.json();
      })
      .then(function (data) {
        if (data.status === "success") {
          $("#spinner-modal").modal("hide");
          populateModal(data.data, data.total_bal);

          modal.style.display = "block";
        } else {
          console.error(data.message);
        }
      })
      .catch(function (error) {
        console.error("Error fetching data:", error);
      });
  }

  // Add click event listeners to all "a" elements inside the table
  var viewGLLinks = document.querySelectorAll(".viewgl-link");
  viewGLLinks.forEach(function (link) {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      var fund = link.dataset.fund;
      var obj = link.dataset.obj;
      var yr = link.dataset.yr;
      console.log(url)
      fetchDataAndPopulateModal(fund, obj, yr, school, year , url);
    });
  });

  // for Year to Date
  function parseTotal(value){
    if (!value){
      return 0
    }
    let multiplier = 1
    if (value && value.includes('(')) {
      multiplier = -1
    }
    const numbersOnly = value.replace(/[^\d.]/g, '')
    try{
      const numericValue = parseFloat(numbersOnly, 10)
      return numericValue * multiplier
    }
    catch(error) {
      return 0
    }

  }

  async function fetchDataAndPopulateModalYTD({fund, func, obj, school, year, url, endpointName}) {
    // TODO do something about this
    // get only months that are viewable in the table
    const yr = ['09','10']
    const ytdData = {gl_data: [], total_bal: 0}

    for (const month of yr) {
      const endpoints = {
        'viewgl': `/viewgl/${fund}/${obj}/${month}/${school}/${year}/${url}`,
        'viewglfunc': `/viewglfunc/${func}/${month}/${school}/${year}/${url}`,
        'viewglexpense': `/viewglexpense/${obj}/${month}/${school}/${year}/${url}`,
        'viewgldna': `/viewgldna/${func}/${month}/${school}/${year}/${url}`,
      }
      try {
        const response = await fetch(endpoints[endpointName])
        if (response.status !== 200){
          return
        }

        const data = await response.json()
        ytdData.gl_data = ytdData.gl_data.concat(data.data.gl_data)
        ytdData.total_bal = ytdData.total_bal + parseTotal(data.data.total_bal)
      } catch(error){
        console.log(error)
      }
    }

      if (ytdData.total_bal < 0) {
        let formattedTotal = ytdData.total_bal.toLocaleString("en-US", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        })
        ytdData.total_bal = '(' + formattedTotal.replace('-', '') + ')'
      }
      else {
        ytdData.total_bal = ytdData.total_bal.toLocaleString("en-US", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        })
      }

      $("#spinner-modal").modal("hide");
      populateModal(ytdData);
      modal.style.display = "block";
  }

  // Add click event listeners to all year to date column links
  const viewGLYearToDate = document.querySelectorAll(".viewgl-link-ytd")
  viewGLYearToDate.forEach((link) => {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      const args = {
        fund: link.dataset.fund,
        obj: link.dataset.obj,
        school: school,
        year: year,
        url: url,
        endpointName: 'viewgl'
      }
      console.log(args)
      // var fund = link.dataset.fund;
      // var obj = link.dataset.obj;
      fetchDataAndPopulateModalYTD(args);
    });
  })

  // FIRST TOTAL
  function fetchDataAndPopulateModal2(func, yr, school, year, url) {
    fetch(`/viewglfunc/${func}/${yr}/${school}/${year}/${url}`)
      .then(function (response) {
        return response.json();
      })
      .then(function (data) {
        if (data.status === "success") {
          populateModal(data.data, data.total_bal);

          modal.style.display = "block";
        } else {
          console.error(data.message);
        }
      })
      .catch(function (error) {
        console.error("Error fetching data:", error);
      });
  }

  var viewGLLinks2 = document.querySelectorAll(".viewglfunc-link");
  viewGLLinks2.forEach(function (link) {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      var func = link.dataset.func;

      var yr = link.dataset.yr;
      fetchDataAndPopulateModal2(func, yr, school, year , url);
    });
  });

  // Add click event listeners to all year to date column links in first total
  const viewGLFuncYearToDate = document.querySelectorAll(".viewglfunc-link-ytd")
  viewGLFuncYearToDate.forEach((link) => {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      const args = {
        func: link.dataset.func,
        school: school,
        year: year,
        url: url,
        endpointName: 'viewglfunc'
      }

      // var yr = link.dataset.yr;
      fetchDataAndPopulateModalYTD(args);
    });
  })

  //FOR EXPENSE BY Object
  function fetchDataAndPopulateModal3(func, yr, school , year, url) {
    fetch(`/viewglexpense/${func}/${yr}/${school}/${year}/${url}`)
      .then(function (response) {
        return response.json();
      })
      .then(function (data) {
        console.log(data);
        if (data.status === "success") {
          // Populate the modal with the fetched data
          populateModal(data.data, data.total_bal);
          // Show the modal
          modal.style.display = "block";
        } else {
          console.error(data.message);
        }
      })
      .catch(function (error) {
        console.error("Error fetching data:", error);
      });
  }

  // Add click event listeners to all "a" elements inside the table
  var viewGLexpenseLinks3 = document.querySelectorAll(".viewglexpense-link");
  viewGLexpenseLinks3.forEach(function (link) {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      var obj = link.dataset.obj;
      var yr = link.dataset.yr;
      fetchDataAndPopulateModal3(obj, yr, school, year , url);
    });
  });

  const viewGLExpenseYearToDate = document.querySelectorAll(".viewglexpense-link-ytd")
  viewGLExpenseYearToDate.forEach((link) => {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      const args = {
        obj: link.dataset.obj,
        school: school,
        year: year,
        url: url,
        endpointName: 'viewglexpense'
      }

      fetchDataAndPopulateModalYTD(args);
    });
  })

  //FOR DnA

  function fetchDataAndPopulateModal4(func, yr, school, year , url) {
    fetch(`/viewgldna/${func}/${yr}/${school}/${year}/${url}`)
      .then(function (response) {
        return response.json();
      })
      .then(function (data) {
        if (data.status === "success") {
          populateModal(data.data, data.total_bal);

          modal.style.display = "block";
        } else {
          console.error(data.message);
        }
      })
      .catch(function (error) {
        console.error("Error fetching data:", error);
      });
  }

  var viewGLLinks4 = document.querySelectorAll(".viewgldna-link");
  viewGLLinks4.forEach(function (link) {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      var func = link.dataset.func;

      var yr = link.dataset.yr;
      fetchDataAndPopulateModal4(func, yr, school, year, url);
    });
  });

  const viewGLDnaYearToDate = document.querySelectorAll(".viewgldna-link-ytd")
  viewGLDnaYearToDate.forEach((link) => {
    link.addEventListener("click", function (event) {
      $("#spinner-modal").modal("show");
      event.preventDefault();
      const args = {
        func: link.dataset.func,
        school: school,
        year: year,
        url: url,
        endpointName: 'viewgldna'
      }

      fetchDataAndPopulateModalYTD(args);
    });
  })

  //////////////////////////// For the totals of each section /////////////////////////////
  // Local Revenue YTD Total
  const localRevenueYtd = document.getElementById('local-revenue-ytd-total')
  localRevenueYtd.addEventListener('click', () => viewglAll({classes:'.local-revenue-row'}))

  // local revenue total for each month
  const localTotals = document.querySelectorAll('.local-total')
  for(let i=0; i < localTotals.length; i++){
    const month = localTotals[i].dataset.yr
    localTotals[i].addEventListener('click', () => viewglAll({classes:'.local-revenue-row', yr:month}))

  }

  // State Program Revenue YTD Total
  const stateRevenueYtd = document.getElementById('state-revenue-ytd-total')
  stateRevenueYtd.addEventListener('click', () => viewglAll({classes: '.spr-row'}))

  // state program revenue total for each month
  const sprTotals = document.querySelectorAll('.spr-total')
  for(let i=0; i < sprTotals.length; i++){
    const month = sprTotals[i].dataset.yr
    sprTotals[i].addEventListener('click', () => viewglAll({classes:'.spr-row', yr:month}))
  }

  // Federal Revenuew YTD Total
  const federalRevenueYtd = document.getElementById('federal-revenue-ytd-total')
  federalRevenueYtd.addEventListener('click', () => viewglAll({classes:'.fpr-row'}))

  // federal revenue total for each month
  const fprTotals = document.querySelectorAll('.fpr-total')
  for(let i=0; i < fprTotals.length; i++){
    const month = fprTotals[i].dataset.yr
    fprTotals[i].addEventListener('click', () => viewglAll({classes:'.fpr-row', yr:month}))
  }

  // Revenue YTD Total
  const revenueYtd = document.getElementById('all-revenue-ytd-total')
  revenueYtd.addEventListener('click', () => viewglAll({classes:['.local-revenue-row', '.spr-row', '.fpr-row']}))

  // revenue total for each month
  const revenueTotals = document.querySelectorAll('.total-revenue')
  for(let i=0; i < revenueTotals.length; i++){
    const month = revenueTotals[i].dataset.yr
    revenueTotals[i].addEventListener('click', () => viewglAll({classes:['.local-revenue-row', '.spr-row', '.fpr-row'], yr:month}))
  }

  async function viewglAll({classes, yr}){
    $("#spinner-modal").modal("show");
    let rows
    if (typeof classes === 'string' ){
      rows = document.querySelectorAll(classes)
    }
    else {
      rows = document.querySelectorAll(classes.join(', '))
    }

    const data = []

    for (const element of rows) {
      const aTag = element.querySelector('.viewgl-link')
      data.push({
        fund: aTag.dataset.fund,
        obj: aTag.dataset.obj,
      })
    }

    let fetchString
    if(yr){
      fetchString = `/viewgl-all/${school}/${year}/${url}/${yr}`
    } else {
      fetchString = `/viewgl-all/${school}/${year}/${url}/`
    }

    const csrftoken = document.querySelector('input[name=csrfmiddlewaretoken]').getAttribute('value')
    fetch(fetchString, {
      method: "POST",
      mode: "same-origin",
      headers: {
        "Content-Type": "application/json",
        "X-CSRFToken": csrftoken,
      },
      body: JSON.stringify(data)
    }).then((response) => {
        return response.json()
    }).then(data => {
      if (data.status === 'success'){
        $("#spinner-modal").modal("hide");
        populateModal(data.data)
        modal.style.display = "block";
      }
      else {
        $("#spinner-modal").modal("hide");
        console.log(data.status)
      }
    }).catch(error => {
      console.log(error)
    })
  }

  const totalYtdTotal = document.getElementById('total-ytd-total')
  totalYtdTotal.addEventListener('click', () => viewglFuncAll({classes:'.total-row1'}))

  const firstTotals = document.querySelectorAll('.first-total')
  for(let i=0; i < firstTotals.length; i++){
    const month = firstTotals[i].dataset.yr
    firstTotals[i].addEventListener('click', () => viewglFuncAll({classes:'.total-row1',yr:month}))

  }

  async function viewglFuncAll({classes, yr}){
    $("#spinner-modal").modal("show");
    let rows
    if (typeof classes === 'string' ){
      rows = document.querySelectorAll(classes)
    }
    else {
      rows = document.querySelectorAll(classes.join(', '))
    }

    const data = []

    for (const element of rows) {
      const aTag = element.querySelector('.viewglfunc-link')
      data.push(aTag.dataset.func)
    }

    let fetchString
    if(yr){
      fetchString = `/viewglfunc-all/${school}/${year}/${url}/${yr}/`
    } else {
      fetchString =  `/viewglfunc-all/${school}/${year}/${url}/`
    }

    const csrftoken = document.querySelector('input[name=csrfmiddlewaretoken]').getAttribute('value')
    fetch(fetchString, {
      method: "POST",
      mode: "same-origin",
      headers: {
        "Content-Type": "application/json",
        "X-CSRFToken": csrftoken,
      },
      body: JSON.stringify(data)
    }).then((response) => {
        return response.json()
    }).then(data => {
      if (data.status === 'success'){
        $("#spinner-modal").modal("hide");
        populateModal(data.data)
        modal.style.display = "block";
      }
      else {
        $("#spinner-modal").modal("hide");
        console.log(data.status)
        console.log(data.message)
      }
    }).catch(error => {
      console.log(error)
    })
  }

  // Payroll row
  const payrollYtdTotal = document.getElementById('payroll-ytd-total')
  payrollYtdTotal.addEventListener('click', () => viewglExpenseAll({classes:'.payroll-row'}))

  // payroll total per month
  const payrollTotals = document.querySelectorAll('.payroll-total')
  for(let i=0; i<payrollTotals.length; i++){
    const month = payrollTotals[i].dataset.yr
    payrollTotals[i].addEventListener('click', () => viewglExpenseAll({classes:'.payroll-row', yr:month}))

  }

  // Professional and Contract Services
  const professionalYtdTotal = document.getElementById('professional-ytd-total')
  professionalYtdTotal.addEventListener('click', () => viewglExpenseAll({classes:'.PCS-row'}))

  // professional and contract services total per month
  const proTotals = document.querySelectorAll('.professional-total')
  for(let i=0; i<proTotals.length; i++){
    const month = proTotals[i].dataset.yr
    proTotals[i].addEventListener('click', () => viewglExpenseAll({classes:'.PCS-row', yr:month}))

  }

  // Supplies and Materials
  const suppliesYtdTotal = document.getElementById('supplies-ytd-total')
  suppliesYtdTotal.addEventListener('click', () => viewglExpenseAll({classes:'.sm-row'}))

  //  supplies and materials total per month
  const suppliesTotals = document.querySelectorAll('.supplies-total')
  for(let i=0; i<suppliesTotals.length; i++){
    const month = suppliesTotals[i].dataset.yr
    suppliesTotals[i].addEventListener('click', () => viewglExpenseAll({classes:'.sm-row', yr:month}))
  }

  // Other operating costs
  const otherYtdTotal = document.getElementById('other-ytd-total')
  otherYtdTotal.addEventListener('click', () => viewglExpenseAll({classes:'.ooe-row'}))

  //  other operating cost total per month
  const otherTotals = document.querySelectorAll('.other-total')
  for(let i=0; i<otherTotals.length; i++){
    const month = otherTotals[i].dataset.yr
    otherTotals[i].addEventListener('click', () => viewglExpenseAll({classes:'.ooe-row', yr:month}))
  }

  // Debt Services
  const debtYtdTotal = document.getElementById('debt-ytd-total')
  debtYtdTotal.addEventListener('click', () => viewglExpenseAll({classes:'.expense-row'}))

  //  debt services total per month
  const debtTotals = document.querySelectorAll('.ds-total')
  for(let i=0; i<debtTotals.length; i++){
    const month = debtTotals[i].dataset.yr
    debtTotals[i].addEventListener('click', () => viewglExpenseAll({classes:'.expense-row', yr:month}))
  }

  // Expense all 
  const EocYtdTotal = document.getElementById('EOC-ytd-total')
  EocYtdTotal.addEventListener('click', () => {viewglExpenseAll({classes:['.expense-row', '.ooe-row', '.sm-row', '.PCS-row', '.payroll-row']})})

  const EocTotals = document.querySelectorAll('.EOC-total')
  for (let i=0; i<EocTotals.length; i++){
    const month = EocTotals[i].dataset.yr
    EocTotals[i].addEventListener('click', () => {viewglExpenseAll({classes:['.expense-row', '.ooe-row', '.sm-row', '.PCS-row', '.payroll-row'], yr:month})})

  }

  async function viewglExpenseAll({classes, yr}){
    $("#spinner-modal").modal("show");
    let rows
    if (typeof classes === 'string' ){
      rows = document.querySelectorAll(classes)
    }
    else {
      rows = document.querySelectorAll(classes.join(', '))
    }

    const data = []

    for (const element of rows) {
      const aTag = element.querySelector('.viewglexpense-link')
      data.push(aTag.dataset.obj)
    }

    let fetchString
    if(yr){
      fetchString =`/viewglexpense-all/${school}/${year}/${url}/${yr}/`
    } else {
      fetchString =`/viewglexpense-all/${school}/${year}/${url}/`
    }

    const csrftoken = document.querySelector('input[name=csrfmiddlewaretoken]').getAttribute('value')
    fetch(fetchString, {
      method: "POST",
      mode: "same-origin",
      headers: {
        "Content-Type": "application/json",
        "X-CSRFToken": csrftoken,
      },
      body: JSON.stringify(data)
    }).then((response) => {
        return response.json()
    }).then(data => {
      if (data.status === 'success'){
        $("#spinner-modal").modal("hide");
        populateModal(data.data)
        modal.style.display = "block";
      }
      else {
        $("#spinner-modal").modal("hide");
        console.log(data.status)
        console.log(data.message)
      }
    }).catch(error => {
      console.log(error)
    })
  }


  // Close the modal when the close button is clicked
  var closeButton = modal.querySelector(".modal-footer button");
  closeButton.addEventListener("click", function () {
    modal.style.display = "none";
    $("#spinner-modal").modal("hide");
  });

  // search box in modal
  $('#myModal').on('shown.bs.modal', function () {
    $('#modalSearch').trigger('focus')
  })
});
