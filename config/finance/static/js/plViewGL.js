document.addEventListener("DOMContentLoaded", function () {
  var modal = document.getElementById("myModal");

  var modalTableBody = document.getElementById("modal-table-body");
  var mdfooter = document.getElementById("myModalFooter");

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
            <td class="text-end" style="white-space: nowrap;">${row.Date}</td>
            <td class="text-end">${row.AcctPer}</td>
            <td class="text-end">${row.Real}</td>
            <td class="text-end">${row.Expend}</td>
            <td class="text-end">${row.Bal}</td>
            <td class="text-end" style="white-space: nowrap;">${row.WorkDescr}</td>
            <td class="text-end">${row.Type}</td>
          `;
      }      else{
        newRow.innerHTML = `
        <td class="text-center">${row.fund}</td>
        <td class="text-center">${row.func}</td>
        <td class="text-center">${row.obj}</td>
        <td class="text-center">${row.org}</td>
        <td class="text-center">${row.fscl_yr}</td>
  

   
        <td class="text-center" style="white-space: nowrap;">${row.Date}</td>
        <td class="text-center">${row.AcctPer}</td>
        <td class="text-center">${row.Real}</td>

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
