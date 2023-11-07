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
          <td style="font-size:25px"> $ ${data.total_bal}</td>
          <td colspan="4"></td>
        `;
    mdfooter.appendChild(totalRow);

    $("#balancesheet-data-table1").DataTable({
      dom: "lrtip",
      paging: false,
    });
  }

  // REVENUE TOTAL

  function fetchDataAndPopulateModal(fund, obj, yr, school) {
    fetch(`/viewgl/${fund}/${obj}/${yr}/${school}`)
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
      fetchDataAndPopulateModal(fund, obj, yr, school);
    });
  });

  // FIRST TOTAL
  function fetchDataAndPopulateModal2(func, yr, school) {
    fetch(`/viewglfunc/${func}/${yr}/${school}`)
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
      fetchDataAndPopulateModal2(func, yr, school);
    });
  });

  //FOR EXPENSE BY Object
  function fetchDataAndPopulateModal3(func, yr, school) {
    fetch(`/viewglexpense/${func}/${yr}/${school}`)
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
      fetchDataAndPopulateModal3(obj, yr, school);
    });
  });

  // Close the modal when the close button is clicked
  var closeButton = modal.querySelector(".modal-footer button");
  closeButton.addEventListener("click", function () {
    modal.style.display = "none";
    $("#spinner-modal").modal("hide");
  });
});
