document.addEventListener("DOMContentLoaded", function () {
  
  
    var modal = document.getElementById("myModal2");

     
      var modalTableBody = document.getElementById("modal2-table-body");

    
      function populateModal(data) {
      
        modalTableBody.innerHTML = "";
      
     
        data.glfunc_data.forEach(function (row) {
          var newRow = document.createElement("tr");
          newRow.innerHTML = `
            <td class="right-align-td">${row.fund}</td>
            <td class="right-align-td">${row.func}</td>
            <td class="right-align-td"> ${row.obj}</td>
            <td class="right-align-td">${row.org}</td>
            <td class="right-align-td">${row.fscl_yr}</td>
            <td class="right-align-td">${row.pgm}</td>
            <td class="right-align-td">${row.projDtl}</td>
            <td class="right-align-td">${row.AcctDescr}</td>
            <td class="right-align-td">${row.Number}</td>
            <td class="right-align-td" style="white-space: nowrap;">${row.Date}</td>
            <td class="right-align-td">${row.AcctPer}</td>
            <td class="right-align-td">${row.Real}</td>
            <td class="right-align-td">${row.Expend}</td>
            <td class="right-align-td">${row.Bal}</td>
            <td class="right-align-td" style="white-space: nowrap;">${row.WorkDescr}</td>
            <td class="right-align-td">${row.Type}</td>
            
          `;
          modalTableBody.appendChild(newRow);
        });
        var totalRow = document.createElement("tr");
        totalRow.innerHTML = `
          <td colspan="12" style="text-align: right;">Total Balance</td>
          <td>${data.total_bal}</td>
          <td colspan="3"></td>
        `;
        modalTableBody.appendChild(totalRow);
      }

   
    function fetchDataAndPopulateModal(func, yr) {
      fetch(`/viewglfunc/${func}/${yr}/`)
        .then(function (response) {
          
          return response.json();
        })
        .then(function (data) {
         
          if (data.status === "success") {
            
            populateModal(data.data,data.total_bal);
           
            modal.style.display = "block";
          } else {
            console.error(data.message);
          }
        })
        .catch(function (error) {
          console.error("Error fetching data:", error);
        });
    }

    var viewGLLinks = document.querySelectorAll(".viewglfunc-link");
    viewGLLinks.forEach(function (link) {
      link.addEventListener("click", function (event) {
        event.preventDefault();
        var func = link.dataset.func;
    
        var yr = link.dataset.yr;
        fetchDataAndPopulateModal(func, yr);
      });
    });

    // Close the modal when the close button is clicked
    var closeButton = modal.querySelector(".modal2-footer button");
    closeButton.addEventListener("click", function () {
      modal.style.display = "none";
    });
  });


 