document.addEventListener("DOMContentLoaded", function () {
  function getCurrentDateFormatted() {
      var today = new Date();
      var year = today.getFullYear();
      var month = today.getMonth() + 1; // January is 0!
      var day = today.getDate();

      // Adding leading zero if month or day is less than 10
      month = month < 10 ? '0' + month : month;
      day = day < 10 ? '0' + day : day;

      return year + '-' + month + '-' + day;
  }

  $('#reset-button').on('click', function(event){
    window.globalDateStart = ""
    window.globalDateEnd = ""
    // const today = new Date();
    // const formattedToday = today.getFullYear() + '-' + 
    //                    ('0' + (today.getMonth() + 1)).slice(-2) + '-' + 
    //                    ('0' + today.getDate()).slice(-2);
    // $('input[name="daterange"]').data('daterangepicker').setStartDate(formattedToday);
    // $('input[name="daterange"]').data('daterangepicker').setEndDate(formattedToday);

    fetchDateRange("","")
  })
  
  $(async function () {
    $('input[name="daterange"]').daterangepicker(
      {
        opens: "left",
      },
      function (start, end, label) {
        console.log(
          "A new date selection was made: " +
            start.format("YYYY-MM-DD") +
            " to " +
            end.format("YYYY-MM-DD")
        );
        window.globalDateStart = start.format("YYYY-MM-DD")
        window.globalDateEnd = end.format("YYYY-MM-DD")
        fetchDateRange( globalDateStart, globalDateEnd )
      }
    );
  });

  async function fetchDateRange(start, end){
    $("#page-load-spinner").modal("show");

    const csrftoken = document.querySelector('input[name=csrfmiddlewaretoken]').getAttribute('value')
    let url
    if (start !== ""){
        url = `/general-ledger-range/${school}/${start}/${end}`
    } else {
        url =  `/general-ledger-range/${school}`
    }

    fetch(url, {
    //   method: "POST",
      mode: "same-origin",
      headers: {
        "Content-Type": "application/json",
        "X-CSRFToken": csrftoken,
      },
    //   body: JSON.stringify(data)
    }).then((response) => {
        return response.json()
    }).then(data => {
        console.log(data)
        replaceTableData(data)
    }).catch(error => {
      console.log(error)
    })
  }
  function replaceTableData(data) {
    // Assuming your data array is structured like this
    // const data = [
    //     { fund: '...', func: '...', /* other properties */ },
    //     // ... more data objects
    // ];

    $('#spinner-modal').modal('show');

    // Initialize your DataTable
    var table = $('#dataTable').DataTable();

    // Clear existing data
    table.clear();

    // Add new data
    data.forEach((row, index) => {
      table.row.add([
        row.fund,
        row.func,
        row.obj,
        row.org,
        row.fscl_yr,
        row.pgm,
        row.edSpan,
        row.projDtl,
        row.AcctDescr,
        row.Number,
        row.Date,
        row.AcctPer,
        row.Est,
        row.Real,
        row.Appr,
        row.Encum,
        row.Expend,
        row.Bal,
        row.WorkDescr,
        row.Type,
        row.Contr,
      ]);
    });

    // Redraw the table
    table.draw();
    
    $("#page-load-spinner").modal("hide");
  }

});