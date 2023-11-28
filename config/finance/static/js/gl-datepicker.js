document.addEventListener("DOMContentLoaded", function () {
  $('#reset-button').on('click', function(event){
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
        fetchDateRange( start.format("YYYY-MM-DD"), end.format("YYYY-MM-DD") )
      }
    );
  });

  async function fetchDateRange(start, end){
    console.log(school)

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
    
  }
  function replaceTableData1(data){
    const tBody = document.querySelector('#dataTable tbody')
    // remove contents
    tBody.innerHTML = ''

    // replace contents with new data
    console.log('replacing...')
    data.forEach((row,index, arr) => {
        const newRow = document.createElement('tr')
        newRow.className = `table-row ${(index + 1) % 2 === 0 ? 'even': 'odd'}`;
        newRow.innerHTML = `
                        <td>${row.fund}</td>
                        <td>${row.func}</td>
                        <td>${row.obj}</td>
                        <td>${row.org}</td>
                        <td>${row.fscl_yr}</td>
                        <td>${row.pgm}</td>
                        <td>${row.edSpan}</td>
                        <td>${row.projDtl}</td>
                        <td>${row.AcctDescr}</td>
                        <td>${row.Number}</td>
                        <td>${row.Date}</td>
                        <td>${row.AcctPer}</td>
                        <td>${row.Est}</td>
                        <td>${row.Real}</td>
                        <td>${row.Appr}</td>
                        <td>${row.Encum}</td>
                        <td>${row.Expend}</td>
                        <td>${row.Bal}</td>
                        <td>${row.WorkDescr}</td>
                        <td>${row.Type}</td>
                        <td>${row.Contr}</td>
                    `
        tBody.appendChild(newRow)
    })
  }



});