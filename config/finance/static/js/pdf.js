document.addEventListener('DOMContentLoaded', function (){
    const exportPDFButton = document.getElementById('export-pdf-button')

    const table = {
        'profit-loss' :  'data-table',
        'balance-sheet' :  'data-table',
        'cashflow-statement': 'data-table'
    }

    exportPDFButton.addEventListener('click', function(event){

      $("#page-load-spinner").modal("show");
      event.preventDefault();
      const url = new URL(window.location.href)
      // url.pathname returns /dashboard/advantage
      const page = url.pathname.split('/')[1]
      const tableBox = document.getElementById('table-box')
      const tableOverview = document.getElementById('data-table-overview')
      const originalTable = document.getElementById(table[page]);

    //////////////////////////////////////////// cloned table ///////////////////////////////////////////////////////
      // export table for overview table
      const clonedTable = originalTable.cloneNode(true); // Clone the table and its contents

      // remove toggle button
    //   const toggleButton = clonedTable.querySelector('#toggle-button')
    // //   toggleButton.parentNode.removeChild(toggleButton)
    //   toggleButton.style.display = 'none'

    //   const buttonHeader = clonedTable.querySelector('thead tr.header th.firstcolumn')
    //   buttonHeader.style.display = 'none'

      
    //   const headerCollapsed = clonedTable.querySelector('thead th.collapsed')
    //   headerCollapsed.style.display = 'none'


    //   const buttonColumns = clonedTable.querySelectorAll('tbody td.firstcolumn')
    //   for( btn of buttonColumns) {
    //     btn.style.display = 'none'
    //     // btn.parentNode.removeChild(btn)
    //   }
    //   const tdCollapsed = clonedTable.querySelector('tbody td.collapsed') 
    //   tdCollapsed.style.display = 'none'
    // //   tdCollapsed.parentNode.removeChild(tdCollapsed)

      // styles
    //   const headers = clonedTable.querySelectorAll('thead tr.header th')
    //   for (header of headers) {
    //     const div = header.querySelector('div')
    //     if (div) {
    //         div.style.textAlign = 'right'
    //     }
    //   }

    // Remove ID from the cloned table
    clonedTable.removeAttribute('id');

    // Remove IDs from all cloned child elements
    var clonedElements = clonedTable.getElementsByTagName('*'); // Get all elements
    for (var i = 0; i < clonedElements.length; i++) {
        clonedElements[i].removeAttribute('id');
    }

    tableOverview.appendChild(clonedTable)
    tableOverview.style.display=''

    /////////////////////////////////////////////////////// original table ///////////////////////////////////////////////////////
    // expand all expandable cols/rows in orig table
    const originalExpandButtons = originalTable.querySelectorAll('a.expand-button')

    for (btn of originalExpandButtons){
      btn.click();
    }


    const allExpandButtons = document.querySelectorAll('a.expand-button')
    for (btn of allExpandButtons){
        btn.style.display='none'
    }

      
    const options = {
        margin: 2,
        filename: "exported_table.pdf",
        image: { type: "jpg", quality: 0.6 },
        html2canvas: {
          scale: 1,
          timeout: 0,
          logging: true, 
          // scrollX: 0,
          // scrollY: 0,
        },
        // jsPDF: { unit: "mm", format: [600, 600], orientation: "portrait" },
        jsPDF: { unit: "mm", format: 'a3', orientation: "landscape" },
        // jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' },
        // pagebreak: { mode: ["avoid-all", "css", "legacy"] },
        // pagebreak: { mode: "css", avoid: "tr", after:"#data-table-overview" },
        pagebreak: { mode: "css", avoid: "tr"},
      };

      html2pdf().from(tableBox).set(options).save().then(() => {
        // tableOverview.style.display='none'
        const tableOverview = document.getElementById('data-table-overview')
        tableOverview.innerHTML = ''
        tableOverview.style.display = 'none'

        const originalExpandButtons = originalTable.querySelectorAll('a.expand-button')
        for (btn of originalExpandButtons){
          btn.click();
        }

        const allExpandButtons = document.querySelectorAll('a.expand-button')
        for (btn of allExpandButtons){
            btn.style.display=''
        }

        $("#page-load-spinner").modal("hide");
      });


    })

})