$(document).ready(function() {

  function hideEmptyColumns(){
    const tableBody = $('#data-table').find("tbody")
    const rows = tableBody.find("tr")

    let ignoreList = [] // contains indexes that have data
    let blankList = [] // contains every index


    // get the indexes with data and save it to the ignoreList
    rows.each(function(rIndex){
      const row = $(this)
      const cols = row.find("td")
      cols.each(function(cIndex){
        const col = $(this)
        if (ignoreList.includes(cIndex)){
          return
        }
        if (col.text()){
          ignoreList.push(cIndex)
        }
        if (!blankList.includes(cIndex)){
          blankList.push(cIndex)
        }
      })
    })


    // if the td is in the ignoreList 
    rows.each(function(rIndex){
      const row = $(this)
      const cols = row.find("td")

      // hide the td not in ignorelist
      cols.each(function(cIndex){
        const col = $(this)
        if (!ignoreList.includes(cIndex)){
          col.hide()
        }
      })
    })

    //for the thead
    const tableHead = $('#data-table').find('thead')
    const headers = tableHead.find("tr")

    const titles = headers.eq(2).find("th")

    titles.each(function(index){
      const title = $(this)
      if (!ignoreList.includes(index)){
        title.hide()
      }
    })
  }
});
