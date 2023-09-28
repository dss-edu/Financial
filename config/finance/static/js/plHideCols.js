$(document).ready(function() {
  hideEmptyRows()

  // hideEmptyColumns()

  function hideEmptyColumns(){
    const table = document.getElementById('data-table')
    const tBody = table.getElementsByTagName('tbody')[0]

    const tRows = tBody.getElementsByTagName('tr')
    const colN = tRows[0].querySelectorAll('td').length
    
    let blankList = new Array(20).fill(true);

    tRows.forEach((row) => {
      const data = row.querySelectorAll('td')
      for (let i = 5; i < colN; i ++){
        if (data[i].textContent.trim() !== ''){
          blankList[i] = false
        }
      }
    })
    // make the first 5 false
    for (let i = 0; i < 5; i++){
      blankList[i] = false
    }
    console.log(blankList)


    // go through every column and remove the blanks
    tRows.forEach((row)=>{
      const data = row.getElementsByTagName('td')
      data.forEach((col, index, arr)=>{
        if (blankList[index]){
          col.remove()
        }
      })
    })

    const tHead = table.getElementsByTagName('thead')[0]
    const headers = tHead.getElementsByTagName('tr')
    const colTitles = headers[2].getElementsByTagName('th')
    console.log(colTitles)

    colTitles.forEach((title,index, arr)=>{
      console.log(title)
      if (blankList[index]){
        title.remove()
      }
    })

  }

  function blankblankblank(){
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

  function hideEmptyRows(){
    const table = document.getElementById('data-table')
    const tBody = table.getElementsByTagName('tbody')[0]

    const rows = tBody.querySelectorAll('tr')

    rows.forEach(function(row){
      const data = row.getElementsByTagName('td')

      // check if all data has content
      let toDelete = true
      for (let i = 3; i < data.length; i++){
        if (data[i].textContent.trim() !== ''){
          toDelete = false
          break
        }
      }
      
      if (toDelete){
        row.remove()
      }
    })
  }
});
