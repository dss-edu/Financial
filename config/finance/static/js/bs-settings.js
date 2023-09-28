document.addEventListener('DOMContentLoaded', function(){
  const settingsNav = document.getElementById('settings-nav')

  settingsNav.addEventListener('click', function(event){
    event.preventDefault()
    $('#settings-modal').modal('show')
  })

  $('#save-settings-btn').on('click', async function(event){
    const csrfToken = document.querySelector('[name=csrfmiddlewaretoken]').value
    event.preventDefault()
    console.log(school)


    const data = getSettingsData()
    let proceed = true
    data.forEach((row) => {
      if (row["activity"].trim() === '') {
        alert("There are still activities without tags please select a tag first")
        proceed = false
        return
      }
    })

    if (!proceed) {
      return
    }

    try {
      $("#settings-modal").modal("hide");
      $("#page-load-spinner").modal("show");
      const response = await fetch("/bs/activity-edits/" + school, {
        method: "POST",
        headers: {
            'X-CSRFToken': csrfToken,
            'Content-Type': 'application/json'

        },
        body: JSON.stringify(data)
      })

      if (response.ok){
        $("#page-load-spinner").modal("hide");
        window.location.href = "/balance-sheet/" + school;
      } else {
          $("page-load-spinner").modal("hide");
          throw new Error("Failed to add manual adjustment.");
      }
    } catch (error) {
      // Handle any errors that occurred during the fetch request
      $("page-load-spinner").modal("hide");
      alert("An error occurred while processing your request.");
    }
    

  })

  function getSettingsData(){
    const table = document.getElementById('settings-table')
    const tBody = table.getElementsByTagName('tbody')[0]
    const tRows = tBody.getElementsByTagName('tr')
    // get all the select values
    const selectedValues = []
    $('select[name="missing-activity"]').each(function(){
      const selectValue = $(this).val()
      selectedValues.push(selectValue)
    })

    let data = []
    tRows.forEach((row, index, arr)=>{
      const td = row.getElementsByTagName('td')
      let activity = selectedValues[index].trim()
      let obj =  td[1].textContent.trim()
      let description =  td[2].textContent.trim()
      data.push(
        {
          activity: activity,
          obj: obj,
          description: description,
        }
      )
    })

    return data
  }


})
