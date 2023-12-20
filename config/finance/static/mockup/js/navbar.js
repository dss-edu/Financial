document.addEventListener("DOMContentLoaded", function() {
    var dropdownSearch = document.getElementById('dropdownSearch');
    dropdownSearch.addEventListener('keyup', function(event) {
        var searchQuery = event.target.value.toLowerCase();
        var dropdownItems = document.querySelectorAll('.dropdown-menu .dropdown-item');

        dropdownItems.forEach(function(item) {
            var text = item.textContent.toLowerCase();
            var match = text.indexOf(searchQuery) > -1;
            item.style.display = match ? '' : 'none';
        });
    });
});
