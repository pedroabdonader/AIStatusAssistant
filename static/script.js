document.getElementById('notesForm').onsubmit = function(event) {
    event.preventDefault(); // Prevent the default form submission
    showSpinner(); // Show the spinner

    const formData = new FormData(this); // Create FormData object

    fetch('/', {
        method: 'POST',
        body: formData,
    })
    .then(response => {
        if (response.ok) {
            return response.json(); // Get the response as JSON
        }
        throw new Error('Network response was not ok.');
    })
    .then(data => {
        // Create a table to display the DataFrame
        const table = document.getElementById('dataTable');
        table.innerHTML = ''; // Clear previous data
    
        // Use the received column names for headers
        const columnOrder = data.columns; // Get the column names from the response
    
        // Create table headers based on the defined order
        const headerRow = document.createElement('tr');
        columnOrder.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header; // Use the header name
            headerRow.appendChild(th);
        });
        table.appendChild(headerRow); // Append header row to the table
    
        // Populate the table with DataFrame data based on the defined order
        data.data.forEach(row => {
            const tr = document.createElement('tr');
            columnOrder.forEach(header => {
                const td = document.createElement('td');
                td.textContent = row[header] || ''; // Use the key to access the correct data
                tr.appendChild(td);
            });
            table.appendChild(tr);
        });
    
        // Show the modal
        document.getElementById('dataModal').style.display = 'block';
    
        // Clear previous download link if it exists
        const existingLink = document.getElementById('downloadLink');
        if (existingLink) {
            existingLink.remove(); // Remove the existing download link
        }
    
        // Create a new download link for the PowerPoint file
        const downloadLink = document.createElement('a');
        downloadLink.id = 'downloadLink'; // Set an ID for the download link
        downloadLink.href = '/download'; // Link to download the PowerPoint file
        downloadLink.textContent = 'Download PowerPoint Report';
        downloadLink.style.display = 'block'; // Make it a block element
        document.getElementById('modalContent').appendChild(downloadLink);
    
        hideSpinner(); // Hide the spinner after processing
    });
};

function showSpinner() {
    document.getElementById('loadingSpinner').style.display = 'flex'; // Show the spinner
}

function hideSpinner() {
    document.getElementById('loadingSpinner').style.display = 'none'; // Hide the spinner
}

// Close modal functionality
document.getElementById('closeModal').addEventListener('click', function() {
    document.getElementById('dataModal').style.display = 'none';
});
