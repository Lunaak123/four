function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

const sheetName = getQueryParam('sheetName');
const fileUrl = getQueryParam('fileUrl');

(async () => {
    if (!fileUrl || !sheetName) {
        alert("Invalid sheet data.");
        return;
    }

    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        
        // Populate the dropdown with available sheets
        const sheetSelect = document.getElementById('sheet-select');
        workbook.SheetNames.forEach((name) => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        // Function to display the chosen sheet
        const displaySheet = (chosenSheet) => {
            const sheet = workbook.Sheets[chosenSheet];
            if (sheet) {
                const html = XLSX.utils.sheet_to_html(sheet);
                document.getElementById('sheet-content').innerHTML = html;
            } else {
                alert("Sheet not found.");
            }
        };

        // Load initial sheet (from URL params)
        displaySheet(sheetName);

        // Add event listener to the "Load Sheet" button
        document.getElementById('load-sheet').addEventListener('click', () => {
            const chosenSheet = sheetSelect.value;
            displaySheet(chosenSheet);
        });
    } catch (error) {
        console.error("Error loading Excel file:", error);
        alert("Failed to load the Excel sheet. Please check the URL and try again.");
    }
})();
