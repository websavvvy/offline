// JavaScript File for Excel Upload and Processing
let uploadedFile = null; // Variable to store the uploaded Excel file

// Handle the upload button click to trigger the file input
document.getElementById('uploadBtn').addEventListener('click', () => {
    document.getElementById('excelFileInput').click();
});

// Handle Excel file upload
document.getElementById('excelFileInput').addEventListener('change', (event) => {
    uploadedFile = event.target.files[0];
    const fileNameDisplay = document.getElementById('fileName');

    if (uploadedFile) {
        fileNameDisplay.textContent = `Uploaded File: ${uploadedFile.name}`;
        showNotification(`File "${uploadedFile.name}" uploaded successfully.`, "success");
    } else {
        fileNameDisplay.textContent = "";
        showNotification("No file selected. Please upload an Excel file.", "error");
    }
});

// Handle Excel processing
document.getElementById('processBtn').addEventListener('click', () => {
    const output = document.getElementById('output');
    const progress = document.getElementById('progress');
    const progressBar = document.querySelector('.progress-bar');

    if (!uploadedFile) {
        showNotification("Please upload an Excel file before processing.", "error");
        return;
    }

    // Clear the output field before processing
    output.textContent = "";

    // Show progress bar and reset
    progress.classList.remove("hidden");
    progressBar.style.width = "0%";

    // We use requestAnimationFrame to handle progress bar smoothly
    let progressInterval;
    const updateProgress = () => {
        const currentWidth = parseInt(progressBar.style.width);
        if (currentWidth < 90) {
            progressBar.style.width = `${currentWidth + 5}%`;
            progressInterval = requestAnimationFrame(updateProgress);
        }
    };

    requestAnimationFrame(updateProgress); // Start progress bar animation

    setTimeout(() => {
        const reader = new FileReader();

        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Log all sheet names for better diagnosis
            console.log("Sheet Names:", workbook.SheetNames);

            // Assuming the first sheet is the one to process, or use the correct sheet if needed
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert sheet data to JSON (ensure full data is parsed)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

            // Log the entire data to understand its structure
            console.log("Sheet Data:", jsonData);

            // Extract specific columns by header name
            const headers = jsonData[0]; // First row contains headers
            console.log("Headers from the uploaded file:", headers); // Debugging line

            const requiredColumns = ["Name", "Account", "Last message time"];

            // Normalize headers (trim spaces, convert to lowercase) for comparison
            const normalizedHeaders = headers.map(header => header.trim().toLowerCase());
            console.log("Normalized Headers:", normalizedHeaders); // Debugging line

            const columnIndexes = requiredColumns.map(col => {
                // Normalize required columns for comparison
                return normalizedHeaders.findIndex(header => header === col.trim().toLowerCase());
            });

            // If required columns are missing, show error message
            if (columnIndexes.some(index => index === -1)) {
                const missingColumns = requiredColumns.filter((col, index) => columnIndexes[index] === -1);
                showNotification(`Missing columns: ${missingColumns.join(", ")}`, "error");
                progress.classList.add("hidden");
                cancelAnimationFrame(progressInterval); // Stop progress bar animation
                return;
            }

            const filteredData = jsonData.map((row, rowIndex) => {
                if (rowIndex === 0) return requiredColumns; // Keep header row
                // Convert the "Last message time" to relative time format
                const lastMessageTime = row[columnIndexes[2]] || "";
                const relativeTime = convertToRelativeTime(lastMessageTime);
                return { Name: row[columnIndexes[0]], Account: row[columnIndexes[1]], LastMessageTime: relativeTime };
            }).slice(1); // Remove the header row

            // Display the result after processing is complete
            output.textContent = JSON.stringify(filteredData, null, 2);

            // Update progress bar to 100% and hide progress bar after 0.5 seconds
            progressBar.style.width = "100%";
            setTimeout(() => {
                progress.classList.add("hidden");
                cancelAnimationFrame(progressInterval); // Stop progress bar animation
                showNotification("Excel file processed successfully.", "success");

                // Group data by Account and prepare table data
                const groupedData = groupByAccount(filteredData);

                // Sort by last message received in descending order (longest offline duration first)
                const sortedData = sortByOfflineDuration(groupedData);

                // Display sorted data as a table
                displayAsTable(sortedData);
            }, 500); // Hide progress bar after 0.5s
        };

        reader.onerror = function () {
            showNotification("Could not read the file.", "error");
            progress.classList.add("hidden");
            cancelAnimationFrame(progressInterval); // Stop progress bar animation
        };

        reader.readAsArrayBuffer(uploadedFile);
    }, 500); // Simulate processing delay
});

// Helper function to show notifications
function showNotification(message, type = "success") {
    const notification = document.getElementById("notification");
    notification.textContent = message;
    notification.className = `notification ${type}`;
    setTimeout(() => {
        notification.className = "notification hidden"; // Hide after 3 seconds
    }, 3000);
}

// Convert date to relative time (e.g., "328 days 13 h 20 min ago" or "6 h 20 min ago")
function convertToRelativeTime(dateString) {
    // Ensure the string is valid and in the correct format
    if (!dateString || typeof dateString !== "string") {
        return "Invalid Date";
    }

    const dateParts = dateString.split(' '); // Split the date and time (e.g., "06.02.2024 10:30:04")
    if (dateParts.length !== 2) {
        return "Invalid Date";
    }

    const dateArray = dateParts[0].split('.'); // Split the date by dot (day, month, year)
    const timeArray = dateParts[1].split(':'); // Split the time by colon (hour, minute, second)

    // Ensure the parts are valid
    if (dateArray.length !== 3 || timeArray.length !== 3) {
        return "Invalid Date";
    }

    // Rebuild the date in a format JavaScript can understand: "YYYY-MM-DD HH:mm:ss"
    const formattedDate = `${dateArray[2]}-${dateArray[1]}-${dateArray[0]} ${timeArray[0]}:${timeArray[1]}:${timeArray[2]}`;
    
    // Create a new Date object
    const date = new Date(formattedDate);
    const now = new Date();
    const diffInMs = now - date;

    if (isNaN(date.getTime())) {
        // If the date is invalid, return an error message
        return "Invalid Date";
    }

    const days = Math.floor(diffInMs / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diffInMs % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((diffInMs % (1000 * 60 * 60)) / (1000 * 60));

    // If it's less than a day, display hours and minutes
    if (days === 0) {
        return `${hours} h ${minutes} min ago`;
    }

    // Otherwise, display days, hours, and minutes
    return `${days} days ${hours} h ${minutes} min ago`;
}

// Helper function to group data by account (Client)
function groupByAccount(data) {
    return data.reduce((acc, row) => {
        const account = row.Account;
        const vehicle = row.Name;
        const lastMessageTime = row.LastMessageTime;

        if (!acc[account]) {
            acc[account] = [];
        }

        acc[account].push({ vehicle, lastMessageTime });
        return acc;
    }, {});
}

// Helper function to sort vehicles by offline duration (days, hours, minutes)
function sortByOfflineDuration(groupedData) {
    const accountData = Object.keys(groupedData).map(account => {
        const vehicles = groupedData[account].map(vehicle => {
            return { vehicle: vehicle.vehicle, lastMessageTime: vehicle.lastMessageTime };
        });

        // Sort vehicles by the offline duration (longest first)
        vehicles.sort((a, b) => {
            const diffA = convertToDuration(a.lastMessageTime);
            const diffB = convertToDuration(b.lastMessageTime);
            return diffB - diffA; // Sort in descending order
        });

        return { account, vehicles };
    });

    // Sort accounts by the highest offline duration of their vehicles
    accountData.sort((a, b) => {
        const diffA = convertToDuration(a.vehicles[0].lastMessageTime);
        const diffB = convertToDuration(b.vehicles[0].lastMessageTime);
        return diffB - diffA; // Sort in descending order
    });

    return accountData;
}

// Helper function to convert relative time (e.g., "days h min ago") to a duration in minutes
function convertToDuration(relativeTime) {
    const daysRegex = /(\d+)\s*days/;
    const hoursRegex = /(\d+)\s*h/;
    const minutesRegex = /(\d+)\s*min/;

    const daysMatch = relativeTime.match(daysRegex);
    const hoursMatch = relativeTime.match(hoursRegex);
    const minutesMatch = relativeTime.match(minutesRegex);

    const days = daysMatch ? parseInt(daysMatch[1]) : 0;
    const hours = hoursMatch ? parseInt(hoursMatch[1]) : 0;
    const minutes = minutesMatch ? parseInt(minutesMatch[1]) : 0;

    return days * 1440 + hours * 60 + minutes; // Convert everything to minutes
}

// Helper function to display data as a table
function displayAsTable(sortedData) {
    const table = document.createElement('table');
    const headerRow = document.createElement('tr');

    // Create the table header row
    const headers = ["Account", "Vehicle", "Last Message Time"];
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Add data rows
    sortedData.forEach(data => {
        data.vehicles.forEach(vehicle => {
            const row = document.createElement('tr');
            const accountCell = document.createElement('td');
            accountCell.textContent = data.account;
            row.appendChild(accountCell);

            const vehicleCell = document.createElement('td');
            vehicleCell.textContent = vehicle.vehicle;
            row.appendChild(vehicleCell);

            const lastMessageTimeCell = document.createElement('td');
            lastMessageTimeCell.textContent = vehicle.lastMessageTime;
            row.appendChild(lastMessageTimeCell);

            table.appendChild(row);
        });
    });

    // Append the table to the output section
    const output = document.getElementById('output');
    output.innerHTML = ""; // Clear previous output
    output.appendChild(table);
    // Function to export the displayed table as an Excel file
document.getElementById('downloadBtn').addEventListener('click', () => {
    const table = document.querySelector('table'); // Get the table element
    if (!table) {
        showNotification("No table available to download.", "error");
        return;
    }

    const wb = XLSX.utils.book_new(); // Create a new workbook
    const ws = XLSX.utils.table_to_sheet(table); // Convert the table to a worksheet

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1"); // Append the worksheet to the workbook

    // Trigger the download
    XLSX.writeFile(wb, "output.xlsx");
});

// Update displayAsTable function to show the download button
function displayAsTable(sortedData) {
    const outputContainer = document.getElementById('output');
    outputContainer.innerHTML = ''; // Clear any existing content

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');

    // Create the table header row
    const headers = ["Account", "Vehicle", "Last Message Time"];
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Add data rows
    sortedData.forEach(data => {
        data.vehicles.forEach(vehicle => {
            const row = document.createElement('tr');
            const accountCell = document.createElement('td');
            accountCell.textContent = data.account;
            row.appendChild(accountCell);

            const vehicleCell = document.createElement('td');
            vehicleCell.textContent = vehicle.vehicle;
            row.appendChild(vehicleCell);

            const lastMessageTimeCell = document.createElement('td');
            lastMessageTimeCell.textContent = vehicle.lastMessageTime;
            row.appendChild(lastMessageTimeCell);

            table.appendChild(row);
        });
    });

    outputContainer.appendChild(table); // Add the table to the output container

    // Show the download button
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.classList.remove('hidden');
}
// Update the download button event listener
document.getElementById('downloadBtn').addEventListener('click', () => {
    const table = document.querySelector('table');
    if (!table) {
        showNotification("No table available to download.", "error");
        return;
    }

    // Convert table to array of data (excluding header row)
    const rows = Array.from(table.querySelectorAll('tr')).slice(1);
    const header = Array.from(table.querySelectorAll('tr')[0].cells).map(cell => cell.textContent);
    
    // Group rows by account
    const accountGroups = {};
    rows.forEach(row => {
        const account = row.cells[0].textContent;
        if (!accountGroups[account]) {
            accountGroups[account] = [];
        }
        accountGroups[account].push(row);
    });

    // Get sorted unique accounts
    const accounts = Object.keys(accountGroups).sort();
    
    // Calculate accounts per manager (trying to distribute evenly)
    const totalAccounts = accounts.length;
    const baseSize = Math.floor(totalAccounts / 3);
    const remainder = totalAccounts % 3;
    
    // Distribute accounts to managers
    const managerAccounts = {
        'Ahmed': accounts.slice(0, baseSize + (remainder > 0 ? 1 : 0)),
        'Bilan': accounts.slice(baseSize + (remainder > 0 ? 1 : 0), baseSize * 2 + (remainder > 1 ? 2 : 1)),
        'Mercy': accounts.slice(baseSize * 2 + (remainder > 1 ? 2 : 1))
    };

    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Create sheets for each manager
    Object.entries(managerAccounts).forEach(([manager, managerAccountList]) => {
        // Get all rows for this manager's accounts
        const managerRows = managerAccountList.reduce((acc, account) => {
            return acc.concat(accountGroups[account]);
        }, []);

        // Convert rows to data array
        const sheetData = [
            header,
            ...managerRows.map(row => Array.from(row.cells).map(cell => cell.textContent))
        ];

        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(sheetData);

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, manager);
    });

    // Add summary sheet
    const summaryData = [
        ['Account Manager', 'Number of Accounts', 'Total Vehicles'],
        ...Object.entries(managerAccounts).map(([manager, accounts]) => {
            const totalVehicles = accounts.reduce((sum, account) => {
                return sum + accountGroups[account].length;
            }, 0);
            return [manager, accounts.length, totalVehicles];
        }),
        ['Total', totalAccounts, rows.length]
    ];
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // Trigger the download
    XLSX.writeFile(wb, "account_managers_distribution.xlsx");
    
    showNotification("File downloaded with accounts distributed among Ahmed, Bilan, and Mercy!", "success");
});

// Update the displayAsTable function to include information about account managers
function displayAsTable(sortedData) {
    const outputContainer = document.getElementById('output');
    outputContainer.innerHTML = ''; // Clear any existing content

    // Add information about the account managers
    const info = document.createElement('div');
    info.style.marginBottom = '20px';
    info.innerHTML = `
        <strong>Account Managers Distribution:</strong>
        <ul>
            <li>Ahmed's Portfolio</li>
            <li>Bilan's Portfolio</li>
            <li>Mercy's Portfolio</li>
        </ul>
        <p>When downloading, accounts will be distributed evenly among these managers, maintaining all vehicles under the same account manager.</p>
    `;
    outputContainer.appendChild(info);

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');

    // Create the table header row
    const headers = ["Account", "Vehicle", "Last Message Time"];
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Add data rows
    sortedData.forEach(data => {
        data.vehicles.forEach(vehicle => {
            const row = document.createElement('tr');
            const accountCell = document.createElement('td');
            accountCell.textContent = data.account;
            row.appendChild(accountCell);

            const vehicleCell = document.createElement('td');
            vehicleCell.textContent = vehicle.vehicle;
            row.appendChild(vehicleCell);

            const lastMessageTimeCell = document.createElement('td');
            lastMessageTimeCell.textContent = vehicle.lastMessageTime;
            row.appendChild(lastMessageTimeCell);

            table.appendChild(row);
        });
    });

    outputContainer.appendChild(table);

    // Show the download button
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.classList.remove('hidden');
}
}
