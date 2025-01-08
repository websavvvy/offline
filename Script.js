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

    output.textContent = ""; // Clear the output field
    initializeProgressBar(progress, progressBar);

    setTimeout(() => processExcelFile(uploadedFile, progress, progressBar), 500);
});

// Initialize and manage progress bar
function initializeProgressBar(progress, progressBar) {
    progress.classList.remove("hidden");
    progressBar.style.width = "0%";
    
    let progressInterval;
    const updateProgress = () => {
        const currentWidth = parseInt(progressBar.style.width);
        if (currentWidth < 90) {
            progressBar.style.width = `${currentWidth + 5}%`;
            progressInterval = requestAnimationFrame(updateProgress);
        }
    };

    requestAnimationFrame(updateProgress);
    return progressInterval;
}

// Process the Excel file
function processExcelFile(file, progress, progressBar) {
    const reader = new FileReader();
    let progressInterval;

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const jsonData = processWorkbook(workbook);

        if (!jsonData) {
            progress.classList.add("hidden");
            cancelAnimationFrame(progressInterval);
            return;
        }

        const groupedData = groupByAccount(jsonData);
        const sortedData = sortByOfflineDuration(groupedData);
        displayAsTable(sortedData);
        createAndDownloadWorkbook(groupedData);

        finishProcessing(progress, progressBar, progressInterval);
    };

    reader.onerror = function() {
        showNotification("Could not read the file.", "error");
        progress.classList.add("hidden");
        cancelAnimationFrame(progressInterval);
    };

    reader.readAsArrayBuffer(file);
}

// Process workbook and validate columns
function processWorkbook(workbook) {
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    const headers = jsonData[0];
    const requiredColumns = ["Name", "Account", "Last message time"];
    const normalizedHeaders = headers.map(header => header.trim().toLowerCase());
    const columnIndexes = requiredColumns.map(col => 
        normalizedHeaders.findIndex(header => header === col.trim().toLowerCase())
    );

    if (columnIndexes.some(index => index === -1)) {
        const missingColumns = requiredColumns.filter((col, index) => columnIndexes[index] === -1);
        showNotification(`Missing columns: ${missingColumns.join(", ")}`, "error");
        return null;
    }

    return jsonData.slice(1).map(row => ({
        Name: row[columnIndexes[0]],
        Account: row[columnIndexes[1]],
        LastMessageTime: convertToRelativeTime(row[columnIndexes[2]] || "")
    }));
}

// Convert date to relative time
function convertToRelativeTime(dateString) {
    if (!dateString || typeof dateString !== "string") return "Invalid Date";

    const dateParts = dateString.split(' ');
    if (dateParts.length !== 2) return "Invalid Date";

    const dateArray = dateParts[0].split('.');
    const timeArray = dateParts[1].split(':');
    if (dateArray.length !== 3 || timeArray.length !== 3) return "Invalid Date";

    const formattedDate = `${dateArray[2]}-${dateArray[1]}-${dateArray[0]} ${timeArray[0]}:${timeArray[1]}:${timeArray[2]}`;
    const date = new Date(formattedDate);
    const now = new Date();
    const diffInMs = now - date;

    if (isNaN(date.getTime())) return "Invalid Date";

    const days = Math.floor(diffInMs / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diffInMs % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((diffInMs % (1000 * 60 * 60)) / (1000 * 60));

    return days === 0 ? `${hours} h ${minutes} min ago` : `${days} days ${hours} h ${minutes} min ago`;
}

// Group data by account
function groupByAccount(data) {
    return data.reduce((acc, row) => {
        const account = row.Account;
        if (!acc[account]) acc[account] = [];
        acc[account].push({ vehicle: row.Name, lastMessageTime: row.LastMessageTime });
        return acc;
    }, {});
}

// Sort vehicles by offline duration
function sortByOfflineDuration(groupedData) {
    const accountData = Object.keys(groupedData).map(account => {
        const vehicles = groupedData[account].sort((a, b) => 
            convertToDuration(b.lastMessageTime) - convertToDuration(a.lastMessageTime)
        );
        return { account, vehicles };
    });

    return accountData.sort((a, b) => 
        convertToDuration(b.vehicles[0].lastMessageTime) - convertToDuration(a.vehicles[0].lastMessageTime)
    );
}

// Convert relative time to duration in minutes
function convertToDuration(relativeTime) {
    const daysRegex = /(\d+)\s*days/;
    const hoursRegex = /(\d+)\s*h/;
    const minutesRegex = /(\d+)\s*min/;

    const days = (relativeTime.match(daysRegex) || [0, 0])[1];
    const hours = (relativeTime.match(hoursRegex) || [0, 0])[1];
    const minutes = (relativeTime.match(minutesRegex) || [0, 0])[1];

    return parseInt(days) * 1440 + parseInt(hours) * 60 + parseInt(minutes);
}

// Display sorted data as a table
function displayAsTable(sortedData) {
    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    ["Account", "Vehicle", "Last Message Time"].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sortedData.forEach(data => {
        data.vehicles.forEach(vehicle => {
            const row = document.createElement('tr');
            [data.account, vehicle.vehicle, vehicle.lastMessageTime].forEach(text => {
                const td = document.createElement('td');
                td.textContent = text;
                row.appendChild(td);
            });
            table.appendChild(row);
        });
    });

    const output = document.getElementById('output');
    output.innerHTML = "";
    output.appendChild(table);
}

// Create and download workbook with account distributions
function createAndDownloadWorkbook(groupedData) {
    const accountDistribution = {
        ahmed: [
            "RAPZAM LOGISTICS", "GREEN GAS TRANSPORT LIMITED", "Peaceway", "Urban Energy Limited", "HYPERLOOP HAULIERS",
            "IZWE", "Eager Eagle Trucking Ltd", "ROUTESTAR LTD", "Intek Logistics", "GALAXY ENGINEERING LTD",
            "Gastec trading & supply co.Ltd", "Haga Logistics", "All Three M Enterprise", "KAAH INVESTMENTS LTD",
            "MSM Logistics", "MMS COMPANY", "BLACK EAGLE ENERGY LIMITED", "GASS TRANSPORT", "Breeze logistics limited",
            "Napi Energy Limited", "Ruby Transporters", "WARSAN GAS TRANSPORT LIMITED", "LUSAKA CITY COUNCIL",
            "MIKA MEATS", "JAAF LOGISTICS", "CHAMPION LOGISTICS", "bluebolt enterprise limited",
            "UNISCON ENGINEERING & TRADING LTD", "Tow Jam Recovery", "INFRATEL", "HANAF TRANSPORT", "Jabri Logistcs",
            "NRL Account", "Mohamed Osman", "KIBONDO GREEN FARM", "MILGO LOGISTICS COMPANY LTD"
        ],
        bilan: [
            "Afrisom", "Guuled Transport Ltd", "Legacy Logistic Ltd", "Kalkaalow Transport", "Hogol", "Abdul kadir",
            "SIMBA SUPPLY CHAIN", "Abushiri Transporter", "Ahmed Abdillah company", "F a kerrow ltd", "Bahad Transport",
            "Ayaan Transport", "LONGPING AGRISCIENCE", "PEMBE CO LTD", "MAMBA EQUIPMENT", "K&P CONSTRUCTIONS LTD",
            "Jankah Transport", "MSF PETROLEUM", "Ahmed Sudi", "Taran Logistics account", "ALLY ADANI MOHAMED",
            "SIMERA TRANSPORT LIMITED TZ", "HONEST LOGISTICS", "ALLYEN", "Shabelle", "AGS Engineering services Ltd",
            "Abdullahi Shire", "FORTE EQUIPMENT TZ LTD", "Dr Ethan account", "Masele Msita Account", "Asad Taran",
            "Tarac Trading Co.LTD", "BUILD MART COMPANY", "BULK DISTRIBUTERS", "JOMA LOGISTICS CO.LTD",
            "GAHER FREIGHT FORWADERS", "TAIFA GAS TANZANIA LTD"
        ],
        mercy: [
            "Highland Estates Transport", "Ariva", "Horyaal Transport", "Deeq", "Warsame investment"
        ]
    };

    const wb = XLSX.utils.book_new();
    
    // Create sheets for each account manager
    Object.entries(accountDistribution).forEach(([manager, accounts]) => {
        const sheetData = groupDataByAccounts(groupedData, accounts);
        wb.SheetNames.push(manager.charAt(0).toUpperCase() + manager.slice(1));
        wb.Sheets[manager.charAt(0).toUpperCase() + manager.slice(1)] = XLSX.utils.json_to_sheet(sheetData);
    });

    XLSX.writeFile(wb, 'Account_Distribution.xlsx');
}

// Filter data by selected accounts
function groupDataByAccounts(groupedData, selectedAccounts) {
    return Object.keys(groupedData)
        .filter(account => selectedAccounts.includes(account))
        .flatMap(account => groupedData[account].map(vehicle => ({
            Account: account,
            Name: vehicle.vehicle,
            LastMessageTime: vehicle.lastMessageTime
        })));
}

// Helper function to show notifications
function showNotification(message, type = "success") {
    const notification = document.getElementById("notification");
    notification.textContent = message;
    notification.className = `notification ${type}`;
    setTimeout(() => {
        notification.className = "notification hidden";
    }, 3000);
}

// Helper function to finish processing
function finishProcessing(progress, progressBar, progressInterval) {
    progressBar.style.width = "100%";
    setTimeout(() => {
        progress.classList.add("hidden");
        cancelAnimationFrame(progressInterval);
        showNotification("Excel file processed successfully.", "success");
    }, 500);
}