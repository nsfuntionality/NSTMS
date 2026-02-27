// ===== Auth Check =====
(function checkAuth() {
    const user = sessionStorage.getItem('tms_user');
    if (!user) {
        window.location.href = 'index.html';
        return;
    }
    const parsed = JSON.parse(user);
    document.getElementById('navUser').textContent = parsed.name;
})();

// ===== State =====
let fuelData = [];
let loadsData = [];
let reportData = []; // miles per day / odometer data
let filteredFuel = [];
let filteredLoads = [];
let filteredReport = [];

// ===== Storage Keys =====
const STORAGE_KEYS = {
    fuel: 'tms_fuel_data',
    loads: 'tms_loads_data',
    report: 'tms_report_data',
    files: 'tms_uploaded_files'
};

// ===== Init =====
document.addEventListener('DOMContentLoaded', function() {
    loadStoredData();
    setupNavigation();
    setupFilters();
    setupUpload();
    setupExport();
    setupActions();
    renderAll();
});

// ===== Data Loading =====
function loadStoredData() {
    try {
        const storedFuel = localStorage.getItem(STORAGE_KEYS.fuel);
        const storedLoads = localStorage.getItem(STORAGE_KEYS.loads);
        const storedReport = localStorage.getItem(STORAGE_KEYS.report);
        if (storedFuel) fuelData = JSON.parse(storedFuel);
        if (storedLoads) loadsData = JSON.parse(storedLoads);
        if (storedReport) reportData = JSON.parse(storedReport);
    } catch (e) {
        console.error('Error loading stored data:', e);
    }
    filteredFuel = [...fuelData];
    filteredLoads = [...loadsData];
    filteredReport = [...reportData];
}

function saveData() {
    try {
        localStorage.setItem(STORAGE_KEYS.fuel, JSON.stringify(fuelData));
        localStorage.setItem(STORAGE_KEYS.loads, JSON.stringify(loadsData));
        localStorage.setItem(STORAGE_KEYS.report, JSON.stringify(reportData));
    } catch (e) {
        console.error('Error saving data:', e);
        showToast('Storage limit reached. Consider clearing old data.', 'error');
    }
}

function saveFileRecord(name, type, rows) {
    const files = JSON.parse(localStorage.getItem(STORAGE_KEYS.files) || '[]');
    files.push({
        name: name,
        type: type,
        rows: rows,
        date: new Date().toISOString()
    });
    localStorage.setItem(STORAGE_KEYS.files, JSON.stringify(files));
}

// ===== Navigation =====
function setupNavigation() {
    document.querySelectorAll('.sidebar-nav li').forEach(function(item) {
        item.addEventListener('click', function() {
            const tab = this.dataset.tab;
            document.querySelectorAll('.sidebar-nav li').forEach(function(el) { el.classList.remove('active'); });
            this.classList.add('active');
            document.querySelectorAll('.tab-content').forEach(function(el) { el.classList.remove('active'); });
            document.getElementById('tab-' + tab).classList.add('active');
        });
    });

    document.getElementById('logoutBtn').addEventListener('click', function() {
        sessionStorage.removeItem('tms_user');
        window.location.href = 'index.html';
    });
}

// ===== Filters =====
function setupFilters() {
    document.getElementById('applyFilters').addEventListener('click', applyFilters);
    document.getElementById('clearFilters').addEventListener('click', clearFilters);
}

function populateFilterDropdowns() {
    const driverSet = new Set();
    const truckSet = new Set();
    const trailerSet = new Set();

    fuelData.forEach(function(r) {
        if (r.driverName) driverSet.add(r.driverName);
        if (r.unit) truckSet.add(r.unit);
    });
    loadsData.forEach(function(r) {
        if (r.driver) {
            // Some drivers have "Driver1 / Driver2" format
            r.driver.split('/').forEach(function(d) { driverSet.add(d.trim()); });
        }
        if (r.truck) truckSet.add(r.truck);
        if (r.trailer) trailerSet.add(r.trailer);
    });

    populateSelect('filterDriver', driverSet, 'All Drivers');
    populateSelect('filterTruck', truckSet, 'All Trucks');
    populateSelect('filterTrailer', trailerSet, 'All Trailers');
}

function populateSelect(id, values, defaultLabel) {
    var sel = document.getElementById(id);
    var currentVal = sel.value;
    sel.innerHTML = '<option value="">' + defaultLabel + '</option>';
    Array.from(values).sort().forEach(function(v) {
        var opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        sel.appendChild(opt);
    });
    sel.value = currentVal;
}

function applyFilters() {
    var startDate = document.getElementById('filterStartDate').value;
    var endDate = document.getElementById('filterEndDate').value;
    var driver = document.getElementById('filterDriver').value;
    var truck = document.getElementById('filterTruck').value;
    var trailer = document.getElementById('filterTrailer').value;

    filteredFuel = fuelData.filter(function(r) {
        if (startDate && r.tranDate && r.tranDate < startDate) return false;
        if (endDate && r.tranDate && r.tranDate > endDate) return false;
        if (driver && r.driverName && r.driverName.indexOf(driver) === -1) return false;
        if (truck && r.unit !== truck) return false;
        return true;
    });

    filteredLoads = loadsData.filter(function(r) {
        if (startDate && r.pickDate && r.pickDate < startDate) return false;
        if (endDate && r.pickDate && r.pickDate > endDate) return false;
        if (driver && r.driver && r.driver.indexOf(driver) === -1) return false;
        if (truck && r.truck !== truck) return false;
        if (trailer && r.trailer !== trailer) return false;
        return true;
    });

    filteredReport = reportData.filter(function(r) {
        if (startDate && r.date && r.date < startDate) return false;
        if (endDate && r.date && r.date > endDate) return false;
        return true;
    });

    renderAll();
    showToast('Filters applied');
}

function clearFilters() {
    document.getElementById('filterStartDate').value = '';
    document.getElementById('filterEndDate').value = '';
    document.getElementById('filterDriver').value = '';
    document.getElementById('filterTruck').value = '';
    document.getElementById('filterTrailer').value = '';
    filteredFuel = [...fuelData];
    filteredLoads = [...loadsData];
    filteredReport = [...reportData];
    renderAll();
    showToast('Filters cleared');
}

// ===== Rendering =====
function renderAll() {
    populateFilterDropdowns();
    renderFuelTable();
    renderLoadsTable();
    renderReport();
    renderStoredFiles();
    updateStats();
}

function renderFuelTable() {
    var tbody = document.querySelector('#fuelTable tbody');
    tbody.innerHTML = '';
    var totalQty = 0;
    var totalAmt = 0;

    if (filteredFuel.length === 0) {
        tbody.innerHTML = '<tr><td colspan="13" class="empty-state">No fuel records found. Upload a fuel file to get started.</td></tr>';
    } else {
        filteredFuel.forEach(function(r) {
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + esc(r.cardNum) + '</td>' +
                '<td>' + formatDate(r.tranDate) + '</td>' +
                '<td>' + esc(r.tranTime) + '</td>' +
                '<td>' + esc(r.invoice) + '</td>' +
                '<td>' + esc(r.unit) + '</td>' +
                '<td>' + esc(r.driverName) + '</td>' +
                '<td>' + esc(r.locationName) + '</td>' +
                '<td>' + esc(r.city) + '</td>' +
                '<td>' + esc(r.state) + '</td>' +
                '<td>' + esc(r.item) + '</td>' +
                '<td class="amount-cell">$' + num(r.unitPrice) + '</td>' +
                '<td class="amount-cell">' + num(r.qty) + '</td>' +
                '<td class="amount-cell">$' + num(r.amt) + '</td>';
            tbody.appendChild(tr);
            totalQty += (parseFloat(r.qty) || 0);
            totalAmt += (parseFloat(r.amt) || 0);
        });
    }

    document.getElementById('fuelTotalQty').innerHTML = '<strong>' + totalQty.toFixed(2) + '</strong>';
    document.getElementById('fuelTotalAmt').innerHTML = '<strong>$' + totalAmt.toFixed(2) + '</strong>';
}

function renderLoadsTable() {
    var tbody = document.querySelector('#loadsTable tbody');
    tbody.innerHTML = '';

    if (filteredLoads.length === 0) {
        tbody.innerHTML = '<tr><td colspan="14" class="empty-state">No load records found. Upload a loads file to get started.</td></tr>';
    } else {
        filteredLoads.forEach(function(r) {
            var tr = document.createElement('tr');
            if (r.notes) tr.classList.add('highlight-row');
            tr.innerHTML =
                '<td>' + esc(r.invoiceId) + '</td>' +
                '<td>' + esc(r.loadNum) + '</td>' +
                '<td>' + esc(r.broker) + '</td>' +
                '<td>' + formatDate(r.pickDate) + '</td>' +
                '<td>' + esc(r.pickup) + '</td>' +
                '<td>' + formatDate(r.dropDate) + '</td>' +
                '<td>' + esc(r.dropoff) + '</td>' +
                '<td>' + esc(r.driver) + '</td>' +
                '<td>' + esc(r.truck) + '</td>' +
                '<td>' + esc(r.trailer) + '</td>' +
                '<td>' + esc(r.shippingId) + '</td>' +
                '<td>' + esc(r.puDatetime) + '</td>' +
                '<td>' + esc(r.doDatetime) + '</td>' +
                '<td>' + esc(r.notes) + '</td>';
            tbody.appendChild(tr);
        });
    }
}

function renderReport() {
    // Determine report meta from data
    var driver = '--';
    var truck = '--';
    var periodStart = '';
    var periodEnd = '';

    if (filteredLoads.length > 0) {
        var drivers = new Set();
        var trucks = new Set();
        filteredLoads.forEach(function(r) {
            if (r.driver) r.driver.split('/').forEach(function(d) { drivers.add(d.trim()); });
            if (r.truck) trucks.add(r.truck);
        });
        driver = Array.from(drivers).join(', ');
        truck = Array.from(trucks).join(', ');
    }
    if (filteredFuel.length > 0 && !driver || driver === '--') {
        var driverNames = new Set();
        filteredFuel.forEach(function(r) { if (r.driverName) driverNames.add(r.driverName); });
        if (driverNames.size) driver = Array.from(driverNames).join(', ');
        var truckNames = new Set();
        filteredFuel.forEach(function(r) { if (r.unit) truckNames.add(r.unit); });
        if (truckNames.size) truck = Array.from(truckNames).join(', ');
    }

    // Calculate period from all data
    var allDates = [];
    filteredFuel.forEach(function(r) { if (r.tranDate) allDates.push(r.tranDate); });
    filteredLoads.forEach(function(r) {
        if (r.pickDate) allDates.push(r.pickDate);
        if (r.dropDate) allDates.push(r.dropDate);
    });
    filteredReport.forEach(function(r) { if (r.date) allDates.push(r.date); });
    allDates.sort();
    if (allDates.length) {
        periodStart = allDates[0];
        periodEnd = allDates[allDates.length - 1];
    }

    document.getElementById('rptDriver').textContent = driver;
    document.getElementById('rptTruck').textContent = truck;
    document.getElementById('rptPeriod').textContent = periodStart && periodEnd
        ? formatDate(periodStart) + ' - ' + formatDate(periodEnd) : '--';

    // Fuel totals
    var fuelTotalAmt = 0;
    var fuelTotalGal = 0;
    filteredFuel.forEach(function(r) {
        fuelTotalAmt += (parseFloat(r.amt) || 0);
        fuelTotalGal += (parseFloat(r.qty) || 0);
    });
    document.getElementById('rptFuelCost').textContent = '$' + fuelTotalAmt.toFixed(2);
    document.getElementById('rptTotalLoads').textContent = filteredLoads.length;

    // Render Miles Per Day table
    var milesBody = document.querySelector('#reportMilesTable tbody');
    milesBody.innerHTML = '';
    var totalMiles = 0;
    var totalMissing = 0;

    if (filteredReport.length === 0) {
        milesBody.innerHTML = '<tr><td colspan="6" class="empty-state">No odometer/miles data. Upload a report file with miles data.</td></tr>';
    } else {
        filteredReport.forEach(function(r) {
            var tr = document.createElement('tr');
            if (r.missingMiles > 0) tr.classList.add('highlight-row');
            tr.innerHTML =
                '<td>' + formatDate(r.date) + '</td>' +
                '<td class="amount-cell">' + esc(r.startOdo) + '</td>' +
                '<td class="amount-cell">' + esc(r.endOdo) + '</td>' +
                '<td class="amount-cell">' + esc(r.totalMiles) + '</td>' +
                '<td class="amount-cell' + (r.missingMiles > 0 ? ' missing-miles' : '') + '">' + esc(r.missingMiles) + '</td>' +
                '<td>' + esc(r.notes || '') + '</td>';
            milesBody.appendChild(tr);
            totalMiles += (parseFloat(r.totalMiles) || 0);
            totalMissing += (parseFloat(r.missingMiles) || 0);
        });
    }

    document.getElementById('rptTotalMiles').textContent = totalMiles.toLocaleString();
    document.getElementById('rptMilesTotal').innerHTML = '<strong>' + totalMiles.toLocaleString() + '</strong>';
    document.getElementById('rptMissingTotal').innerHTML = '<strong>' + totalMissing.toLocaleString() + '</strong>';

    // Render Fuel Summary in report
    var fuelBody = document.querySelector('#reportFuelTable tbody');
    fuelBody.innerHTML = '';
    if (filteredFuel.length === 0) {
        fuelBody.innerHTML = '<tr><td colspan="5" class="empty-state">No fuel data for this period.</td></tr>';
    } else {
        filteredFuel.forEach(function(r) {
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + formatDate(r.tranDate) + '</td>' +
                '<td>' + esc(r.locationName) + '</td>' +
                '<td>' + esc(r.item) + '</td>' +
                '<td class="amount-cell">' + num(r.qty) + '</td>' +
                '<td class="amount-cell">$' + num(r.amt) + '</td>';
            fuelBody.appendChild(tr);
        });
    }
    document.getElementById('rptFuelGallons').innerHTML = '<strong>' + fuelTotalGal.toFixed(2) + '</strong>';
    document.getElementById('rptFuelTotal').innerHTML = '<strong>$' + fuelTotalAmt.toFixed(2) + '</strong>';

    // Render Load Summary in report
    var loadsBody = document.querySelector('#reportLoadsTable tbody');
    loadsBody.innerHTML = '';
    if (filteredLoads.length === 0) {
        loadsBody.innerHTML = '<tr><td colspan="8" class="empty-state">No load data for this period.</td></tr>';
    } else {
        filteredLoads.forEach(function(r) {
            var tr = document.createElement('tr');
            if (r.notes) tr.classList.add('highlight-row');
            tr.innerHTML =
                '<td>' + esc(r.invoiceId) + '</td>' +
                '<td>' + formatDate(r.pickDate) + '</td>' +
                '<td>' + esc(r.pickup) + '</td>' +
                '<td>' + formatDate(r.dropDate) + '</td>' +
                '<td>' + esc(r.dropoff) + '</td>' +
                '<td>' + esc(r.trailer) + '</td>' +
                '<td>' + esc(r.shippingId) + '</td>' +
                '<td>' + esc(r.notes) + '</td>';
            loadsBody.appendChild(tr);
        });
    }
}

function renderStoredFiles() {
    var container = document.getElementById('storedFilesList');
    var files = JSON.parse(localStorage.getItem(STORAGE_KEYS.files) || '[]');

    if (files.length === 0) {
        container.innerHTML = '<p class="empty-state">No files stored yet. Upload files above to get started.</p>';
        return;
    }

    container.innerHTML = '';
    files.forEach(function(f, i) {
        var div = document.createElement('div');
        div.className = 'stored-file-item';
        div.innerHTML =
            '<div class="stored-file-info">' +
                '<span class="stored-file-name">' + esc(f.name) + '</span>' +
                '<span class="stored-file-meta">' + esc(f.type) + ' | ' + f.rows + ' records | Uploaded ' + new Date(f.date).toLocaleDateString() + '</span>' +
            '</div>' +
            '<button class="btn btn-sm btn-outline btn-danger" data-idx="' + i + '">Remove</button>';
        container.appendChild(div);
    });

    container.querySelectorAll('button[data-idx]').forEach(function(btn) {
        btn.addEventListener('click', function() {
            var idx = parseInt(this.dataset.idx);
            files.splice(idx, 1);
            localStorage.setItem(STORAGE_KEYS.files, JSON.stringify(files));
            renderStoredFiles();
        });
    });
}

function updateStats() {
    document.getElementById('statFuel').textContent = fuelData.length;
    document.getElementById('statLoads').textContent = loadsData.length;
    var totalCost = 0;
    fuelData.forEach(function(r) { totalCost += (parseFloat(r.amt) || 0); });
    document.getElementById('statFuelCost').textContent = '$' + totalCost.toFixed(2);
}

// ===== Upload Handlers =====
function setupUpload() {
    // Fuel file
    var fuelInput = document.getElementById('fuelFileInput');
    fuelInput.addEventListener('change', function() {
        document.getElementById('fuelFileName').textContent = this.files[0] ? this.files[0].name : 'No file chosen';
        document.getElementById('uploadFuelBtn').disabled = !this.files[0];
    });
    document.getElementById('uploadFuelBtn').addEventListener('click', function() {
        var file = fuelInput.files[0];
        if (!file) return;
        parseExcelFile(file, 'fuel');
    });

    // Loads file
    var loadsInput = document.getElementById('loadsFileInput');
    loadsInput.addEventListener('change', function() {
        document.getElementById('loadsFileName').textContent = this.files[0] ? this.files[0].name : 'No file chosen';
        document.getElementById('uploadLoadsBtn').disabled = !this.files[0];
    });
    document.getElementById('uploadLoadsBtn').addEventListener('click', function() {
        var file = loadsInput.files[0];
        if (!file) return;
        parseExcelFile(file, 'loads');
    });

    // Report file
    var reportInput = document.getElementById('reportFileInput');
    reportInput.addEventListener('change', function() {
        document.getElementById('reportFileName').textContent = this.files[0] ? this.files[0].name : 'No file chosen';
        document.getElementById('uploadReportBtn').disabled = !this.files[0];
    });
    document.getElementById('uploadReportBtn').addEventListener('click', function() {
        var file = reportInput.files[0];
        if (!file) return;
        parseExcelFile(file, 'report');
    });
}

function parseExcelFile(file, type) {
    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array', cellDates: true });

            if (type === 'fuel') {
                parseFuelWorkbook(workbook, file.name);
            } else if (type === 'loads') {
                parseLoadsWorkbook(workbook, file.name);
            } else if (type === 'report') {
                parseReportWorkbook(workbook, file.name);
            }
        } catch (err) {
            console.error('Parse error:', err);
            showToast('Error parsing file: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseFuelWorkbook(workbook, fileName) {
    // Try to find Fuel sheet, otherwise use first sheet
    var sheetName = workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('fuel') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row
    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (row && row.some(function(c) {
            return String(c).toLowerCase().indexOf('card') !== -1 ||
                   String(c).toLowerCase().indexOf('tran date') !== -1;
        })) {
            headerIdx = i;
            break;
        }
    }

    if (headerIdx === -1) {
        showToast('Could not find fuel data headers in file', 'error');
        return;
    }

    // Map columns
    var headers = rows[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
    var colMap = {
        cardNum: findCol(headers, ['card']),
        tranDate: findCol(headers, ['tran date', 'date']),
        tranTime: findCol(headers, ['time', 'trans. time']),
        invoice: findCol(headers, ['invoice']),
        unit: findCol(headers, ['unit']),
        driverName: findCol(headers, ['driver']),
        odometer: findCol(headers, ['odometer', 'odo']),
        locationName: findCol(headers, ['location']),
        city: findCol(headers, ['city']),
        state: findCol(headers, ['state', 'prov']),
        fees: findCol(headers, ['fees', 'fee']),
        item: findCol(headers, ['item']),
        unitPrice: findCol(headers, ['unit price', 'price']),
        qty: findCol(headers, ['qty', 'quantity']),
        amt: findCol(headers, ['amt', 'amount']),
        db: findCol(headers, ['db']),
        currency: findCol(headers, ['currency'])
    };

    var newData = [];
    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row || !row[colMap.cardNum] && !row[colMap.tranDate]) continue;
        newData.push({
            cardNum: row[colMap.cardNum] || '',
            tranDate: excelDateToStr(row[colMap.tranDate]),
            tranTime: formatExcelTime(row[colMap.tranTime]),
            invoice: row[colMap.invoice] || '',
            unit: row[colMap.unit] || '',
            driverName: row[colMap.driverName] || '',
            odometer: row[colMap.odometer] || 0,
            locationName: row[colMap.locationName] || '',
            city: row[colMap.city] || '',
            state: row[colMap.state] || '',
            fees: row[colMap.fees] || 0,
            item: row[colMap.item] || '',
            unitPrice: row[colMap.unitPrice] || 0,
            qty: row[colMap.qty] || 0,
            amt: row[colMap.amt] || 0,
            db: row[colMap.db] || '',
            currency: row[colMap.currency] || ''
        });
    }

    fuelData = fuelData.concat(newData);
    filteredFuel = [...fuelData];
    saveData();
    saveFileRecord(fileName, 'Fuel', newData.length);
    renderAll();
    showToast('Loaded ' + newData.length + ' fuel records from ' + fileName, 'success');
}

function parseLoadsWorkbook(workbook, fileName) {
    var sheetName = workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('load') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row
    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (row && row.some(function(c) {
            return String(c).toLowerCase().indexOf('invoiceid') !== -1 ||
                   String(c).toLowerCase().indexOf('load #') !== -1 ||
                   String(c).toLowerCase().indexOf('load') !== -1 && String(c).toLowerCase().indexOf('broker') !== -1;
        })) {
            headerIdx = i;
            break;
        }
    }

    if (headerIdx === -1) {
        showToast('Could not find load data headers in file', 'error');
        return;
    }

    var headers = rows[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
    var colMap = {
        invoiceId: findCol(headers, ['invoiceid', 'invoice']),
        loadNum: findCol(headers, ['load #', 'load']),
        broker: findCol(headers, ['broker']),
        pickDate: findCol(headers, ['pick date', 'pickup date']),
        pickup: findCol(headers, ['pickup']),
        dropDate: findCol(headers, ['drop date', 'dropoff date']),
        dropoff: findCol(headers, ['dropoff']),
        driver: findCol(headers, ['driver']),
        truck: findCol(headers, ['truck', 'truckname']),
        trailer: findCol(headers, ['trailer']),
        shippingId: findCol(headers, ['shipping', 'shipping id']),
        puDatetime: findCol(headers, ['pickup datetime', 'pu datetime']),
        doDatetime: findCol(headers, ['delivery datetime', 'do datetime']),
        notes: findCol(headers, ['notes', 'note'])
    };

    var newData = [];
    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row || (!row[colMap.invoiceId] && !row[colMap.loadNum])) continue;
        newData.push({
            invoiceId: row[colMap.invoiceId] || '',
            loadNum: row[colMap.loadNum] || '',
            broker: row[colMap.broker] || '',
            pickDate: excelDateToStr(row[colMap.pickDate]),
            pickup: row[colMap.pickup] || '',
            dropDate: excelDateToStr(row[colMap.dropDate]),
            dropoff: row[colMap.dropoff] || '',
            driver: row[colMap.driver] || '',
            truck: row[colMap.truck] || '',
            trailer: row[colMap.trailer] || '',
            shippingId: row[colMap.shippingId] || '',
            puDatetime: row[colMap.puDatetime] || '',
            doDatetime: row[colMap.doDatetime] || '',
            notes: row[colMap.notes] || ''
        });
    }

    loadsData = loadsData.concat(newData);
    filteredLoads = [...loadsData];
    saveData();
    saveFileRecord(fileName, 'Loads', newData.length);
    renderAll();
    showToast('Loaded ' + newData.length + ' load records from ' + fileName, 'success');
}

function parseReportWorkbook(workbook, fileName) {
    // Try to find "Report Date - Date" sheet, or sheet with odometer data
    var sheetName = workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('report') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Strategy: Find columns with Date, Start Odo, End Odo, Total Mile, Missing Miles
    // In the sample file, Miles Per Day data starts around row 9 (column R onwards)
    // We need to find the date column and odometer columns

    var newData = [];

    // First, try to find explicit header row for miles data
    var headerIdx = -1;
    var dateCol = -1;
    var startOdoCol = -1;
    var endOdoCol = -1;
    var totalMileCol = -1;
    var missingCol = -1;
    var notesCol = -1;

    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (!row) continue;
        for (var j = 0; j < row.length; j++) {
            var cellStr = String(row[j]).toLowerCase().trim();
            if (cellStr === 'date' && dateCol === -1) { dateCol = j; headerIdx = i; }
            if (cellStr.indexOf('start odo') !== -1 || cellStr.indexOf('start_odo') !== -1) startOdoCol = j;
            if (cellStr.indexOf('end odo') !== -1 || cellStr.indexOf('end_odo') !== -1) endOdoCol = j;
            if (cellStr.indexOf('total mile') !== -1 || cellStr === 'total miles') totalMileCol = j;
            if (cellStr.indexOf('missing') !== -1) missingCol = j;
        }
        if (dateCol !== -1 && startOdoCol !== -1) break;
    }

    if (headerIdx === -1 || dateCol === -1) {
        // Fallback: try to find the Miles Per Day section in the combined report file
        // The sample file has this data starting at column R (index 17)
        for (var i = 0; i < Math.min(rows.length, 15); i++) {
            var row = rows[i];
            if (!row) continue;
            for (var j = 0; j < row.length; j++) {
                var cellStr = String(row[j]).toLowerCase().trim();
                if (cellStr === 'date') { dateCol = j; headerIdx = i; }
                if (cellStr === 'start odo') startOdoCol = j;
                if (cellStr === 'end odo') endOdoCol = j;
                if (cellStr === 'total mile') totalMileCol = j;
                if (cellStr === 'missing miles') missingCol = j;
                if (cellStr === 'details' || cellStr === 'notes') notesCol = j;
            }
            if (dateCol !== -1 && startOdoCol !== -1) break;
        }
    }

    if (dateCol === -1) {
        showToast('Could not find date/odometer columns in report file. Looking for: Date, Start Odo, End Odo, Total Mile, Missing Miles', 'error');
        return;
    }

    // Defaults for columns not found
    if (startOdoCol === -1) startOdoCol = dateCol + 1;
    if (endOdoCol === -1) endOdoCol = dateCol + 2;
    if (totalMileCol === -1) totalMileCol = dateCol + 3;
    if (missingCol === -1) missingCol = dateCol + 4;

    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row) continue;
        var dateVal = row[dateCol];
        if (!dateVal) continue;
        var dateStr = excelDateToStr(dateVal);
        if (!dateStr) continue;

        var record = {
            date: dateStr,
            startOdo: row[startOdoCol] || 0,
            endOdo: row[endOdoCol] || 0,
            totalMiles: row[totalMileCol] || 0,
            missingMiles: row[missingCol] || 0,
            notes: notesCol !== -1 ? (row[notesCol] || '') : ''
        };
        newData.push(record);
    }

    reportData = reportData.concat(newData);
    filteredReport = [...reportData];
    saveData();
    saveFileRecord(fileName, 'Report (Miles/Odo)', newData.length);
    renderAll();
    showToast('Loaded ' + newData.length + ' daily mile records from ' + fileName, 'success');
}

// ===== Export =====
function setupExport() {
    document.getElementById('exportFuelCSV').addEventListener('click', function() {
        exportCSV(filteredFuel, [
            'cardNum', 'tranDate', 'tranTime', 'invoice', 'unit', 'driverName',
            'odometer', 'locationName', 'city', 'state', 'fees', 'item',
            'unitPrice', 'qty', 'amt', 'db', 'currency'
        ], [
            'Card #', 'Tran Date', 'Trans. Time', 'Invoice', 'Unit', 'Driver Name',
            'Odometer', 'Location Name', 'City', 'State', 'Fees', 'Item',
            'Unit Price', 'Qty', 'Amount', 'DB', 'Currency'
        ], 'fuel_export.csv');
    });

    document.getElementById('exportLoadsCSV').addEventListener('click', function() {
        exportCSV(filteredLoads, [
            'invoiceId', 'loadNum', 'broker', 'pickDate', 'pickup', 'dropDate',
            'dropoff', 'driver', 'truck', 'trailer', 'shippingId', 'puDatetime', 'doDatetime', 'notes'
        ], [
            'Invoice ID', 'Load #', 'Broker', 'Pick Date', 'Pickup', 'Drop Date',
            'Dropoff', 'Driver', 'Truck', 'Trailer', 'Shipping ID', 'PU Datetime', 'DO Datetime', 'Notes'
        ], 'loads_export.csv');
    });

    document.getElementById('exportReportCSV').addEventListener('click', function() {
        exportCSV(filteredReport, [
            'date', 'startOdo', 'endOdo', 'totalMiles', 'missingMiles', 'notes'
        ], [
            'Date', 'Start Odometer', 'End Odometer', 'Total Miles', 'Missing Miles', 'Notes'
        ], 'report_export.csv');
    });
}

function exportCSV(data, fields, headers, filename) {
    if (!data.length) {
        showToast('No data to export', 'error');
        return;
    }
    var csvRows = [headers.join(',')];
    data.forEach(function(row) {
        var vals = fields.map(function(f) {
            var v = row[f] != null ? String(row[f]) : '';
            if (v.indexOf(',') !== -1 || v.indexOf('"') !== -1) {
                v = '"' + v.replace(/"/g, '""') + '"';
            }
            return v;
        });
        csvRows.push(vals.join(','));
    });
    var blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
    showToast('Exported ' + data.length + ' records to ' + filename, 'success');
}

// ===== Actions =====
function setupActions() {
    document.getElementById('clearAllData').addEventListener('click', function() {
        if (!confirm('Are you sure you want to clear ALL stored data? This cannot be undone.')) return;
        fuelData = [];
        loadsData = [];
        reportData = [];
        filteredFuel = [];
        filteredLoads = [];
        filteredReport = [];
        localStorage.removeItem(STORAGE_KEYS.fuel);
        localStorage.removeItem(STORAGE_KEYS.loads);
        localStorage.removeItem(STORAGE_KEYS.report);
        localStorage.removeItem(STORAGE_KEYS.files);
        renderAll();
        showToast('All data cleared', 'success');
    });

    document.getElementById('loadSampleData').addEventListener('click', loadSampleData);
}

function loadSampleData() {
    // Load sample data embedded directly for GitHub Pages (no server needed)
    var sampleFuel = [
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.812,qty:161.23,amt:614.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"DEFD",unitPrice:4.399,qty:15.14,amt:66.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"ULSD",unitPrice:3.685,qty:161.65,amt:595.68,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:11.29,amt:49.66,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-07",tranTime:"07:00",invoice:57250,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:3.277,qty:164.76,amt:539.92,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-07",tranTime:"07:00",invoice:57250,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"DEFD",unitPrice:4.399,qty:10.75,amt:47.30,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-08",tranTime:"19:17",invoice:47082,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.576,qty:174.86,amt:625.30,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-08",tranTime:"19:17",invoice:47082,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:9.94,amt:43.73,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-12",tranTime:"05:02",invoice:6237,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.925,qty:170.72,amt:670.07,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-12",tranTime:"05:02",invoice:6237,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"DEFD",unitPrice:4.399,qty:11.61,amt:51.05,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-13",tranTime:"02:22",invoice:52717,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.573,qty:167.15,amt:597.22,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-14",tranTime:"09:01",invoice:66434,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:3.427,qty:83.04,amt:284.58,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-16",tranTime:"12:14",invoice:66469,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:3.427,qty:137.63,amt:471.66,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-16",tranTime:"12:14",invoice:66469,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"DEFD",unitPrice:4.399,qty:18.42,amt:81.01,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-17",tranTime:"19:32",invoice:44196,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"ULSD",unitPrice:3.923,qty:162.37,amt:636.98,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-17",tranTime:"19:32",invoice:44196,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:8.75,amt:38.50,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-19",tranTime:"18:11",invoice:10597,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:4.026,qty:175.38,amt:706.08,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-21",tranTime:"10:26",invoice:60652,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.567,qty:132.79,amt:473.67,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-21",tranTime:"10:26",invoice:60652,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:19.59,amt:86.19,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-24",tranTime:"02:27",invoice:47331,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"ULSD",unitPrice:3.491,qty:167.77,amt:585.69,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-24",tranTime:"02:27",invoice:47331,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:11.58,amt:50.95,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-27",tranTime:"23:42",invoice:81071,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:2.863,qty:172.43,amt:493.67,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-27",tranTime:"23:42",invoice:81071,unit:"MPL5046",driverName:"Manjinder Bajwa",odometer:0,locationName:"PETRO RACINE - FUEL 491",city:"STURTEVANT",state:"WI",fees:0,item:"DEFD",unitPrice:4.399,qty:9.65,amt:42.47,db:"N",currency:"USD/Gallons"}
    ];

    var sampleLoads = [
        {invoiceId:"KLL56230",loadNum:15368128,broker:"Avenger Logistics Llc",pickDate:"2025-10-31",pickup:"Sharonville, OH 45241",dropDate:"2025-11-03",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1016",shippingId:580988,puDatetime:"PU 10/31 | Time 03PM",doDatetime:"DO 11/03 | Local Miles",notes:""},
        {invoiceId:"KLL56375",loadNum:15393089,broker:"Avenger Logistics Llc",pickDate:"2025-11-03",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-03",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1016",shippingId:6851123,puDatetime:"PU 11/03 | Local Miles",doDatetime:"DO 11/03 | Time 02PM",notes:""},
        {invoiceId:"KLL56376",loadNum:15392627,broker:"Avenger Logistics Llc",pickDate:"2025-11-03",pickup:"Sharonville, OH 45241",dropDate:"2025-11-04",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1008",shippingId:581331,puDatetime:"PU 11/03 | Time 02PM",doDatetime:"DO 11/04 | Local Miles",notes:""},
        {invoiceId:"KLL56397",loadNum:15392834,broker:"Avenger Logistics Llc",pickDate:"2025-11-04",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-04",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1008",shippingId:6851159,puDatetime:"PU 11/04 | Local Miles",doDatetime:"DO 11/04 | Time 07PM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56398",loadNum:15392694,broker:"Avenger Logistics Llc",pickDate:"2025-11-04",pickup:"Sharonville, OH 45241",dropDate:"2025-11-06",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa / Rohit Kumar",truck:"MPL5046",trailer:"MPL1015",shippingId:581572,puDatetime:"PU 11/04 | Time 07PM",doDatetime:"DO 11/06 | Local Driver",notes:""},
        {invoiceId:"KLL56419",loadNum:15392663,broker:"Avenger Logistics Llc",pickDate:"2025-11-06",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-06",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa / Rohit Kumar",truck:"MPL5046",trailer:"MPL1002",shippingId:6851195,puDatetime:"PU 11/06 | Local Driver",doDatetime:"DO 11/06 | Time 02PM",notes:""},
        {invoiceId:"KLL56420",loadNum:15393302,broker:"Avenger Logistics Llc",pickDate:"2025-11-06",pickup:"Sharonville, OH 45241",dropDate:"2025-11-07",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1047",shippingId:581776,puDatetime:"PU 11/06 | Time 02PM",doDatetime:"DO 11/07 | Local Miles",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56451",loadNum:15392866,broker:"Avenger Logistics Llc",pickDate:"2025-11-08",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-08",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa / Vikas Saini",truck:"MPL5046",trailer:"MPL1018",shippingId:6851242,puDatetime:"PU 11/08 | Local Driver",doDatetime:"DO 11/08 | Time 03PM",notes:""},
        {invoiceId:"KLL56452",loadNum:15393042,broker:"Avenger Logistics Llc",pickDate:"2025-11-08",pickup:"Sharonville, OH 45241",dropDate:"2025-11-10",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1047",shippingId:582017,puDatetime:"PU 11/08 | Time 03PM",doDatetime:"DO 11/10 | Local Miles",notes:""},
        {invoiceId:"KLL56568",loadNum:15425555,broker:"Avenger Logistics Llc",pickDate:"2025-11-10",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-11",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1047",shippingId:6851294,puDatetime:"PU 11/10 | Local Miles",doDatetime:"DO 11/11 | Time 03AM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56569",loadNum:15425564,broker:"Avenger Logistics Llc",pickDate:"2025-11-11",pickup:"Sharonville, OH 45241",dropDate:"2025-11-12",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1011",shippingId:582075,puDatetime:"PU 11/11 | Time 03AM",doDatetime:"DO 11/12 | Local Miles",notes:""},
        {invoiceId:"KLL56588",loadNum:15417727,broker:"Avenger Logistics Llc",pickDate:"2025-11-12",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-12",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1011",shippingId:6851324,puDatetime:"PU 11/12 | Local Miles",doDatetime:"DO 11/12 | Time 11PM",notes:""},
        {invoiceId:"KLL56589",loadNum:15417160,broker:"Avenger Logistics Llc",pickDate:"2025-11-12",pickup:"Sharonville, OH 45241",dropDate:"2025-11-13",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa / Baljit Singh",truck:"MPL5046",trailer:"MPL1018",shippingId:582576,puDatetime:"PU 11/12 | Time 11PM",doDatetime:"DO 11/13 | Local Driver",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56604",loadNum:15417280,broker:"Avenger Logistics Llc",pickDate:"2025-11-13",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-13",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1014",shippingId:6851352,puDatetime:"PU 11/13 | Local Driver",doDatetime:"DO 11/13 | Time 11PM",notes:"Wrong trailer no. on Eld"},
        {invoiceId:"KLL56605",loadNum:15417700,broker:"Avenger Logistics Llc",pickDate:"2025-11-13",pickup:"Sharonville, OH 45241",dropDate:"2025-11-14",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa / Nikhal Panday",truck:"MPL5046",trailer:"MPL1004",shippingId:582677,puDatetime:"PU 11/13 | Time 11PM",doDatetime:"DO 11/14 | Local Miles",notes:"Wrong Shiping id on Eld"},
        {invoiceId:"KLL56628",loadNum:15417205,broker:"Avenger Logistics Llc",pickDate:"2025-11-14",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-15",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1016",shippingId:6851389,puDatetime:"PU 11/14 | Local Miles",doDatetime:"DO 11/15 | Time 02AM",notes:""},
        {invoiceId:"KLL56629",loadNum:15417158,broker:"Avenger Logistics Llc",pickDate:"2025-11-15",pickup:"Sharonville, OH 45241",dropDate:"2025-11-16",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1027",shippingId:582976,puDatetime:"PU 11/15 | Time 02AM",doDatetime:"DO 11/16 | Local Driver",notes:"Need to recheck Local Driver"},
        {invoiceId:"KLL56756",loadNum:15440628,broker:"Avenger Logistics Llc",pickDate:"2025-11-17",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-17",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1023",shippingId:6851448,puDatetime:"PU 11/17 | Local Miles",doDatetime:"DO 11/17 | Time 06PM",notes:""},
        {invoiceId:"KLL56757",loadNum:15440770,broker:"Avenger Logistics Llc",pickDate:"2025-11-17",pickup:"Sharonville, OH 45241",dropDate:"2025-11-18",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1027",shippingId:583188,puDatetime:"PU 11/17 | Time 06PM",doDatetime:"DO 11/18 | Local Miles",notes:""},
        {invoiceId:"KLL56776",loadNum:15441263,broker:"Avenger Logistics Llc",pickDate:"2025-11-18",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-18",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1027",shippingId:6851480,puDatetime:"PU 11/18 | Local Miles",doDatetime:"DO 11/18 | Time 09PM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56777",loadNum:15440413,broker:"Avenger Logistics Llc",pickDate:"2025-11-18",pickup:"Sharonville, OH 45241",dropDate:"2025-11-19",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1014",shippingId:583590,puDatetime:"PU 11/18 | Time 09PM",doDatetime:"DO 11/19 | Local Miles",notes:""},
        {invoiceId:"KLL56796",loadNum:15440397,broker:"Avenger Logistics Llc",pickDate:"2025-11-19",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-20",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1014",shippingId:6851514,puDatetime:"PU 11/19 | Local Miles",doDatetime:"DO 11/20 | Time 11AM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56797",loadNum:15440941,broker:"Avenger Logistics Llc",pickDate:"2025-11-20",pickup:"Sharonville, OH 45241",dropDate:"2025-11-20",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1011",shippingId:583860,puDatetime:"PU 11/20 | Time 11AM",doDatetime:"DO 11/20 | Local Driver",notes:"Needs to recheck*"},
        {invoiceId:"KLL56812",loadNum:15440662,broker:"Avenger Logistics Llc",pickDate:"2025-11-20",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-21",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1011",shippingId:6851544,puDatetime:"PU 11/20 | Local Driver",doDatetime:"DO 11/21 | Time 01PM",notes:"Needs to recheck*"},
        {invoiceId:"KLL56813",loadNum:15440633,broker:"Avenger Logistics Llc",pickDate:"2025-11-21",pickup:"Sharonville, OH 45241",dropDate:"2025-11-23",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1011",shippingId:584083,puDatetime:"PU 11/21 | Time 01PM",doDatetime:"DO 11/23 | Local Miles",notes:"Wrong Shiping id on Eld"},
        {invoiceId:"KLL56936",loadNum:15464416,broker:"Avenger Logistics Llc",pickDate:"2025-11-23",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-24",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa / Nikhal Panday",truck:"MPL5046",trailer:"MPL1009",shippingId:6851594,puDatetime:"PU 11/23 | Local Driver",doDatetime:"DO 11/24 | Time 01AM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56937",loadNum:15464688,broker:"Avenger Logistics Llc",pickDate:"2025-11-24",pickup:"Sharonville, OH 45241",dropDate:"2025-11-24",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1026",shippingId:584271,puDatetime:"PU 11/24 | Time 01AM",doDatetime:"DO 11/24 | Local Miles",notes:""},
        {invoiceId:"KLL56956",loadNum:15464436,broker:"Avenger Logistics Llc",pickDate:"2025-11-24",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-25",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1026",shippingId:6851621,puDatetime:"PU 11/24 | Local Miles",doDatetime:"DO 11/25 | Time 04AM",notes:""},
        {invoiceId:"KLL56957",loadNum:15464538,broker:"Avenger Logistics Llc",pickDate:"2025-11-25",pickup:"Sharonville, OH 45241",dropDate:"2025-11-27",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1018",shippingId:584150,puDatetime:"PU 11/25 | Time 04AM",doDatetime:"DO 11/27 | Local Miles",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56978",loadNum:15464705,broker:"Avenger Logistics Llc",pickDate:"2025-11-27",pickup:"Sheboygan, WI 53081",dropDate:"2025-11-28",dropoff:"Sharonville, OH 45241",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1018",shippingId:6851664,puDatetime:"PU 11/27 | Local Miles",doDatetime:"DO 11/28 | Time 10PM",notes:"Wrong Shiping id and trailer no. on Eld"},
        {invoiceId:"KLL56979",loadNum:15464776,broker:"Avenger Logistics Llc",pickDate:"2025-11-28",pickup:"Sharonville, OH 45241",dropDate:"2025-12-01",dropoff:"Sheboygan, WI 53081",driver:"Manjinder Bajwa",truck:"MPL5046",trailer:"MPL1023",shippingId:584829,puDatetime:"PU 11/28 | Time 10PM",doDatetime:"DO 12/01 | Local Miles",notes:""}
    ];

    var sampleReport = [
        {date:"2025-11-01",startOdo:57685,endOdo:57858,totalMiles:173,missingMiles:0,notes:""},
        {date:"2025-11-02",startOdo:57858,endOdo:57858,totalMiles:0,missingMiles:0,notes:""},
        {date:"2025-11-03",startOdo:57858,endOdo:57858,totalMiles:0,missingMiles:428,notes:"Maybe book transfer to someone"},
        {date:"2025-11-03",startOdo:58286,endOdo:58924,totalMiles:638,missingMiles:0,notes:""},
        {date:"2025-11-04",startOdo:58924,endOdo:59067,totalMiles:143,missingMiles:136,notes:""},
        {date:"2025-11-04",startOdo:59203,endOdo:59587,totalMiles:384,missingMiles:0,notes:""},
        {date:"2025-11-05",startOdo:59587,endOdo:60274,totalMiles:687,missingMiles:0,notes:""},
        {date:"2025-11-06",startOdo:60274,endOdo:60274,totalMiles:0,missingMiles:15,notes:""},
        {date:"2025-11-06",startOdo:60289,endOdo:60981,totalMiles:692,missingMiles:0,notes:""},
        {date:"2025-11-07",startOdo:60981,endOdo:61048,totalMiles:67,missingMiles:134,notes:""},
        {date:"2025-11-07",startOdo:61182,endOdo:61939,totalMiles:757,missingMiles:0,notes:"Book transfer to SodhiTanvir Singh"},
        {date:"2025-11-08",startOdo:61939,endOdo:61939,totalMiles:0,missingMiles:0,notes:""},
        {date:"2025-11-08",startOdo:61939,endOdo:62632,totalMiles:693,missingMiles:0,notes:""},
        {date:"2025-11-09",startOdo:62632,endOdo:62696,totalMiles:64,missingMiles:0,notes:""},
        {date:"2025-11-10",startOdo:62696,endOdo:62696,totalMiles:0,missingMiles:134,notes:""},
        {date:"2025-11-10",startOdo:62830,endOdo:63013,totalMiles:183,missingMiles:0,notes:""},
        {date:"2025-11-11",startOdo:63013,endOdo:63534,totalMiles:521,missingMiles:0,notes:""},
        {date:"2025-11-12",startOdo:63534,endOdo:63597,totalMiles:63,missingMiles:135,notes:""},
        {date:"2025-11-12",startOdo:63732,endOdo:64176,totalMiles:444,missingMiles:0,notes:""},
        {date:"2025-11-13",startOdo:64176,endOdo:64945,totalMiles:769,missingMiles:0,notes:""},
        {date:"2025-11-14",startOdo:64945,endOdo:65324,totalMiles:379,missingMiles:138,notes:""}
    ];

    fuelData = sampleFuel;
    loadsData = sampleLoads;
    reportData = sampleReport;
    filteredFuel = [...fuelData];
    filteredLoads = [...loadsData];
    filteredReport = [...reportData];
    saveData();
    saveFileRecord('Sample - Fuel Nov 2025', 'Fuel', sampleFuel.length);
    saveFileRecord('Sample - Loads Nov 2025', 'Loads', sampleLoads.length);
    saveFileRecord('Sample - Report Nov 2025', 'Report (Miles/Odo)', sampleReport.length);
    renderAll();
    showToast('Sample data loaded successfully!', 'success');
}

// ===== Utility Functions =====
function findCol(headers, keywords) {
    for (var i = 0; i < headers.length; i++) {
        for (var k = 0; k < keywords.length; k++) {
            if (headers[i].indexOf(keywords[k]) !== -1) return i;
        }
    }
    return 0;
}

function excelDateToStr(val) {
    if (!val) return '';
    if (val instanceof Date) {
        var y = val.getFullYear();
        var m = String(val.getMonth() + 1).padStart(2, '0');
        var d = String(val.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    if (typeof val === 'string') {
        // Try to parse common date formats
        if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.substring(0, 10);
        var parsed = new Date(val);
        if (!isNaN(parsed.getTime())) {
            return parsed.getFullYear() + '-' + String(parsed.getMonth()+1).padStart(2,'0') + '-' + String(parsed.getDate()).padStart(2,'0');
        }
    }
    if (typeof val === 'number') {
        // Excel serial date
        var date = new Date((val - 25569) * 86400000);
        return date.getFullYear() + '-' + String(date.getMonth()+1).padStart(2,'0') + '-' + String(date.getDate()).padStart(2,'0');
    }
    return String(val);
}

function formatExcelTime(val) {
    if (!val) return '';
    if (val instanceof Date) {
        return String(val.getHours()).padStart(2, '0') + ':' + String(val.getMinutes()).padStart(2, '0');
    }
    if (typeof val === 'number') {
        var totalMinutes = Math.round(val * 24 * 60);
        var hours = Math.floor(totalMinutes / 60);
        var minutes = totalMinutes % 60;
        return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
    }
    return String(val);
}

function formatDate(dateStr) {
    if (!dateStr) return '';
    // dateStr is expected as YYYY-MM-DD
    var parts = String(dateStr).split('-');
    if (parts.length === 3) {
        return parts[1] + '/' + parts[2] + '/' + parts[0];
    }
    return dateStr;
}

function esc(val) {
    if (val == null || val === '') return '';
    var s = String(val);
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(s));
    return div.innerHTML;
}

function num(val) {
    var n = parseFloat(val);
    return isNaN(n) ? '0.00' : n.toFixed(2);
}

function showToast(message, type) {
    type = type || '';
    var container = document.getElementById('toastContainer');
    var toast = document.createElement('div');
    toast.className = 'toast ' + type;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(function() {
        toast.remove();
    }, 4000);
}
