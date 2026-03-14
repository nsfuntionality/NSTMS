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
let tripData = [];
let filteredFuel = [];
let filteredLoads = [];
let filteredReport = [];
let filteredTrip = [];

// ===== Storage Keys =====
const STORAGE_KEYS = {
    fuel: 'tms_fuel_data',
    loads: 'tms_loads_data',
    report: 'tms_report_data',
    trip: 'tms_trip_data',
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
    setupModal();
    setupSettings();
    applyAuditMode();
    renderAll();
    // Auto-load from Excel data files if no data in localStorage
    if (fuelData.length === 0 && loadsData.length === 0 && reportData.length === 0) {
        loadFromDataFiles();
    }
});

// ===== Data Loading =====
function loadStoredData() {
    try {
        const storedFuel = localStorage.getItem(STORAGE_KEYS.fuel);
        const storedLoads = localStorage.getItem(STORAGE_KEYS.loads);
        const storedReport = localStorage.getItem(STORAGE_KEYS.report);
        const storedTrip = localStorage.getItem(STORAGE_KEYS.trip);
        if (storedFuel) fuelData = JSON.parse(storedFuel);
        if (storedLoads) loadsData = JSON.parse(storedLoads);
        if (storedReport) reportData = JSON.parse(storedReport);
        if (storedTrip) tripData = JSON.parse(storedTrip);
    } catch (e) {
        console.error('Error loading stored data:', e);
    }
    filteredFuel = [...fuelData];
    filteredLoads = [...loadsData];
    filteredReport = [...reportData];
    filteredTrip = [...tripData];
}

function saveData() {
    try {
        localStorage.setItem(STORAGE_KEYS.fuel, JSON.stringify(fuelData));
        localStorage.setItem(STORAGE_KEYS.loads, JSON.stringify(loadsData));
        localStorage.setItem(STORAGE_KEYS.report, JSON.stringify(reportData));
        localStorage.setItem(STORAGE_KEYS.trip, JSON.stringify(tripData));
    } catch (e) {
        console.error('Error saving data:', e);
        showToast('Storage limit reached. Consider clearing old data.', 'error');
    }
}

// ===== Auto-load from Excel data files =====
function loadFromDataFiles() {
    var dataFiles = [
        { url: 'data/Fuel.xlsx', type: 'fuel' },
        { url: 'data/Loads.xlsx', type: 'loads' },
        { url: 'data/Report.xlsx', type: 'report' },
        { url: 'data/LocalTripsheet.xlsx', type: 'trip' }
    ];
    var loaded = 0;
    var total = dataFiles.length;

    dataFiles.forEach(function(fileInfo) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', fileInfo.url, true);
        xhr.responseType = 'arraybuffer';
        xhr.onload = function() {
            if (xhr.status === 200) {
                try {
                    var data = new Uint8Array(xhr.response);
                    var workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    if (fileInfo.type === 'fuel') {
                        parseFuelWorkbook(workbook, fileInfo.url);
                    } else if (fileInfo.type === 'loads') {
                        parseLoadsWorkbook(workbook, fileInfo.url);
                    } else if (fileInfo.type === 'report') {
                        parseReportWorkbook(workbook, fileInfo.url);
                    } else if (fileInfo.type === 'trip') {
                        parseTripWorkbook(workbook, fileInfo.url);
                    }
                } catch (err) {
                    console.warn('Could not parse ' + fileInfo.url + ':', err.message);
                }
            }
            loaded++;
            if (loaded === total) {
                renderAll();
                if (fuelData.length > 0 || loadsData.length > 0 || reportData.length > 0) {
                    showToast('Data loaded from Excel files', 'success');
                }
            }
        };
        xhr.onerror = function() {
            console.warn('Could not fetch ' + fileInfo.url);
            loaded++;
            if (loaded === total) renderAll();
        };
        xhr.send();
    });
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

    document.getElementById('logoutBtn').addEventListener('click', function(e) {
        e.preventDefault();
        sessionStorage.removeItem('tms_user');
        window.location.href = 'index.html';
    });

    // Navbar dropdown toggle
    var navUserBtn = document.getElementById('navUserBtn');
    var navDropdownMenu = document.getElementById('navDropdownMenu');
    navUserBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        navDropdownMenu.classList.toggle('show');
    });
    document.addEventListener('click', function() {
        navDropdownMenu.classList.remove('show');
    });

    // Settings link in dropdown
    document.getElementById('navSettingsBtn').addEventListener('click', function(e) {
        e.preventDefault();
        navDropdownMenu.classList.remove('show');
        document.querySelectorAll('.sidebar-nav li').forEach(function(el) { el.classList.remove('active'); });
        document.querySelectorAll('.tab-content').forEach(function(el) { el.classList.remove('active'); });
        document.getElementById('tab-settings').classList.add('active');
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
            r.driver.split('/').forEach(function(d) { driverSet.add(d.trim()); });
        }
        if (r.truck) truckSet.add(r.truck);
        if (r.trailer) trailerSet.add(r.trailer);
    });
    tripData.forEach(function(r) {
        if (r.driverName) driverSet.add(r.driverName);
        if (r.truck) truckSet.add(r.truck);
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

    filteredTrip = tripData.filter(function(r) {
        if (startDate && r.day && r.day < startDate) return false;
        if (endDate && r.day && r.day > endDate) return false;
        if (driver && r.driverName && r.driverName.indexOf(driver) === -1) return false;
        if (truck && r.truck !== truck) return false;
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
    filteredTrip = [...tripData];
    renderAll();
    showToast('Filters cleared');
}

// ===== Rendering =====
function renderAll() {
    populateFilterDropdowns();
    renderFuelTable();
    renderLoadsTable();
    renderTripTable();
    renderReport();
    renderStoredFiles();
}

function renderFuelTable() {
    var tbody = document.querySelector('#fuelTable tbody');
    tbody.innerHTML = '';
    var totalQty = 0;
    var totalAmt = 0;

    if (filteredFuel.length === 0) {
        tbody.innerHTML = '<tr><td colspan="18" class="empty-state">No fuel records found. Upload a fuel file to get started.</td></tr>';
    } else {
        filteredFuel.forEach(function(r) {
            var realIdx = fuelData.indexOf(r);
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + esc(r.cardNum) + '</td>' +
                '<td>' + formatDate(r.tranDate) + '</td>' +
                '<td>' + esc(r.tranTime) + '</td>' +
                '<td>' + esc(r.invoice) + '</td>' +
                '<td>' + esc(r.unit) + '</td>' +
                '<td>' + esc(r.driverName) + '</td>' +
                '<td class="amount-cell">' + esc(r.odometer) + '</td>' +
                '<td>' + esc(r.locationName) + '</td>' +
                '<td>' + esc(r.city) + '</td>' +
                '<td>' + esc(r.state) + '</td>' +
                '<td class="amount-cell">' + num(r.fees) + '</td>' +
                '<td>' + esc(r.item) + '</td>' +
                '<td class="amount-cell">$' + num(r.unitPrice) + '</td>' +
                '<td class="amount-cell">' + num(r.qty) + '</td>' +
                '<td class="amount-cell">$' + num(r.amt) + '</td>' +
                '<td>' + esc(r.db) + '</td>' +
                '<td>' + esc(r.currency) + '</td>' +
                '<td class="actions-cell"><button class="btn-edit" data-type="fuel" data-idx="' + realIdx + '">Edit</button><button class="btn-delete" data-type="fuel" data-idx="' + realIdx + '">Delete</button></td>';
            tbody.appendChild(tr);
            totalQty += (parseFloat(r.qty) || 0);
            totalAmt += (parseFloat(r.amt) || 0);
        });
    }

    document.getElementById('fuelTotalQty').innerHTML = '<strong>' + totalQty.toFixed(2) + '</strong>';
    document.getElementById('fuelTotalAmt').innerHTML = '<strong>$' + totalAmt.toFixed(2) + '</strong>';

    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            deleteRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
    tbody.querySelectorAll('.btn-edit').forEach(function(btn) {
        btn.addEventListener('click', function() {
            editRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
}

function renderLoadsTable() {
    var tbody = document.querySelector('#loadsTable tbody');
    tbody.innerHTML = '';

    if (filteredLoads.length === 0) {
        tbody.innerHTML = '<tr><td colspan="15" class="empty-state">No load records found. Upload a loads file to get started.</td></tr>';
    } else {
        filteredLoads.forEach(function(r) {
            var realIdx = loadsData.indexOf(r);
            var tr = document.createElement('tr');
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
                '<td class="actions-cell"><button class="btn-edit" data-type="loads" data-idx="' + realIdx + '">Edit</button><button class="btn-delete" data-type="loads" data-idx="' + realIdx + '">Delete</button></td>';
            tbody.appendChild(tr);
        });
    }

    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            deleteRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
    tbody.querySelectorAll('.btn-edit').forEach(function(btn) {
        btn.addEventListener('click', function() {
            editRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
}

function renderTripTable() {
    var tbody = document.querySelector('#tripTable tbody');
    tbody.innerHTML = '';
    var totalHours = 0;

    if (filteredTrip.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="empty-state">No trip sheet records found.</td></tr>';
    } else {
        filteredTrip.forEach(function(r) {
            var realIdx = tripData.indexOf(r);
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + esc(r.driverName) + '</td>' +
                '<td>' + esc(r.truck) + '</td>' +
                '<td>' + formatDate(r.day) + '</td>' +
                '<td>' + esc(r.startTime) + '</td>' +
                '<td>' + esc(r.endTime) + '</td>' +
                '<td class="amount-cell">' + num(r.totalHours) + '</td>' +
                '<td>' + (r.offDutyDay ? 'Yes' : '') + '</td>' +
                '<td>' + esc(r.destination) + '</td>' +
                '<td class="actions-cell"><button class="btn-edit" data-type="trip" data-idx="' + realIdx + '">Edit</button><button class="btn-delete" data-type="trip" data-idx="' + realIdx + '">Delete</button></td>';
            tbody.appendChild(tr);
            totalHours += (parseFloat(r.totalHours) || 0);
        });
    }

    document.getElementById('tripTotalHours').innerHTML = '<strong>' + totalHours.toFixed(2) + '</strong>';

    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            deleteRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
    tbody.querySelectorAll('.btn-edit').forEach(function(btn) {
        btn.addEventListener('click', function() {
            editRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
}

// Helper: get the Saturday end-of-week for a given date
function getWeekEndSaturday(d) {
    var day = d.getDay(); // 0=Sun, 6=Sat
    var diff = 6 - day;
    var sat = new Date(d);
    sat.setDate(sat.getDate() + diff);
    return sat;
}

// Helper: group daily report records into weekly buckets (Sun 00:05 - Sat 23:55)
function groupByWeek(records) {
    if (!records.length) return [];

    // Sort by date
    var sorted = records.slice().sort(function(a, b) {
        return (a.date || '').localeCompare(b.date || '');
    });

    // Determine month boundaries from the data
    var firstDate = new Date(sorted[0].date + 'T00:00:00');
    var lastDate = new Date(sorted[sorted.length - 1].date + 'T00:00:00');

    var weeks = [];
    var currentStart = new Date(firstDate);

    while (currentStart <= lastDate) {
        var weekEnd = getWeekEndSaturday(currentStart);
        // Cap week end at last date of data
        if (weekEnd > lastDate) weekEnd = new Date(lastDate);

        var startStr = currentStart.toISOString().slice(0, 10);
        var endStr = weekEnd.toISOString().slice(0, 10);

        var weekRecords = sorted.filter(function(r) {
            return r.date >= startStr && r.date <= endStr;
        });

        var trucks = new Set();
        var drivers = new Set();
        var totalMiles = 0;
        var startOdoMin = Infinity;
        var endOdoMax = 0;

        weekRecords.forEach(function(r) {
            if (r.truck) r.truck.split(',').forEach(function(t) { trucks.add(t.trim()); });
            if (r.driverName) drivers.add(r.driverName);
            totalMiles += (parseFloat(r.totalMiles) || 0);
            var so = parseFloat(r.startOdo) || 0;
            var eo = parseFloat(r.endOdo) || 0;
            if (so > 0 && so < startOdoMin) startOdoMin = so;
            if (eo > endOdoMax) endOdoMax = eo;
        });

        var truckList = Array.from(trucks);
        var reportName = truckList.length ? 'Use Truck - ' + truckList.join(',') : '';

        weeks.push({
            reportName: reportName,
            driver: Array.from(drivers).join(', '),
            date: endStr,
            state: 'LL',
            startOdo: startOdoMin === Infinity ? 0 : startOdoMin,
            endOdo: endOdoMax,
            miles: totalMiles
        });

        // Move to next Sunday
        var nextSun = new Date(weekEnd);
        nextSun.setDate(nextSun.getDate() + 1);
        currentStart = nextSun;
    }

    return weeks;
}

function renderReport() {
    var milesBody = document.querySelector('#reportMilesTable tbody');
    milesBody.innerHTML = '';
    var totalMiles = 0;

    if (filteredReport.length === 0) {
        milesBody.innerHTML = '<tr><td colspan="7" class="empty-state">No miles data. Upload a report file with miles data.</td></tr>';
    } else {
        // Determine driver name for the title
        var driverNames = new Set();
        filteredReport.forEach(function(r) { if (r.driverName) driverNames.add(r.driverName); });
        var driverLabel = driverNames.size ? Array.from(driverNames).join(', ') : '';
        if (driverLabel) {
            document.getElementById('rptWeeklyTitle').textContent = 'Weekly - ' + driverLabel + ' Miles Service Information';
        }

        var weeks = groupByWeek(filteredReport);
        weeks.forEach(function(w) {
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td>' + esc(w.reportName) + '</td>' +
                '<td>' + esc(w.driver) + '</td>' +
                '<td>' + formatDate(w.date) + '</td>' +
                '<td>' + esc(w.state) + '</td>' +
                '<td class="odo-col amount-cell">' + (w.startOdo ? w.startOdo.toLocaleString() : '') + '</td>' +
                '<td class="odo-col amount-cell">' + (w.endOdo ? w.endOdo.toLocaleString() : '') + '</td>' +
                '<td class="amount-cell">' + w.miles.toFixed(2) + '</td>';
            milesBody.appendChild(tr);
            totalMiles += w.miles;
        });
    }

    document.getElementById('rptMilesTotal').innerHTML = '<strong>' + totalMiles.toFixed(2) + '</strong>';

    // Total Earnings at $0.50/mile
    var totalEarnings = totalMiles * 0.50;
    document.getElementById('rptTotalEarnings').textContent = '$' + totalEarnings.toFixed(2);
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

// ===== Delete Record =====
function deleteRecord(type, idx) {
    if (!confirm('Are you sure you want to delete this record?')) return;

    if (type === 'fuel') {
        fuelData.splice(idx, 1);
        filteredFuel = [...fuelData];
    } else if (type === 'loads') {
        loadsData.splice(idx, 1);
        filteredLoads = [...loadsData];
    } else if (type === 'report') {
        reportData.splice(idx, 1);
        filteredReport = [...reportData];
    } else if (type === 'trip') {
        tripData.splice(idx, 1);
        filteredTrip = [...tripData];
    }

    saveData();
    applyFilters();
    showToast('Record deleted. Click "Save to Excel" to update the Excel file.', 'success');
}

// ===== Edit Record =====
function editRecord(type, idx) {
    if (type === 'fuel') {
        openEditFuelModal(idx);
    } else if (type === 'loads') {
        openEditLoadModal(idx);
    } else if (type === 'trip') {
        openEditTripModal(idx);
    }
}

function openEditFuelModal(idx) {
    var r = fuelData[idx];
    if (!r) return;
    document.getElementById('modalTitle').textContent = 'Edit Fuel Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>Card #</label><input type="text" id="mFuelCard" value="' + escAttr(r.cardNum) + '"></div>' +
            '<div class="form-group"><label>Date</label><input type="date" id="mFuelDate" value="' + escAttr(r.tranDate) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Time</label><input type="time" id="mFuelTime" value="' + escAttr(r.tranTime) + '"></div>' +
            '<div class="form-group"><label>Invoice</label><input type="text" id="mFuelInvoice" value="' + escAttr(r.invoice) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Unit/Truck</label><input type="text" id="mFuelUnit" value="' + escAttr(r.unit) + '"></div>' +
            '<div class="form-group"><label>Driver Name</label><input type="text" id="mFuelDriver" value="' + escAttr(r.driverName) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Odometer</label><input type="number" id="mFuelOdometer" value="' + escAttr(r.odometer) + '"></div>' +
            '<div class="form-group"><label>Location</label><input type="text" id="mFuelLocation" value="' + escAttr(r.locationName) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>City</label><input type="text" id="mFuelCity" value="' + escAttr(r.city) + '"></div>' +
            '<div class="form-group"><label>State</label><input type="text" id="mFuelState" value="' + escAttr(r.state) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Fees</label><input type="number" step="0.01" id="mFuelFees" value="' + escAttr(r.fees) + '"></div>' +
            '<div class="form-group"><label>Item</label><input type="text" id="mFuelItem" value="' + escAttr(r.item) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Unit Price</label><input type="number" step="0.001" id="mFuelPrice" value="' + escAttr(r.unitPrice) + '"></div>' +
            '<div class="form-group"><label>Qty (Gallons)</label><input type="number" step="0.01" id="mFuelQty" value="' + escAttr(r.qty) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Amount</label><input type="number" step="0.01" id="mFuelAmt" value="' + escAttr(r.amt) + '"></div>' +
            '<div class="form-group"><label>Currency</label><input type="text" id="mFuelCurrency" value="' + escAttr(r.currency) + '"></div>' +
        '</div>';

    document.getElementById('modalSave').onclick = function() {
        fuelData[idx] = {
            cardNum: document.getElementById('mFuelCard').value,
            tranDate: document.getElementById('mFuelDate').value,
            tranTime: document.getElementById('mFuelTime').value,
            invoice: document.getElementById('mFuelInvoice').value,
            unit: document.getElementById('mFuelUnit').value,
            driverName: document.getElementById('mFuelDriver').value,
            odometer: parseFloat(document.getElementById('mFuelOdometer').value) || 0,
            locationName: document.getElementById('mFuelLocation').value,
            city: document.getElementById('mFuelCity').value,
            state: document.getElementById('mFuelState').value,
            fees: parseFloat(document.getElementById('mFuelFees').value) || 0,
            item: document.getElementById('mFuelItem').value,
            unitPrice: parseFloat(document.getElementById('mFuelPrice').value) || 0,
            qty: parseFloat(document.getElementById('mFuelQty').value) || 0,
            amt: parseFloat(document.getElementById('mFuelAmt').value) || 0,
            db: r.db || '',
            currency: document.getElementById('mFuelCurrency').value
        };
        filteredFuel = [...fuelData];
        saveData();
        renderAll();
        closeModal();
        showToast('Fuel record updated. Click "Save to Excel" to update the Excel file.', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

function openEditLoadModal(idx) {
    var r = loadsData[idx];
    if (!r) return;
    document.getElementById('modalTitle').textContent = 'Edit Load Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>InvoiceID</label><input type="text" id="mLoadInvoiceId" value="' + escAttr(r.invoiceId) + '"></div>' +
            '<div class="form-group"><label>Load #</label><input type="text" id="mLoadNum" value="' + escAttr(r.loadNum) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Broker</label><input type="text" id="mLoadBroker" value="' + escAttr(r.broker) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Pick Date</label><input type="date" id="mLoadPickDate" value="' + escAttr(r.pickDate) + '"></div>' +
            '<div class="form-group"><label>Pickup Location</label><input type="text" id="mLoadPickup" value="' + escAttr(r.pickup) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Drop Date</label><input type="date" id="mLoadDropDate" value="' + escAttr(r.dropDate) + '"></div>' +
            '<div class="form-group"><label>Dropoff Location</label><input type="text" id="mLoadDropoff" value="' + escAttr(r.dropoff) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Driver</label><input type="text" id="mLoadDriver" value="' + escAttr(r.driver) + '"></div>' +
            '<div class="form-group"><label>Truck</label><input type="text" id="mLoadTruck" value="' + escAttr(r.truck) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Trailer</label><input type="text" id="mLoadTrailer" value="' + escAttr(r.trailer) + '"></div>' +
        '</div>';

    document.getElementById('modalSave').onclick = function() {
        loadsData[idx] = {
            invoiceId: document.getElementById('mLoadInvoiceId').value,
            loadNum: document.getElementById('mLoadNum').value,
            broker: document.getElementById('mLoadBroker').value,
            pickDate: document.getElementById('mLoadPickDate').value,
            pickup: document.getElementById('mLoadPickup').value,
            dropDate: document.getElementById('mLoadDropDate').value,
            dropoff: document.getElementById('mLoadDropoff').value,
            driver: document.getElementById('mLoadDriver').value,
            truck: document.getElementById('mLoadTruck').value,
            trailer: document.getElementById('mLoadTrailer').value
        };
        filteredLoads = [...loadsData];
        saveData();
        renderAll();
        closeModal();
        showToast('Load record updated. Click "Save to Excel" to update the Excel file.', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

function openEditTripModal(idx) {
    var r = tripData[idx];
    if (!r) return;
    document.getElementById('modalTitle').textContent = 'Edit Trip Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>Driver Name</label><input type="text" id="mTripDriver" value="' + escAttr(r.driverName) + '"></div>' +
            '<div class="form-group"><label>Truck</label><input type="text" id="mTripTruck" value="' + escAttr(r.truck) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Day</label><input type="date" id="mTripDay" value="' + escAttr(r.day) + '"></div>' +
            '<div class="form-group"><label>Destination City/State</label><input type="text" id="mTripDest" value="' + escAttr(r.destination) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Start Time</label><input type="time" id="mTripStart" value="' + escAttr(r.startTime) + '"></div>' +
            '<div class="form-group"><label>End Time</label><input type="time" id="mTripEnd" value="' + escAttr(r.endTime) + '"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Total Hours</label><input type="number" step="0.01" id="mTripHours" value="' + escAttr(r.totalHours) + '"></div>' +
            '<div class="form-group"><label>Off Duty Day</label><select id="mTripOffDuty"><option value="false"' + (!r.offDutyDay ? ' selected' : '') + '>No</option><option value="true"' + (r.offDutyDay ? ' selected' : '') + '>Yes</option></select></div>' +
        '</div>';

    document.getElementById('modalSave').onclick = function() {
        tripData[idx] = {
            driverName: document.getElementById('mTripDriver').value,
            truck: document.getElementById('mTripTruck').value,
            day: document.getElementById('mTripDay').value,
            startTime: document.getElementById('mTripStart').value,
            endTime: document.getElementById('mTripEnd').value,
            totalHours: parseFloat(document.getElementById('mTripHours').value) || 0,
            offDutyDay: document.getElementById('mTripOffDuty').value === 'true',
            destination: document.getElementById('mTripDest').value
        };
        filteredTrip = [...tripData];
        saveData();
        renderAll();
        closeModal();
        showToast('Trip record updated. Click "Save to Excel" to update the Excel file.', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

function openAddTripModal() {
    document.getElementById('modalTitle').textContent = 'Add Trip Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>Driver Name</label><input type="text" id="mTripDriver" value=""></div>' +
            '<div class="form-group"><label>Truck</label><input type="text" id="mTripTruck" value=""></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Day</label><input type="date" id="mTripDay" value=""></div>' +
            '<div class="form-group"><label>Destination City/State</label><input type="text" id="mTripDest" value=""></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Start Time</label><input type="time" id="mTripStart" value=""></div>' +
            '<div class="form-group"><label>End Time</label><input type="time" id="mTripEnd" value=""></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Total Hours</label><input type="number" step="0.01" id="mTripHours" value=""></div>' +
            '<div class="form-group"><label>Off Duty Day</label><select id="mTripOffDuty"><option value="false" selected>No</option><option value="true">Yes</option></select></div>' +
        '</div>';

    document.getElementById('modalSave').onclick = function() {
        tripData.push({
            driverName: document.getElementById('mTripDriver').value,
            truck: document.getElementById('mTripTruck').value,
            day: document.getElementById('mTripDay').value,
            startTime: document.getElementById('mTripStart').value,
            endTime: document.getElementById('mTripEnd').value,
            totalHours: parseFloat(document.getElementById('mTripHours').value) || 0,
            offDutyDay: document.getElementById('mTripOffDuty').value === 'true',
            destination: document.getElementById('mTripDest').value
        });
        filteredTrip = [...tripData];
        saveData();
        renderAll();
        closeModal();
        showToast('Trip record added. Click "Save to Excel" to update the Excel file.', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

// ===== Save to Excel =====
function saveToExcel(type) {
    var ws, wb, filename;

    if (type === 'fuel') {
        var fuelHeaders = ['Card #', 'Tran Date', 'Trans. Time', 'Invoice', 'Unit', 'Driver Name', 'Odometer', 'Location Name', 'City', 'State/Prov', 'Fees', 'Item', 'Unit Price', 'Qty', 'Amt', 'DB', 'Currency'];
        var fuelRows = [fuelHeaders];
        fuelData.forEach(function(r) {
            fuelRows.push([
                r.cardNum, r.tranDate, r.tranTime, r.invoice, r.unit, r.driverName,
                r.odometer, r.locationName, r.city, r.state, r.fees, r.item,
                r.unitPrice, r.qty, r.amt, r.db, r.currency
            ]);
        });
        ws = XLSX.utils.aoa_to_sheet(fuelRows);
        ws['!cols'] = fuelHeaders.map(function() { return { wch: 14 }; });
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Fuel');
        filename = 'Fuel.xlsx';
    } else if (type === 'loads') {
        var loadsHeaders = ['InvoiceID', 'Load #', 'Broker', 'Pick Date', 'Pickup', 'Drop Date', 'Dropoff', 'Driver', 'TruckName', 'Trailer'];
        var loadsRows = [loadsHeaders];
        loadsData.forEach(function(r) {
            loadsRows.push([
                r.invoiceId, r.loadNum, r.broker, r.pickDate, r.pickup, r.dropDate,
                r.dropoff, r.driver, r.truck, r.trailer
            ]);
        });
        ws = XLSX.utils.aoa_to_sheet(loadsRows);
        ws['!cols'] = loadsHeaders.map(function() { return { wch: 16 }; });
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Loads');
        filename = 'Loads.xlsx';
    } else if (type === 'report') {
        var weeks = groupByWeek(filteredReport);
        var reportHeaders = ['Report Name', 'Driver ID', 'Date', 'State', 'Miles'];
        var reportRows = [reportHeaders];
        weeks.forEach(function(w) {
            reportRows.push([w.reportName, w.driver, w.date, w.state, w.miles]);
        });
        ws = XLSX.utils.aoa_to_sheet(reportRows);
        ws['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 12 }];
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Weekly Miles');
        filename = 'Weekly_Miles_Report.xlsx';
    } else if (type === 'trip') {
        var tripHeaders = ['Driver Name', 'Truck', 'Day', 'Start Time', 'End Time', 'Total Hours', 'Off Duty Day', 'Destination City/State'];
        var tripRows = [tripHeaders];
        tripData.forEach(function(r) {
            tripRows.push([
                r.driverName || '', r.truck || '', r.day, r.startTime, r.endTime,
                r.totalHours, r.offDutyDay ? 'Yes' : '', r.destination || ''
            ]);
        });
        ws = XLSX.utils.aoa_to_sheet(tripRows);
        ws['!cols'] = tripHeaders.map(function() { return { wch: 18 }; });
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'TripSheet');
        filename = 'LocalTripsheet.xlsx';
    }

    if (wb) {
        XLSX.writeFile(wb, filename);
        showToast('Saved ' + filename + ' - replace the file in data/ folder and push to update.', 'success');
    }
}

// ===== Modal (Add Record) =====
function setupModal() {
    document.getElementById('modalClose').addEventListener('click', closeModal);
    document.getElementById('modalCancel').addEventListener('click', closeModal);
    document.getElementById('modalOverlay').addEventListener('click', function(e) {
        if (e.target === this) closeModal();
    });

    document.getElementById('addFuelBtn').addEventListener('click', function() {
        openAddFuelModal();
    });
    document.getElementById('addLoadBtn').addEventListener('click', function() {
        openAddLoadModal();
    });
    document.getElementById('addTripBtn').addEventListener('click', function() {
        openAddTripModal();
    });
}

function closeModal() {
    document.getElementById('modalOverlay').classList.remove('active');
}

function openAddFuelModal() {
    document.getElementById('modalTitle').textContent = 'Add Fuel Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>Card #</label><input type="text" id="mFuelCard"></div>' +
            '<div class="form-group"><label>Date</label><input type="date" id="mFuelDate"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Time</label><input type="time" id="mFuelTime"></div>' +
            '<div class="form-group"><label>Invoice</label><input type="text" id="mFuelInvoice"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Unit/Truck</label><input type="text" id="mFuelUnit"></div>' +
            '<div class="form-group"><label>Driver Name</label><input type="text" id="mFuelDriver"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Location</label><input type="text" id="mFuelLocation"></div>' +
            '<div class="form-group"><label>City</label><input type="text" id="mFuelCity"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>State</label><input type="text" id="mFuelState"></div>' +
            '<div class="form-group"><label>Item</label><input type="text" id="mFuelItem" placeholder="ULSD, DEFD, etc."></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Unit Price</label><input type="number" step="0.001" id="mFuelPrice"></div>' +
            '<div class="form-group"><label>Qty (Gallons)</label><input type="number" step="0.01" id="mFuelQty"></div>' +
        '</div>' +
        '<div class="form-group"><label>Amount</label><input type="number" step="0.01" id="mFuelAmt"></div>';

    document.getElementById('modalSave').onclick = function() {
        var record = {
            cardNum: document.getElementById('mFuelCard').value,
            tranDate: document.getElementById('mFuelDate').value,
            tranTime: document.getElementById('mFuelTime').value,
            invoice: document.getElementById('mFuelInvoice').value,
            unit: document.getElementById('mFuelUnit').value,
            driverName: document.getElementById('mFuelDriver').value,
            odometer: 0,
            locationName: document.getElementById('mFuelLocation').value,
            city: document.getElementById('mFuelCity').value,
            state: document.getElementById('mFuelState').value,
            fees: 0,
            item: document.getElementById('mFuelItem').value,
            unitPrice: parseFloat(document.getElementById('mFuelPrice').value) || 0,
            qty: parseFloat(document.getElementById('mFuelQty').value) || 0,
            amt: parseFloat(document.getElementById('mFuelAmt').value) || 0,
            db: '',
            currency: 'USD/Gallons'
        };
        if (!record.tranDate) {
            showToast('Date is required', 'error');
            return;
        }
        fuelData.push(record);
        filteredFuel = [...fuelData];
        saveData();
        renderAll();
        closeModal();
        showToast('Fuel record added', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

function openAddLoadModal() {
    document.getElementById('modalTitle').textContent = 'Add Load Record';
    document.getElementById('modalBody').innerHTML =
        '<div class="form-row">' +
            '<div class="form-group"><label>InvoiceID</label><input type="text" id="mLoadInvoiceId"></div>' +
            '<div class="form-group"><label>Load #</label><input type="text" id="mLoadNum"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Broker</label><input type="text" id="mLoadBroker"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Pick Date</label><input type="date" id="mLoadPickDate"></div>' +
            '<div class="form-group"><label>Pickup Location</label><input type="text" id="mLoadPickup"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Drop Date</label><input type="date" id="mLoadDropDate"></div>' +
            '<div class="form-group"><label>Dropoff Location</label><input type="text" id="mLoadDropoff"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Driver</label><input type="text" id="mLoadDriver"></div>' +
            '<div class="form-group"><label>Truck</label><input type="text" id="mLoadTruck"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>Trailer</label><input type="text" id="mLoadTrailer"></div>' +
        '</div>';

    document.getElementById('modalSave').onclick = function() {
        var record = {
            invoiceId: document.getElementById('mLoadInvoiceId').value,
            loadNum: document.getElementById('mLoadNum').value,
            broker: document.getElementById('mLoadBroker').value,
            pickDate: document.getElementById('mLoadPickDate').value,
            pickup: document.getElementById('mLoadPickup').value,
            dropDate: document.getElementById('mLoadDropDate').value,
            dropoff: document.getElementById('mLoadDropoff').value,
            driver: document.getElementById('mLoadDriver').value,
            truck: document.getElementById('mLoadTruck').value,
            trailer: document.getElementById('mLoadTrailer').value
        };
        if (!record.pickDate) {
            showToast('Pick Date is required', 'error');
            return;
        }
        loadsData.push(record);
        filteredLoads = [...loadsData];
        saveData();
        renderAll();
        closeModal();
        showToast('Load record added', 'success');
    };

    document.getElementById('modalOverlay').classList.add('active');
}

// ===== Upload Handlers =====
function setupUpload() {
    // Combined file upload
    var combinedInput = document.getElementById('combinedFileInput');
    combinedInput.addEventListener('change', function() {
        document.getElementById('combinedFileName').textContent = this.files[0] ? this.files[0].name : 'No file chosen';
        document.getElementById('uploadCombinedBtn').disabled = !this.files[0];
    });
    document.getElementById('uploadCombinedBtn').addEventListener('click', function() {
        var file = combinedInput.files[0];
        if (!file) return;
        parseCombinedFile(file);
    });

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

    // Trip Sheet file
    var tripInput = document.getElementById('tripFileInput');
    tripInput.addEventListener('change', function() {
        document.getElementById('tripFileName').textContent = this.files[0] ? this.files[0].name : 'No file chosen';
        document.getElementById('uploadTripBtn').disabled = !this.files[0];
    });
    document.getElementById('uploadTripBtn').addEventListener('click', function() {
        var file = tripInput.files[0];
        if (!file) return;
        parseExcelFile(file, 'trip');
    });
}

// ===== Combined File Parser =====
function parseCombinedFile(file) {
    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array', cellDates: true });
            var totalRecords = { fuel: 0, loads: 0, report: 0 };

            var fuelSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('fuel') !== -1;
            });
            if (fuelSheet) {
                totalRecords.fuel = parseFuelWorkbook(workbook, file.name, fuelSheet) || 0;
            }

            var loadsSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('load') !== -1;
            });
            if (loadsSheet) {
                totalRecords.loads = parseLoadsWorkbook(workbook, file.name, loadsSheet) || 0;
            }

            var reportSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('report') !== -1;
            });
            if (reportSheet) {
                totalRecords.report = parseReportWorkbook(workbook, file.name, reportSheet) || 0;
            }

            var parts = [];
            if (totalRecords.fuel > 0) parts.push(totalRecords.fuel + ' fuel');
            if (totalRecords.loads > 0) parts.push(totalRecords.loads + ' loads');
            if (totalRecords.report > 0) parts.push(totalRecords.report + ' report');
            if (parts.length === 0) {
                showToast('No data sheets found. Expected sheets named Fuel, Loads, or Report.', 'error');
            } else {
                showToast('Loaded from ' + file.name + ': ' + parts.join(', ') + ' records', 'success');
            }
        } catch (err) {
            console.error('Parse error:', err);
            showToast('Error parsing file: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
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
            } else if (type === 'trip') {
                parseTripWorkbook(workbook, file.name);
            }
        } catch (err) {
            console.error('Parse error:', err);
            showToast('Error parsing file: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// ===== Fuel Parser =====
function parseFuelWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('fuel') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row by looking for known fuel column names
    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (!row) continue;
        var rowStr = row.map(function(c) { return String(c).toLowerCase(); }).join('|');
        if (rowStr.indexOf('card') !== -1 && rowStr.indexOf('date') !== -1) {
            headerIdx = i;
            break;
        }
        if (rowStr.indexOf('tran date') !== -1) {
            headerIdx = i;
            break;
        }
    }

    if (headerIdx === -1) {
        if (!sheetNameOverride) showToast('Could not find fuel data headers in file', 'error');
        return 0;
    }

    var headers = rows[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });

    // Build column map using exact position-based mapping
    // Expected order: Card#, TranDate, Trans.Time, Invoice, Unit, DriverName, Odometer, LocationName, City, State, Fees, Item, UnitPrice, Qty, Amt, DB, Currency
    var colMap = {};
    for (var c = 0; c < headers.length; c++) {
        var h = headers[c];
        if (h.indexOf('card') !== -1 && !colMap.cardNum) colMap.cardNum = c;
        else if ((h.indexOf('tran date') !== -1 || h === 'date') && !colMap.tranDate) colMap.tranDate = c;
        else if ((h.indexOf('time') !== -1) && h.indexOf('date') === -1 && !colMap.tranTime) colMap.tranTime = c;
        else if (h.indexOf('invoice') !== -1 && !colMap.invoice) colMap.invoice = c;
        else if (h === 'unit' && !colMap.unit) colMap.unit = c;
        else if (h.indexOf('driver') !== -1 && !colMap.driverName) colMap.driverName = c;
        else if ((h.indexOf('odometer') !== -1 || h.indexOf('odo') !== -1) && !colMap.odometer) colMap.odometer = c;
        else if (h.indexOf('location') !== -1 && !colMap.locationName) colMap.locationName = c;
        else if (h.indexOf('city') !== -1 && !colMap.city) colMap.city = c;
        else if ((h.indexOf('state') !== -1 || h.indexOf('prov') !== -1) && !colMap.state) colMap.state = c;
        else if ((h.indexOf('fee') !== -1) && !colMap.fees) colMap.fees = c;
        else if (h === 'item' && !colMap.item) colMap.item = c;
        else if (h.indexOf('unit price') !== -1 || h.indexOf('price') !== -1 && !colMap.unitPrice) colMap.unitPrice = c;
        else if ((h.indexOf('qty') !== -1 || h.indexOf('quantity') !== -1) && !colMap.qty) colMap.qty = c;
        else if ((h.indexOf('amt') !== -1 || h.indexOf('amount') !== -1) && !colMap.amt) colMap.amt = c;
        else if (h === 'db' && !colMap.db) colMap.db = c;
        else if (h.indexOf('currency') !== -1 && !colMap.currency) colMap.currency = c;
    }

    var newData = [];
    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row) continue;

        // Check if row has actual data (not empty)
        var dateVal = colMap.tranDate !== undefined ? row[colMap.tranDate] : null;
        var cardVal = colMap.cardNum !== undefined ? row[colMap.cardNum] : null;
        if (!dateVal && !cardVal) continue;

        // Skip rows where the "date" cell looks like a header or is empty text
        if (typeof dateVal === 'string' && dateVal.toLowerCase().indexOf('date') !== -1) continue;

        newData.push({
            cardNum: safeVal(row, colMap.cardNum),
            tranDate: excelDateToStr(safeVal(row, colMap.tranDate)),
            tranTime: formatExcelTime(safeVal(row, colMap.tranTime)),
            invoice: safeVal(row, colMap.invoice),
            unit: safeVal(row, colMap.unit),
            driverName: safeVal(row, colMap.driverName),
            odometer: safeVal(row, colMap.odometer) || 0,
            locationName: safeVal(row, colMap.locationName),
            city: safeVal(row, colMap.city),
            state: safeVal(row, colMap.state),
            fees: safeVal(row, colMap.fees) || 0,
            item: safeVal(row, colMap.item),
            unitPrice: safeVal(row, colMap.unitPrice) || 0,
            qty: safeVal(row, colMap.qty) || 0,
            amt: safeVal(row, colMap.amt) || 0,
            db: safeVal(row, colMap.db),
            currency: safeVal(row, colMap.currency)
        });
    }

    fuelData = fuelData.concat(newData);
    filteredFuel = [...fuelData];
    saveData();
    saveFileRecord(fileName, 'Fuel', newData.length);
    renderAll();
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' fuel records from ' + fileName, 'success');
    }
    return newData.length;
}

// ===== Loads Parser =====
function parseLoadsWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('load') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row
    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (!row) continue;
        var rowStr = row.map(function(c) { return String(c).toLowerCase(); }).join('|');
        if (rowStr.indexOf('pick date') !== -1 || rowStr.indexOf('pickup') !== -1 ||
            rowStr.indexOf('invoiceid') !== -1 || rowStr.indexOf('load #') !== -1) {
            headerIdx = i;
            break;
        }
    }

    if (headerIdx === -1) {
        if (!sheetNameOverride) showToast('Could not find load data headers in file', 'error');
        return 0;
    }

    var headers = rows[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });

    // Build column map
    var colMap = {};
    for (var c = 0; c < headers.length; c++) {
        var h = headers[c];
        if ((h.indexOf('invoiceid') !== -1 || h === 'invoice id') && !colMap.invoiceId) colMap.invoiceId = c;
        else if ((h.indexOf('load #') !== -1 || h === 'load') && !colMap.loadNum) colMap.loadNum = c;
        else if (h.indexOf('broker') !== -1 && !colMap.broker) colMap.broker = c;
        else if (h.indexOf('pick date') !== -1 && !colMap.pickDate) colMap.pickDate = c;
        else if (h === 'pickup' && !colMap.pickup) colMap.pickup = c;
        else if (h.indexOf('drop date') !== -1 && !colMap.dropDate) colMap.dropDate = c;
        else if ((h === 'dropoff' || h === 'drop off') && !colMap.dropoff) colMap.dropoff = c;
        else if (h.indexOf('driver') !== -1 && !colMap.driver) colMap.driver = c;
        else if ((h.indexOf('truck') !== -1) && !colMap.truck) colMap.truck = c;
        else if (h.indexOf('trailer') !== -1 && !colMap.trailer) colMap.trailer = c;
    }

    var newData = [];
    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row) continue;

        var pickVal = safeVal(row, colMap.pickDate);
        var invoiceVal = safeVal(row, colMap.invoiceId);
        var loadVal = safeVal(row, colMap.loadNum);
        if (!pickVal && !invoiceVal && !loadVal) continue;

        newData.push({
            invoiceId: invoiceVal || '',
            loadNum: loadVal || '',
            broker: safeVal(row, colMap.broker) || '',
            pickDate: excelDateToStr(pickVal),
            pickup: safeVal(row, colMap.pickup) || '',
            dropDate: excelDateToStr(safeVal(row, colMap.dropDate)),
            dropoff: safeVal(row, colMap.dropoff) || '',
            driver: safeVal(row, colMap.driver) || '',
            truck: safeVal(row, colMap.truck) || '',
            trailer: safeVal(row, colMap.trailer) || ''
        });
    }

    loadsData = loadsData.concat(newData);
    filteredLoads = [...loadsData];
    saveData();
    saveFileRecord(fileName, 'Loads', newData.length);
    renderAll();
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' load records from ' + fileName, 'success');
    }
    return newData.length;
}

// ===== Report Parser =====
function parseReportWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('report') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    var newData = [];
    var headerIdx = -1;
    var driverNameCol = -1;
    var truckCol = -1;
    var dateCol = -1;
    var startOdoCol = -1;
    var endOdoCol = -1;
    var totalMileCol = -1;
    var missingCol = -1;
    var notesCol = -1;

    // Search all rows for the miles data header
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (!row) continue;
        for (var j = 0; j < row.length; j++) {
            var cellStr = String(row[j]).toLowerCase().trim();
            if ((cellStr.indexOf('driver') !== -1) && driverNameCol === -1) driverNameCol = j;
            if ((cellStr.indexOf('truck') !== -1) && truckCol === -1) truckCol = j;
            if ((cellStr === 'date' || cellStr === 'dates') && dateCol === -1) { dateCol = j; headerIdx = i; }
            if (cellStr.indexOf('start odo') !== -1 || cellStr.indexOf('start_odo') !== -1) startOdoCol = j;
            if (cellStr.indexOf('end odo') !== -1 || cellStr.indexOf('end_odo') !== -1) endOdoCol = j;
            if (cellStr.indexOf('total mile') !== -1 || cellStr === 'total miles') totalMileCol = j;
            if (cellStr.indexOf('missing') !== -1) missingCol = j;
            if (cellStr === 'details' || cellStr === 'notes') notesCol = j;
        }
        if (dateCol !== -1 && startOdoCol !== -1) break;
    }

    // If headers found at a low column index, also check the R+ range (index 17+) for a second "Date" column
    if (dateCol !== -1 && dateCol < 17) {
        for (var i = 0; i < Math.min(rows.length, 15); i++) {
            var row = rows[i];
            if (!row) continue;
            for (var j = 17; j < row.length; j++) {
                var cellStr = String(row[j]).toLowerCase().trim();
                if (cellStr === 'date' || cellStr === 'dates') {
                    dateCol = j;
                    headerIdx = i;
                    driverNameCol = -1;
                    truckCol = -1;
                    startOdoCol = -1;
                    endOdoCol = -1;
                    totalMileCol = -1;
                    missingCol = -1;
                    // Re-scan this row from j-2 (driver/truck may be before date) through end
                    for (var k = Math.max(0, j - 5); k < row.length; k++) {
                        var cs = String(row[k]).toLowerCase().trim();
                        if (cs.indexOf('driver') !== -1 && driverNameCol === -1) driverNameCol = k;
                        if (cs.indexOf('truck') !== -1 && truckCol === -1) truckCol = k;
                        if (cs.indexOf('start odo') !== -1) startOdoCol = k;
                        if (cs.indexOf('end odo') !== -1) endOdoCol = k;
                        if (cs.indexOf('total mile') !== -1) totalMileCol = k;
                        if (cs.indexOf('missing') !== -1) missingCol = k;
                    }
                    break;
                }
            }
            if (dateCol >= 17) break;
        }
    }

    // Fallback: check R9:U50 layout without headers
    if (dateCol === -1) {
        for (var i = 7; i < Math.min(rows.length, 12); i++) {
            var row = rows[i];
            if (!row || !row[17]) continue;
            var testDate = excelDateToStr(row[17]);
            if (testDate && /^\d{4}-\d{2}-\d{2}$/.test(testDate)) {
                dateCol = 17;
                headerIdx = i - 1;
                startOdoCol = 18;
                endOdoCol = 19;
                totalMileCol = 20;
                break;
            }
        }
    }

    if (dateCol === -1) {
        if (!sheetNameOverride) showToast('Could not find date/odometer columns in report file', 'error');
        return 0;
    }

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
        if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;

        newData.push({
            driverName: driverNameCol !== -1 && driverNameCol < row.length ? (row[driverNameCol] || '') : '',
            truck: truckCol !== -1 && truckCol < row.length ? (row[truckCol] || '') : '',
            date: dateStr,
            startOdo: row[startOdoCol] || 0,
            endOdo: row[endOdoCol] || 0,
            totalMiles: row[totalMileCol] || 0,
            missingMiles: missingCol < row.length ? (row[missingCol] || 0) : 0,
            notes: notesCol !== -1 && notesCol < row.length ? (row[notesCol] || '') : ''
        });
    }

    reportData = reportData.concat(newData);
    filteredReport = [...reportData];
    saveData();
    saveFileRecord(fileName, 'Report (Miles/Odo)', newData.length);
    renderAll();
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' daily mile records from ' + fileName, 'success');
    }
    return newData.length;
}

// ===== Trip Sheet Parser =====
function parseTripWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('trip') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    var newData = [];
    var headerIdx = -1;
    var colMap = {};

    // Find header row
    for (var i = 0; i < Math.min(rows.length, 10); i++) {
        var row = rows[i];
        if (!row) continue;
        for (var j = 0; j < row.length; j++) {
            var cs = String(row[j]).toLowerCase().trim();
            if (cs === 'day' || cs === 'date') colMap.day = j;
            if (cs.indexOf('start') !== -1 && cs.indexOf('time') !== -1) colMap.startTime = j;
            if (cs.indexOf('end') !== -1 && cs.indexOf('time') !== -1) colMap.endTime = j;
            if (cs.indexOf('total') !== -1 && cs.indexOf('hour') !== -1) colMap.totalHours = j;
            if (cs.indexOf('off duty') !== -1 || cs.indexOf('check box') !== -1) colMap.offDutyDay = j;
            if (cs.indexOf('destination') !== -1 || cs.indexOf('city/state') !== -1 || cs.indexOf('trip') !== -1) colMap.destination = j;
            if (cs.indexOf('driver') !== -1) colMap.driverName = j;
            if (cs.indexOf('truck') !== -1) colMap.truck = j;
        }
        if (colMap.day !== undefined) { headerIdx = i; break; }
    }

    if (headerIdx === -1) {
        if (!sheetNameOverride) showToast('Could not find trip sheet headers', 'error');
        return 0;
    }

    for (var r = headerIdx + 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row) continue;
        var dayVal = colMap.day !== undefined ? row[colMap.day] : '';
        if (!dayVal) continue;

        var dayStr = excelDateToStr(dayVal);
        if (!dayStr) dayStr = String(dayVal);

        var offDuty = colMap.offDutyDay !== undefined ? row[colMap.offDutyDay] : '';
        var isOffDuty = offDuty === true || offDuty === 'Yes' || offDuty === 'yes' || offDuty === 'X' || offDuty === 'x' || offDuty === 1;

        newData.push({
            driverName: colMap.driverName !== undefined ? (row[colMap.driverName] || '') : '',
            truck: colMap.truck !== undefined ? (row[colMap.truck] || '') : '',
            day: dayStr,
            startTime: colMap.startTime !== undefined ? formatTime(row[colMap.startTime]) : '',
            endTime: colMap.endTime !== undefined ? formatTime(row[colMap.endTime]) : '',
            totalHours: colMap.totalHours !== undefined ? (parseFloat(row[colMap.totalHours]) || 0) : 0,
            offDutyDay: isOffDuty,
            destination: colMap.destination !== undefined ? (row[colMap.destination] || '') : ''
        });
    }

    tripData = tripData.concat(newData);
    filteredTrip = [...tripData];
    saveData();
    saveFileRecord(fileName, 'Local Trip Sheet', newData.length);
    renderAll();
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' trip records from ' + fileName, 'success');
    }
    return newData.length;
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
            'invoiceId', 'loadNum', 'broker', 'pickDate', 'pickup', 'dropDate', 'dropoff',
            'driver', 'truck', 'trailer'
        ], [
            'InvoiceID', 'Load #', 'Broker', 'Pick Date', 'Pickup', 'Drop Date', 'Dropoff',
            'Driver', 'TruckName', 'Trailer'
        ], 'loads_export.csv');
    });

    document.getElementById('exportReportCSV').addEventListener('click', function() {
        var weeks = groupByWeek(filteredReport);
        exportCSV(weeks, [
            'reportName', 'driver', 'date', 'state', 'miles'
        ], [
            'Report Name', 'Driver ID', 'Date', 'State', 'Miles'
        ], 'weekly_miles_report.csv');
    });

    // Save to Excel buttons
    document.getElementById('saveFuelExcel').addEventListener('click', function() { saveToExcel('fuel'); });
    document.getElementById('saveLoadsExcel').addEventListener('click', function() { saveToExcel('loads'); });
    document.getElementById('saveReportExcel').addEventListener('click', function() { saveToExcel('report'); });

    // Export PDF buttons
    document.getElementById('exportFuelPDF').addEventListener('click', function() { exportPDF('fuel'); });
    document.getElementById('exportLoadsPDF').addEventListener('click', function() { exportPDF('loads'); });
    document.getElementById('exportReportPDF').addEventListener('click', function() { exportPDF('report'); });

    // Trip Sheet exports
    document.getElementById('exportTripCSV').addEventListener('click', function() {
        exportCSV(filteredTrip, [
            'driverName', 'truck', 'day', 'startTime', 'endTime', 'totalHours', 'offDutyDay', 'destination'
        ], [
            'Driver Name', 'Truck', 'Day', 'Start Time', 'End Time', 'Total Hours', 'Off Duty Day', 'Destination City/State'
        ], 'tripsheet_export.csv');
    });
    document.getElementById('saveTripExcel').addEventListener('click', function() { saveToExcel('trip'); });
    document.getElementById('exportTripPDF').addEventListener('click', function() { exportPDF('trip'); });
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

// ===== Export PDF =====
var COMPANY_INFO = {
    name: 'Khalsa Logistics LLC',
    address: '9875 S 76th Street',
    cityState: 'Franklin, WI 53132',
    email: 'Khalsalogisticsllc@gmail.com',
    phone: '800-811-7308'
};

function addPDFHeader(doc, title) {
    var pageWidth = doc.internal.pageSize.getWidth();

    // Company name
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text(COMPANY_INFO.name, pageWidth / 2, 18, { align: 'center' });

    // Address
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.text(COMPANY_INFO.address, pageWidth / 2, 24, { align: 'center' });
    doc.text(COMPANY_INFO.cityState, pageWidth / 2, 29, { align: 'center' });

    // Contact
    doc.text(COMPANY_INFO.email + '  |  ' + COMPANY_INFO.phone, pageWidth / 2, 34, { align: 'center' });

    // Divider line
    doc.setDrawColor(26, 86, 219);
    doc.setLineWidth(0.5);
    doc.line(14, 38, pageWidth - 14, 38);

    // Report title
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(26, 86, 219);
    doc.text(title, pageWidth / 2, 45, { align: 'center' });
    doc.setTextColor(0, 0, 0);

    // Date
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('Generated: ' + new Date().toLocaleDateString(), pageWidth - 14, 45, { align: 'right' });

    return 50; // Y position after header
}

function exportPDF(type) {
    if (!window.jspdf || !window.jspdf.jsPDF) {
        showToast('PDF library not loaded. Please refresh the page and try again.', 'error');
        return;
    }
    var jsPDF = window.jspdf.jsPDF;

    if (type === 'fuel') {
        if (!filteredFuel.length) { showToast('No fuel data to export', 'error'); return; }
        var doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'letter' });
        var startY = addPDFHeader(doc, 'Fuel Transaction Report');

        var headers = ['Card #', 'Date', 'Time', 'Invoice', 'Unit', 'Driver', 'Odometer', 'Location', 'City', 'State', 'Fees', 'Item', 'Price', 'Qty', 'Amt', 'DB', 'Currency'];
        var rows = filteredFuel.map(function(r) {
            return [r.cardNum, formatDate(r.tranDate), r.tranTime, r.invoice, r.unit, r.driverName,
                r.odometer, r.locationName, r.city, r.state, num(r.fees), r.item,
                '$' + num(r.unitPrice), num(r.qty), '$' + num(r.amt), r.db, r.currency];
        });

        // Totals row
        var totalQty = 0, totalAmt = 0;
        filteredFuel.forEach(function(r) {
            totalQty += (parseFloat(r.qty) || 0);
            totalAmt += (parseFloat(r.amt) || 0);
        });
        rows.push(['', '', '', '', '', '', '', '', '', '', '', '', 'TOTAL:', totalQty.toFixed(2), '$' + totalAmt.toFixed(2), '', '']);

        doc.autoTable({
            head: [headers],
            body: rows,
            startY: startY,
            styles: { fontSize: 6.5, cellPadding: 1.5 },
            headStyles: { fillColor: [26, 86, 219], fontSize: 6.5, fontStyle: 'bold' },
            alternateRowStyles: { fillColor: [245, 247, 250] },
            didParseCell: function(data) {
                if (data.row.index === rows.length - 1 && data.section === 'body') {
                    data.cell.styles.fontStyle = 'bold';
                    data.cell.styles.fillColor = [219, 234, 254];
                }
            },
            margin: { left: 8, right: 8 }
        });

        doc.save('Fuel_Report.pdf');
        showToast('Fuel PDF exported', 'success');

    } else if (type === 'loads') {
        if (!filteredLoads.length) { showToast('No loads data to export', 'error'); return; }
        var doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'letter' });
        var startY = addPDFHeader(doc, 'Load Records Report');

        var headers = ['InvoiceID', 'Load #', 'Broker', 'Pick Date', 'Pickup', 'Drop Date', 'Dropoff', 'Driver', 'Truck', 'Trailer'];
        var rows = filteredLoads.map(function(r) {
            return [r.invoiceId, r.loadNum, r.broker, formatDate(r.pickDate), r.pickup,
                formatDate(r.dropDate), r.dropoff, r.driver, r.truck, r.trailer];
        });

        doc.autoTable({
            head: [headers],
            body: rows,
            startY: startY,
            styles: { fontSize: 7, cellPadding: 1.5 },
            headStyles: { fillColor: [26, 86, 219], fontSize: 7, fontStyle: 'bold' },
            alternateRowStyles: { fillColor: [245, 247, 250] },
            margin: { left: 8, right: 8 }
        });

        doc.save('Loads_Report.pdf');
        showToast('Loads PDF exported', 'success');

    } else if (type === 'report') {
        var doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });
        var pageWidth = doc.internal.pageSize.getWidth();

        // Title from table header
        var titleText = document.getElementById('rptWeeklyTitle').textContent || 'Weekly Miles Service Information';
        var startY = addPDFHeader(doc, titleText);

        var tableY = startY + 5;

        // Weekly Miles table (NO odometer columns in PDF)
        var weeks = groupByWeek(filteredReport);
        if (weeks.length) {
            var milesHeaders = ['Report Name', 'Driver ID', 'Date', 'State', 'Miles'];
            var milesRows = weeks.map(function(w) {
                return [w.reportName, w.driver, formatDate(w.date), w.state, w.miles.toFixed(2)];
            });

            var totalMilesCalc = 0;
            weeks.forEach(function(w) { totalMilesCalc += w.miles; });
            milesRows.push(['', '', '', 'Total', totalMilesCalc.toFixed(2)]);

            doc.autoTable({
                head: [milesHeaders],
                body: milesRows,
                startY: tableY,
                styles: { fontSize: 8, cellPadding: 2 },
                headStyles: { fillColor: [26, 86, 219], fontSize: 8, fontStyle: 'bold' },
                alternateRowStyles: { fillColor: [245, 247, 250] },
                columnStyles: {
                    4: { halign: 'right' }
                },
                didParseCell: function(data) {
                    if (data.row.index === milesRows.length - 1 && data.section === 'body') {
                        data.cell.styles.fontStyle = 'bold';
                        data.cell.styles.fillColor = [219, 234, 254];
                    }
                },
                margin: { left: 14, right: 14 }
            });

            tableY = doc.lastAutoTable.finalY + 10;

            // Total Earnings
            var totalEarnings = totalMilesCalc * 0.50;
            doc.setFontSize(11);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(26, 86, 219);
            doc.text('Total Earnings (@ $0.50/mile): $' + totalEarnings.toFixed(2), 14, tableY);
            doc.setTextColor(0, 0, 0);
        }

        doc.save('Driver_Earning_Report.pdf');
        showToast('Report PDF exported', 'success');

    } else if (type === 'trip') {
        if (!filteredTrip.length) { showToast('No trip sheet data to export', 'error'); return; }

        // Group by driver+truck for separate pages
        var groups = {};
        filteredTrip.forEach(function(r) {
            var key = (r.driverName || 'Unknown') + '|' + (r.truck || '');
            if (!groups[key]) groups[key] = [];
            groups[key].push(r);
        });

        var doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });
        var pageWidth = doc.internal.pageSize.getWidth();
        var pageHeight = doc.internal.pageSize.getHeight();
        var isFirst = true;

        Object.keys(groups).forEach(function(key) {
            if (!isFirst) doc.addPage();
            isFirst = false;

            var parts = key.split('|');
            var driverName = parts[0];
            var truckName = parts[1];
            var records = groups[key];

            var y = 14;

            // Company header
            doc.setFontSize(14);
            doc.setFont('helvetica', 'bold');
            doc.text(COMPANY_INFO.name, pageWidth / 2, y, { align: 'center' });
            y += 5;
            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');
            doc.text(COMPANY_INFO.address + ', ' + COMPANY_INFO.cityState, pageWidth / 2, y, { align: 'center' });
            y += 4;
            doc.text(COMPANY_INFO.email + '  |  ' + COMPANY_INFO.phone, pageWidth / 2, y, { align: 'center' });
            y += 8;

            // Driver Name / Truck #
            doc.setFontSize(10);
            doc.setFont('helvetica', 'bold');
            doc.text("Driver's Name: ", 14, y);
            doc.setFont('helvetica', 'normal');
            doc.text(driverName, 50, y);
            doc.line(48, y + 1, 110, y + 1);

            doc.setFont('helvetica', 'bold');
            doc.text('Truck #: ', 120, y);
            doc.setFont('helvetica', 'normal');
            doc.text(truckName, 142, y);
            doc.line(140, y + 1, pageWidth - 14, y + 1);
            y += 8;

            // DOT disclaimer block
            doc.setFontSize(7.5);
            doc.setFont('helvetica', 'bold');
            doc.text('DRIVERS MAY PREPARE THIS REPORT INSTEAD OF A "DRIVER\'S DAILY LOG" WHEN OPERATING WITHIN 150 AIR-', pageWidth / 2, y, { align: 'center' });
            y += 3.5;
            doc.text('MILES OF THE DRIVER\'S WORK REPORTING LOCATION IF THE FOLLOWING APPLIES:', pageWidth / 2, y, { align: 'center' });
            y += 5;

            // Two-column rules table
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(6.5);
            var rulesLeft = [
                'Drivers of CDL Vehicles:',
                'Returns to work reporting location within 14 hours',
                'Has 10 consecutive hours off between shifts',
                'Does not drive after 60 or 70 hours in 7 or 8 consecutive days',
                'Passenger Carriers',
                'Must have 8 consecutive hrs. off duty between 14 hr. duty shift No 34 hour',
                'restart, No 16 hour provision'
            ];
            var rulesRight = [
                'WI Drivers of Non CDL Vehicles When Crossing State Lines:',
                'Returns to the work reporting location at end of each shift',
                'Has 10 consecutive hours off between shifts',
                'Maximum 11 hours driving time',
                'Does not drive after 60 or 70 hours in 7 or 8 consecutive days',
                'Does not drive after 14th hour 5 days of 7 consecutive days',
                'Does not drive after the 16th hour 2 days in 7 consecutive days'
            ];

            var ruleStartY = y;
            var midX = pageWidth / 2;
            doc.setDrawColor(0);
            doc.setLineWidth(0.2);

            // Draw rules box
            for (var ri = 0; ri < rulesLeft.length; ri++) {
                var ry = ruleStartY + (ri * 4);
                if (ri === 0) {
                    doc.setFont('helvetica', 'bold');
                } else {
                    doc.setFont('helvetica', 'normal');
                }
                doc.text(rulesLeft[ri], 15, ry + 3);
                doc.text(rulesRight[ri], midX + 2, ry + 3);
                doc.line(14, ry, pageWidth - 14, ry);
                if (ri === 0) doc.line(midX, ruleStartY, midX, ruleStartY + rulesLeft.length * 4);
            }
            doc.line(14, ruleStartY + rulesLeft.length * 4, pageWidth - 14, ruleStartY + rulesLeft.length * 4);
            doc.rect(14, ruleStartY, pageWidth - 28, rulesLeft.length * 4);

            y = ruleStartY + rulesLeft.length * 4 + 5;

            // Notices
            doc.setFontSize(7);
            doc.setFont('helvetica', 'bold');
            doc.text('All hours of service records MUST Be kept for 6 months by the motor carrier', pageWidth / 2, y, { align: 'center' });
            y += 5;
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(6.5);
            doc.text('INTERMITTENT DRIVERS: Shall complete this form for 7 days preceding any day regulated driving is performed. This includes the preceding month.', 14, y);
            y += 5;
            var longText = 'To be prepared monthly by each DOT-certified driver unless time record is exclusively kept on a Driver\'s Daily Log. Indicate "days off" Check box if NO driving is performed during this month and the first 7 days of the following month.';
            var splitText = doc.splitTextToSize(longText, pageWidth - 28);
            doc.text(splitText, 14, y);
            y += splitText.length * 3.5 + 4;

            // Trip data table
            var tripHeaders = ['Day', 'Start Time\n"All Duty"', 'End Time\n"All Duty"', 'Total\nHours', 'Off Duty\nDay', 'Trip Destination City/State'];
            var tripRows = records.map(function(r) {
                return [
                    formatDate(r.day),
                    r.startTime || '',
                    r.endTime || '',
                    r.totalHours ? r.totalHours.toFixed(2) : '',
                    r.offDutyDay ? 'X' : '',
                    r.destination || ''
                ];
            });

            // Pad to at least 31 rows for full month
            while (tripRows.length < 31) {
                tripRows.push(['', '', '', '', '', '']);
            }

            // Add total row
            var totalHrs = 0;
            records.forEach(function(r) { totalHrs += (parseFloat(r.totalHours) || 0); });
            tripRows.push(['', '', 'Total:', totalHrs.toFixed(2), '', '']);

            doc.autoTable({
                head: [tripHeaders],
                body: tripRows,
                startY: y,
                styles: { fontSize: 7, cellPadding: 1.2, lineColor: [0, 0, 0], lineWidth: 0.2 },
                headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontSize: 6.5, fontStyle: 'bold', lineColor: [0, 0, 0], lineWidth: 0.2 },
                columnStyles: {
                    0: { cellWidth: 22 },
                    1: { cellWidth: 22 },
                    2: { cellWidth: 22 },
                    3: { cellWidth: 16 },
                    4: { cellWidth: 18 },
                    5: { cellWidth: pageWidth - 28 - 100 }
                },
                theme: 'grid',
                didParseCell: function(data) {
                    if (data.row.index === tripRows.length - 1 && data.section === 'body') {
                        data.cell.styles.fontStyle = 'bold';
                    }
                },
                margin: { left: 14, right: 14 }
            });

            var finalY = doc.lastAutoTable.finalY + 10;

            // Driver signature line
            if (finalY < pageHeight - 20) {
                doc.setFontSize(9);
                doc.setFont('helvetica', 'bold');
                doc.text('Driver Signature:', 14, finalY);
                doc.line(50, finalY + 1, 130, finalY + 1);
            }
        });

        doc.save('Local_TripSheet.pdf');
        showToast('Trip Sheet PDF exported', 'success');
    }
}

// ===== Actions =====
function setupActions() {
    document.getElementById('clearAllData').addEventListener('click', function() {
        if (!confirm('Are you sure you want to clear ALL stored data? This cannot be undone.')) return;
        fuelData = [];
        loadsData = [];
        reportData = [];
        tripData = [];
        filteredFuel = [];
        filteredLoads = [];
        filteredReport = [];
        filteredTrip = [];
        localStorage.removeItem(STORAGE_KEYS.fuel);
        localStorage.removeItem(STORAGE_KEYS.loads);
        localStorage.removeItem(STORAGE_KEYS.report);
        localStorage.removeItem(STORAGE_KEYS.trip);
        localStorage.removeItem(STORAGE_KEYS.files);
        renderAll();
        showToast('All data cleared', 'success');
    });

    document.getElementById('loadSampleData').addEventListener('click', loadSampleData);
}

function loadSampleData() {
    var sampleFuel = [
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.812,qty:161.23,amt:614.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"DEFD",unitPrice:4.399,qty:15.14,amt:66.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"ULSD",unitPrice:3.685,qty:161.65,amt:595.68,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:11.29,amt:49.66,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-07",tranTime:"07:00",invoice:57250,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO RACINE",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:3.277,qty:164.76,amt:539.92,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-08",tranTime:"19:17",invoice:47082,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.576,qty:174.86,amt:625.30,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-12",tranTime:"05:02",invoice:6237,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.925,qty:170.72,amt:670.07,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-13",tranTime:"02:22",invoice:52717,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.573,qty:167.15,amt:597.22,db:"N",currency:"USD/Gallons"}
    ];

    var sampleLoads = [
        {invoiceId:"INV001",loadNum:15368128,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City A, OH 45241",dropDate:"2025-11-03",dropoff:"City B, WI 53081",driver:"Driver A",truck:"MPL5046",trailer:"MPL1016"},
        {invoiceId:"INV002",loadNum:15393089,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City B, WI 53081",dropDate:"2025-11-03",dropoff:"City A, OH 45241",driver:"Driver A",truck:"MPL5046",trailer:"MPL1016"},
        {invoiceId:"INV003",loadNum:15392627,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City A, OH 45241",dropDate:"2025-11-04",dropoff:"City B, WI 53081",driver:"Driver A",truck:"MPL5046",trailer:"MPL1008"},
        {invoiceId:"INV004",loadNum:15392834,broker:"Sample Logistics LLC",pickDate:"2025-11-04",pickup:"City B, WI 53081",dropDate:"2025-11-04",dropoff:"City A, OH 45241",driver:"Driver A",truck:"MPL5046",trailer:"MPL1008"},
        {invoiceId:"INV005",loadNum:15392694,broker:"Sample Logistics LLC",pickDate:"2025-11-04",pickup:"City A, OH 45241",dropDate:"2025-11-06",dropoff:"City B, WI 53081",driver:"Driver A / Driver B",truck:"MPL5046",trailer:"MPL1015"}
    ];

    var sampleReport = [
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-01",startOdo:57685,endOdo:57858,totalMiles:173},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-02",startOdo:57858,endOdo:57858,totalMiles:0},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-03",startOdo:57858,endOdo:57858,totalMiles:0},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-03",startOdo:58286,endOdo:58924,totalMiles:638},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-04",startOdo:58924,endOdo:59067,totalMiles:143},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-04",startOdo:59203,endOdo:59587,totalMiles:384},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-05",startOdo:59587,endOdo:60274,totalMiles:687},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-06",startOdo:60274,endOdo:60274,totalMiles:0},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-06",startOdo:60289,endOdo:60981,totalMiles:692},
        {driverName:"Driver A",truck:"MPL5046",date:"2025-11-07",startOdo:60981,endOdo:61048,totalMiles:67}
    ];

    fuelData = sampleFuel;
    loadsData = sampleLoads;
    reportData = sampleReport;
    filteredFuel = [...fuelData];
    filteredLoads = [...loadsData];
    filteredReport = [...reportData];
    saveData();
    saveFileRecord('Sample - Fuel', 'Fuel', sampleFuel.length);
    saveFileRecord('Sample - Loads', 'Loads', sampleLoads.length);
    saveFileRecord('Sample - Report', 'Report (Miles/Odo)', sampleReport.length);
    renderAll();
    showToast('Sample data loaded successfully!', 'success');
}

// ===== Utility Functions =====

// Safe value getter - returns empty string if column index is undefined
function safeVal(row, colIdx) {
    if (colIdx === undefined || colIdx === null) return '';
    if (colIdx >= row.length) return '';
    return row[colIdx];
}

function excelDateToStr(val) {
    if (!val && val !== 0) return '';
    if (val instanceof Date) {
        // Use UTC methods to avoid timezone shifting
        var y = val.getUTCFullYear();
        var m = String(val.getUTCMonth() + 1).padStart(2, '0');
        var d = String(val.getUTCDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    if (typeof val === 'string') {
        // Already YYYY-MM-DD
        if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.substring(0, 10);
        // Try to parse other date string formats
        var parsed = new Date(val);
        if (!isNaN(parsed.getTime())) {
            return parsed.getUTCFullYear() + '-' + String(parsed.getUTCMonth()+1).padStart(2,'0') + '-' + String(parsed.getUTCDate()).padStart(2,'0');
        }
    }
    if (typeof val === 'number' && val > 1000) {
        // Excel serial date (only if it looks like a date serial, not a regular number)
        var date = new Date((val - 25569) * 86400000);
        return date.getUTCFullYear() + '-' + String(date.getUTCMonth()+1).padStart(2,'0') + '-' + String(date.getUTCDate()).padStart(2,'0');
    }
    return '';
}

function formatExcelTime(val) {
    if (!val && val !== 0) return '';
    if (val instanceof Date) {
        return String(val.getUTCHours()).padStart(2, '0') + ':' + String(val.getUTCMinutes()).padStart(2, '0');
    }
    if (typeof val === 'number' && val < 1) {
        // Excel decimal time (0.0 to 0.999)
        var totalMinutes = Math.round(val * 24 * 60);
        var hours = Math.floor(totalMinutes / 60);
        var minutes = totalMinutes % 60;
        return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
    }
    if (typeof val === 'string') {
        // Already a time string like "08:08" or "8:08 AM"
        return val;
    }
    return '';
}

function formatTime(val) {
    if (!val) return '';
    // If it's a Date object (from Excel time-only cells), extract just the time
    if (val instanceof Date) {
        var h = val.getHours();
        var m = val.getMinutes();
        var ampm = h >= 12 ? 'PM' : 'AM';
        h = h % 12;
        if (h === 0) h = 12;
        return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ampm;
    }
    var s = String(val);
    // If it looks like a full Date string (e.g. "Sun Dec 31 1899 11:26:35 ..."), extract the time
    var match = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
    if (match) {
        var hr = parseInt(match[1]);
        var min = match[2];
        var ap = match[4];
        if (!ap) {
            ap = hr >= 12 ? 'PM' : 'AM';
            hr = hr % 12;
            if (hr === 0) hr = 12;
        }
        return hr + ':' + min + ' ' + ap.toUpperCase();
    }
    return s;
}

function formatDate(dateStr) {
    if (!dateStr) return '';
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

function escAttr(val) {
    if (val == null || val === '') return '';
    return String(val).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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

// ===== Settings & Audit Mode =====
var settingsData = { users: [], auditMode: false };

function setupSettings() {
    // Load settings from JSON
    fetch('data/settings.json')
        .then(function(res) { return res.json(); })
        .then(function(data) {
            settingsData = data;
            renderUsersTable();
            // Sync audit toggle with loaded settings
            var toggle = document.getElementById('auditModeToggle');
            // Check sessionStorage first (set at login), fallback to file
            var sessionAudit = sessionStorage.getItem('tms_audit_mode');
            if (sessionAudit !== null) {
                toggle.checked = sessionAudit === 'true';
            } else {
                toggle.checked = settingsData.auditMode || false;
            }
            applyAuditMode();
        })
        .catch(function() {
            settingsData = { users: [], auditMode: false };
        });

    // Audit mode toggle
    document.getElementById('auditModeToggle').addEventListener('change', function() {
        settingsData.auditMode = this.checked;
        sessionStorage.setItem('tms_audit_mode', this.checked ? 'true' : 'false');
        applyAuditMode();
    });

    // Reload from Excel files button
    document.getElementById('reloadFromExcelBtn').addEventListener('click', function() {
        if (!confirm('This will replace ALL current data (Fuel, Loads, Report, Trip Sheet) with data from the Excel files in the data/ folder. Continue?')) return;
        var btn = this;
        var statusEl = document.getElementById('reloadStatus');
        btn.disabled = true;
        statusEl.textContent = 'Loading...';

        // Clear existing data
        fuelData = []; loadsData = []; reportData = []; tripData = [];
        filteredFuel = []; filteredLoads = []; filteredReport = []; filteredTrip = [];
        localStorage.removeItem(STORAGE_KEYS.fuel);
        localStorage.removeItem(STORAGE_KEYS.loads);
        localStorage.removeItem(STORAGE_KEYS.report);
        localStorage.removeItem(STORAGE_KEYS.trip);
        localStorage.removeItem(STORAGE_KEYS.files);

        // Re-load from Excel files
        var dataFiles = [
            { url: 'data/Fuel.xlsx', type: 'fuel' },
            { url: 'data/Loads.xlsx', type: 'loads' },
            { url: 'data/Report.xlsx', type: 'report' },
            { url: 'data/LocalTripsheet.xlsx', type: 'trip' }
        ];
        var loaded = 0;
        var successCount = 0;

        dataFiles.forEach(function(fileInfo) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', fileInfo.url + '?t=' + Date.now(), true);
            xhr.responseType = 'arraybuffer';
            xhr.onload = function() {
                if (xhr.status === 200) {
                    try {
                        var data = new Uint8Array(xhr.response);
                        var workbook = XLSX.read(data, { type: 'array', cellDates: true });
                        if (fileInfo.type === 'fuel') parseFuelWorkbook(workbook, fileInfo.url);
                        else if (fileInfo.type === 'loads') parseLoadsWorkbook(workbook, fileInfo.url);
                        else if (fileInfo.type === 'report') parseReportWorkbook(workbook, fileInfo.url);
                        else if (fileInfo.type === 'trip') parseTripWorkbook(workbook, fileInfo.url);
                        successCount++;
                    } catch (err) {
                        console.warn('Could not parse ' + fileInfo.url + ':', err.message);
                    }
                }
                loaded++;
                if (loaded === dataFiles.length) {
                    renderAll();
                    populateFilterDropdowns();
                    btn.disabled = false;
                    statusEl.textContent = successCount + ' file(s) loaded successfully.';
                    showToast('Data reloaded from Excel files (' + successCount + ' files)', 'success');
                    setTimeout(function() { statusEl.textContent = ''; }, 5000);
                }
            };
            xhr.onerror = function() {
                loaded++;
                if (loaded === dataFiles.length) {
                    renderAll();
                    btn.disabled = false;
                    statusEl.textContent = successCount + ' file(s) loaded.';
                }
            };
            xhr.send();
        });
    });

    // Add user button
    document.getElementById('addUserBtn').addEventListener('click', function() {
        settingsData.users.push({ username: '', password: '', name: '' });
        renderUsersTable();
    });

    // Save settings to file (downloads updated settings.json)
    document.getElementById('saveSettingsBtn').addEventListener('click', function() {
        // Read current values from table inputs
        syncUsersFromTable();
        settingsData.auditMode = document.getElementById('auditModeToggle').checked;

        var blob = new Blob([JSON.stringify(settingsData, null, 2)], { type: 'application/json' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'settings.json';
        a.click();
        URL.revokeObjectURL(url);
        showToast('Settings file downloaded. Replace data/settings.json and push to update.', 'success');
    });
}

function renderUsersTable() {
    var tbody = document.querySelector('#usersTable tbody');
    tbody.innerHTML = '';
    settingsData.users.forEach(function(u, i) {
        var tr = document.createElement('tr');
        tr.innerHTML =
            '<td><input type="text" class="settings-input" data-field="username" data-idx="' + i + '" value="' + escAttr(u.username) + '"></td>' +
            '<td><input type="password" class="settings-input" data-field="password" data-idx="' + i + '" value="' + escAttr(u.password) + '"></td>' +
            '<td><input type="text" class="settings-input" data-field="name" data-idx="' + i + '" value="' + escAttr(u.name) + '"></td>' +
            '<td><button class="btn-delete" data-idx="' + i + '">Remove</button></td>';
        tbody.appendChild(tr);
    });

    // Wire up remove buttons
    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            var idx = parseInt(this.dataset.idx);
            settingsData.users.splice(idx, 1);
            renderUsersTable();
        });
    });

    // Wire up input changes
    tbody.querySelectorAll('.settings-input').forEach(function(input) {
        input.addEventListener('change', function() {
            var idx = parseInt(this.dataset.idx);
            var field = this.dataset.field;
            settingsData.users[idx][field] = this.value;
        });
    });
}

function syncUsersFromTable() {
    document.querySelectorAll('#usersTable .settings-input').forEach(function(input) {
        var idx = parseInt(input.dataset.idx);
        var field = input.dataset.field;
        if (settingsData.users[idx]) {
            settingsData.users[idx][field] = input.value;
        }
    });
}

function applyAuditMode() {
    var isAudit = sessionStorage.getItem('tms_audit_mode') === 'true';
    var toggle = document.getElementById('auditModeToggle');
    if (toggle) toggle.checked = isAudit;

    var statusEl = document.getElementById('auditModeStatus');
    if (isAudit) {
        document.body.classList.add('audit-mode');
        if (statusEl) {
            statusEl.textContent = 'AUDIT MODE IS ON — All data modification features are hidden.';
            statusEl.className = 'audit-status active';
        }
    } else {
        document.body.classList.remove('audit-mode');
        if (statusEl) {
            statusEl.textContent = 'Audit mode is off — Normal operation.';
            statusEl.className = 'audit-status inactive';
        }
    }
}
