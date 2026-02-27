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
    setupModal();
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
        tbody.innerHTML = '<tr><td colspan="14" class="empty-state">No fuel records found. Upload a fuel file to get started.</td></tr>';
    } else {
        filteredFuel.forEach(function(r, idx) {
            // Find the real index in fuelData
            var realIdx = fuelData.indexOf(r);
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
                '<td class="amount-cell">$' + num(r.amt) + '</td>' +
                '<td><button class="btn-delete" data-type="fuel" data-idx="' + realIdx + '">Delete</button></td>';
            tbody.appendChild(tr);
            totalQty += (parseFloat(r.qty) || 0);
            totalAmt += (parseFloat(r.amt) || 0);
        });
    }

    document.getElementById('fuelTotalQty').innerHTML = '<strong>' + totalQty.toFixed(2) + '</strong>';
    document.getElementById('fuelTotalAmt').innerHTML = '<strong>$' + totalAmt.toFixed(2) + '</strong>';

    // Attach delete handlers
    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            deleteRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
}

function renderLoadsTable() {
    var tbody = document.querySelector('#loadsTable tbody');
    tbody.innerHTML = '';

    if (filteredLoads.length === 0) {
        tbody.innerHTML = '<tr><td colspan="15" class="empty-state">No load records found. Upload a loads file to get started.</td></tr>';
    } else {
        filteredLoads.forEach(function(r, idx) {
            var realIdx = loadsData.indexOf(r);
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
                '<td>' + esc(r.notes) + '</td>' +
                '<td><button class="btn-delete" data-type="loads" data-idx="' + realIdx + '">Delete</button></td>';
            tbody.appendChild(tr);
        });
    }

    // Attach delete handlers
    tbody.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function() {
            deleteRecord(this.dataset.type, parseInt(this.dataset.idx));
        });
    });
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
    if (filteredFuel.length > 0 && (!driver || driver === '--')) {
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
    }

    saveData();
    applyFilters();
    showToast('Record deleted', 'success');
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
            '<div class="form-group"><label>Invoice ID</label><input type="text" id="mLoadInvoice"></div>' +
            '<div class="form-group"><label>Load #</label><input type="text" id="mLoadNum"></div>' +
        '</div>' +
        '<div class="form-group"><label>Broker</label><input type="text" id="mLoadBroker"></div>' +
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
            '<div class="form-group"><label>Shipping ID</label><input type="text" id="mLoadShipping"></div>' +
        '</div>' +
        '<div class="form-row">' +
            '<div class="form-group"><label>PU Datetime</label><input type="text" id="mLoadPU" placeholder="PU 11/03 | Time 02PM"></div>' +
            '<div class="form-group"><label>DO Datetime</label><input type="text" id="mLoadDO" placeholder="DO 11/04 | Local Miles"></div>' +
        '</div>' +
        '<div class="form-group"><label>Notes</label><input type="text" id="mLoadNotes"></div>';

    document.getElementById('modalSave').onclick = function() {
        var record = {
            invoiceId: document.getElementById('mLoadInvoice').value,
            loadNum: document.getElementById('mLoadNum').value,
            broker: document.getElementById('mLoadBroker').value,
            pickDate: document.getElementById('mLoadPickDate').value,
            pickup: document.getElementById('mLoadPickup').value,
            dropDate: document.getElementById('mLoadDropDate').value,
            dropoff: document.getElementById('mLoadDropoff').value,
            driver: document.getElementById('mLoadDriver').value,
            truck: document.getElementById('mLoadTruck').value,
            trailer: document.getElementById('mLoadTrailer').value,
            shippingId: document.getElementById('mLoadShipping').value,
            puDatetime: document.getElementById('mLoadPU').value,
            doDatetime: document.getElementById('mLoadDO').value,
            notes: document.getElementById('mLoadNotes').value
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
}

// ===== Combined File Parser =====
function parseCombinedFile(file) {
    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array', cellDates: true });
            var totalRecords = { fuel: 0, loads: 0, report: 0 };

            // Find and parse Fuel sheet
            var fuelSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('fuel') !== -1;
            });
            if (fuelSheet) {
                var count = parseFuelWorkbook(workbook, file.name, fuelSheet);
                totalRecords.fuel = count || 0;
            }

            // Find and parse Loads sheet
            var loadsSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('load') !== -1;
            });
            if (loadsSheet) {
                var count = parseLoadsWorkbook(workbook, file.name, loadsSheet);
                totalRecords.loads = count || 0;
            }

            // Find and parse Report sheet
            var reportSheet = workbook.SheetNames.find(function(n) {
                return n.toLowerCase().indexOf('report') !== -1;
            });
            if (reportSheet) {
                var count = parseReportWorkbook(workbook, file.name, reportSheet);
                totalRecords.report = count || 0;
            }

            var msg = 'Loaded from ' + file.name + ': ';
            var parts = [];
            if (totalRecords.fuel > 0) parts.push(totalRecords.fuel + ' fuel');
            if (totalRecords.loads > 0) parts.push(totalRecords.loads + ' loads');
            if (totalRecords.report > 0) parts.push(totalRecords.report + ' report');
            if (parts.length === 0) {
                showToast('No data sheets found in file. Expected sheets named Fuel, Loads, or Report.', 'error');
            } else {
                showToast(msg + parts.join(', ') + ' records', 'success');
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
            }
        } catch (err) {
            console.error('Parse error:', err);
            showToast('Error parsing file: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseFuelWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('fuel') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row - Fuel data is in A1:Q50
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
        if (!sheetNameOverride) showToast('Could not find fuel data headers in file', 'error');
        return 0;
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
    // Parse up to row 50 (A1:Q50 range) but allow more if data exists
    var maxRow = Math.min(rows.length, 50);
    for (var r = headerIdx + 1; r < maxRow; r++) {
        var row = rows[r];
        if (!row || (!row[colMap.cardNum] && !row[colMap.tranDate])) continue;
        // Skip if the row looks empty
        var hasData = false;
        for (var c = 0; c < Math.min(row.length, 17); c++) {
            if (row[c] !== '' && row[c] !== null && row[c] !== undefined) { hasData = true; break; }
        }
        if (!hasData) continue;
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
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' fuel records from ' + fileName, 'success');
    }
    return newData.length;
}

function parseLoadsWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('load') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Find header row - Loads data useful range is D1:J49
    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (row && row.some(function(c) {
            var s = String(c).toLowerCase();
            return s.indexOf('invoiceid') !== -1 ||
                   s.indexOf('load #') !== -1 ||
                   s.indexOf('load') !== -1 ||
                   s.indexOf('pick date') !== -1 ||
                   s.indexOf('broker') !== -1;
        })) {
            headerIdx = i;
            break;
        }
    }

    if (headerIdx === -1) {
        if (!sheetNameOverride) showToast('Could not find load data headers in file', 'error');
        return 0;
    }

    var headers = rows[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
    var colMap = {
        invoiceId: findCol(headers, ['invoiceid', 'invoice id', 'invoice']),
        loadNum: findCol(headers, ['load #', 'load']),
        broker: findCol(headers, ['broker']),
        pickDate: findCol(headers, ['pick date', 'pickup date']),
        pickup: findCol(headers, ['pickup']),
        dropDate: findCol(headers, ['drop date', 'dropoff date']),
        dropoff: findCol(headers, ['dropoff', 'drop off']),
        driver: findCol(headers, ['driver']),
        truck: findCol(headers, ['truck', 'truckname']),
        trailer: findCol(headers, ['trailer']),
        shippingId: findCol(headers, ['shipping', 'shipping id']),
        puDatetime: findCol(headers, ['pickup datetime', 'pu datetime', 'pu date']),
        doDatetime: findCol(headers, ['delivery datetime', 'do datetime', 'do date']),
        notes: findCol(headers, ['notes', 'note'])
    };

    var newData = [];
    var maxRow = Math.min(rows.length, 49);
    for (var r = headerIdx + 1; r < maxRow; r++) {
        var row = rows[r];
        if (!row || (!row[colMap.invoiceId] && !row[colMap.loadNum] && !row[colMap.pickDate])) continue;
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
    if (!sheetNameOverride) {
        showToast('Loaded ' + newData.length + ' load records from ' + fileName, 'success');
    }
    return newData.length;
}

function parseReportWorkbook(workbook, fileName, sheetNameOverride) {
    var sheetName = sheetNameOverride || workbook.SheetNames.find(function(n) {
        return n.toLowerCase().indexOf('report') !== -1;
    }) || workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    var newData = [];
    var headerIdx = -1;
    var dateCol = -1;
    var startOdoCol = -1;
    var endOdoCol = -1;
    var totalMileCol = -1;
    var missingCol = -1;
    var notesCol = -1;

    // Report miles data is in R9:U50 (column R = index 17, row 9 = index 8)
    // First try to find headers in column R area (index 17+) around row 8-9
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
        var row = rows[i];
        if (!row) continue;
        for (var j = 0; j < row.length; j++) {
            var cellStr = String(row[j]).toLowerCase().trim();
            if (cellStr === 'date' || cellStr === 'dates') {
                // Prefer date columns in the R+ range (index 17+) for report
                if (dateCol === -1 || j >= 17) {
                    dateCol = j;
                    headerIdx = i;
                }
            }
            if (cellStr.indexOf('start odo') !== -1 || cellStr.indexOf('start_odo') !== -1 || cellStr === 'start odo') startOdoCol = j;
            if (cellStr.indexOf('end odo') !== -1 || cellStr.indexOf('end_odo') !== -1 || cellStr === 'end odo') endOdoCol = j;
            if (cellStr.indexOf('total mile') !== -1 || cellStr === 'total miles') totalMileCol = j;
            if (cellStr.indexOf('missing') !== -1) missingCol = j;
            if (cellStr === 'details' || cellStr === 'notes') notesCol = j;
        }
        if (dateCol !== -1 && startOdoCol !== -1) break;
    }

    // If no explicit headers found, try R9:U50 layout (col R=17, starting row 8)
    if (dateCol === -1) {
        // Check if column R (index 17) has date-like data starting around row 8
        for (var i = 7; i < Math.min(rows.length, 12); i++) {
            var row = rows[i];
            if (!row || !row[17]) continue;
            var cellStr = String(row[17]).toLowerCase().trim();
            if (cellStr === 'date' || cellStr === 'dates') {
                dateCol = 17;
                headerIdx = i;
                startOdoCol = 18; // S
                endOdoCol = 19;   // T
                totalMileCol = 20; // U
                break;
            }
            // Check if it's actual date data (no header row)
            var testDate = excelDateToStr(row[17]);
            if (testDate && /^\d{4}-\d{2}-\d{2}$/.test(testDate)) {
                dateCol = 17;
                headerIdx = i - 1; // data starts here
                startOdoCol = 18;
                endOdoCol = 19;
                totalMileCol = 20;
                break;
            }
        }
    }

    if (dateCol === -1) {
        if (!sheetNameOverride) showToast('Could not find date/odometer columns in report file. Looking for: Date, Start Odo, End Odo, Total Miles', 'error');
        return 0;
    }

    // Defaults for columns not found
    if (startOdoCol === -1) startOdoCol = dateCol + 1;
    if (endOdoCol === -1) endOdoCol = dateCol + 2;
    if (totalMileCol === -1) totalMileCol = dateCol + 3;
    if (missingCol === -1) missingCol = dateCol + 4;

    var maxRow = Math.min(rows.length, 50);
    for (var r = headerIdx + 1; r < maxRow; r++) {
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
            missingMiles: missingCol < row.length ? (row[missingCol] || 0) : 0,
            notes: notesCol !== -1 && notesCol < row.length ? (row[notesCol] || '') : ''
        };
        newData.push(record);
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
    var sampleFuel = [
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.812,qty:161.23,amt:614.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-03",tranTime:"08:08",invoice:1570,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"DEFD",unitPrice:4.399,qty:15.14,amt:66.61,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"ULSD",unitPrice:3.685,qty:161.65,amt:595.68,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-05",tranTime:"07:42",invoice:38944,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO GREENSBURG",city:"GREENSBURG",state:"IN",fees:0,item:"DEFD",unitPrice:4.399,qty:11.29,amt:49.66,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-07",tranTime:"07:00",invoice:57250,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO RACINE",city:"STURTEVANT",state:"WI",fees:0,item:"ULSD",unitPrice:3.277,qty:164.76,amt:539.92,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-07",tranTime:"07:00",invoice:57250,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO RACINE",city:"STURTEVANT",state:"WI",fees:0,item:"DEFD",unitPrice:4.399,qty:10.75,amt:47.30,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-08",tranTime:"19:17",invoice:47082,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"PETRO REMINGTON",city:"REMINGTON",state:"IN",fees:0,item:"ULSD",unitPrice:3.576,qty:174.86,amt:625.30,db:"N",currency:"USD/Gallons"},
        {cardNum:959,tranDate:"2025-11-12",tranTime:"05:02",invoice:6237,unit:"MPL5046",driverName:"Driver A",odometer:0,locationName:"TA CHICAGO NORTH",city:"RUSSELL",state:"IL",fees:0,item:"ULSD",unitPrice:3.925,qty:170.72,amt:670.07,db:"N",currency:"USD/Gallons"}
    ];

    var sampleLoads = [
        {invoiceId:"INV001",loadNum:15368128,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City A, OH 45241",dropDate:"2025-11-03",dropoff:"City B, WI 53081",driver:"Driver A",truck:"MPL5046",trailer:"MPL1016",shippingId:580988,puDatetime:"PU 11/03 | Time 03PM",doDatetime:"DO 11/03 | Local Miles",notes:""},
        {invoiceId:"INV002",loadNum:15393089,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City B, WI 53081",dropDate:"2025-11-03",dropoff:"City A, OH 45241",driver:"Driver A",truck:"MPL5046",trailer:"MPL1016",shippingId:6851123,puDatetime:"PU 11/03 | Local Miles",doDatetime:"DO 11/03 | Time 02PM",notes:""},
        {invoiceId:"INV003",loadNum:15392627,broker:"Sample Logistics LLC",pickDate:"2025-11-03",pickup:"City A, OH 45241",dropDate:"2025-11-04",dropoff:"City B, WI 53081",driver:"Driver A",truck:"MPL5046",trailer:"MPL1008",shippingId:581331,puDatetime:"PU 11/03 | Time 02PM",doDatetime:"DO 11/04 | Local Miles",notes:""},
        {invoiceId:"INV004",loadNum:15392834,broker:"Sample Logistics LLC",pickDate:"2025-11-04",pickup:"City B, WI 53081",dropDate:"2025-11-04",dropoff:"City A, OH 45241",driver:"Driver A",truck:"MPL5046",trailer:"MPL1008",shippingId:6851159,puDatetime:"PU 11/04 | Local Miles",doDatetime:"DO 11/04 | Time 07PM",notes:"Sample note"},
        {invoiceId:"INV005",loadNum:15392694,broker:"Sample Logistics LLC",pickDate:"2025-11-04",pickup:"City A, OH 45241",dropDate:"2025-11-06",dropoff:"City B, WI 53081",driver:"Driver A / Driver B",truck:"MPL5046",trailer:"MPL1015",shippingId:581572,puDatetime:"PU 11/04 | Time 07PM",doDatetime:"DO 11/06 | Local Driver",notes:""}
    ];

    var sampleReport = [
        {date:"2025-11-01",startOdo:57685,endOdo:57858,totalMiles:173,missingMiles:0,notes:""},
        {date:"2025-11-02",startOdo:57858,endOdo:57858,totalMiles:0,missingMiles:0,notes:""},
        {date:"2025-11-03",startOdo:57858,endOdo:57858,totalMiles:0,missingMiles:428,notes:""},
        {date:"2025-11-03",startOdo:58286,endOdo:58924,totalMiles:638,missingMiles:0,notes:""},
        {date:"2025-11-04",startOdo:58924,endOdo:59067,totalMiles:143,missingMiles:136,notes:""},
        {date:"2025-11-04",startOdo:59203,endOdo:59587,totalMiles:384,missingMiles:0,notes:""},
        {date:"2025-11-05",startOdo:59587,endOdo:60274,totalMiles:687,missingMiles:0,notes:""},
        {date:"2025-11-06",startOdo:60274,endOdo:60274,totalMiles:0,missingMiles:15,notes:""},
        {date:"2025-11-06",startOdo:60289,endOdo:60981,totalMiles:692,missingMiles:0,notes:""},
        {date:"2025-11-07",startOdo:60981,endOdo:61048,totalMiles:67,missingMiles:134,notes:""}
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
        if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.substring(0, 10);
        var parsed = new Date(val);
        if (!isNaN(parsed.getTime())) {
            return parsed.getFullYear() + '-' + String(parsed.getMonth()+1).padStart(2,'0') + '-' + String(parsed.getDate()).padStart(2,'0');
        }
    }
    if (typeof val === 'number') {
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
