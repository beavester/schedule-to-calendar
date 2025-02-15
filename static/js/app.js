let currentFilePath = null;
let shiftMappings = {};

// Setup drag and drop handlers
const uploadArea = document.querySelector('.upload-area');
const fileInput = document.getElementById('file-input');
const employeeSection = document.getElementById('employee-section');
const employeeSelect = document.getElementById('employee-select');
const generateButton = document.getElementById('generate-calendar');
const loadingIndicator = document.getElementById('loading');
const editShiftsButton = document.getElementById('edit-shifts-button');
const shiftsModal = new bootstrap.Modal(document.getElementById('shifts-modal'));
const shiftMappingsTable = document.getElementById('shift-mappings-table');
const saveShiftsButton = document.getElementById('save-shifts');

// Load shift mappings on page load
fetch('/shifts')
    .then(response => response.json())
    .then(data => {
        shiftMappings = data;
        updateShiftMappingsTable();
    })
    .catch(error => showAlert('Error loading shift mappings'));

function updateShiftMappingsTable() {
    const tbody = shiftMappingsTable.querySelector('tbody');
    tbody.innerHTML = '';

    Object.entries(shiftMappings).forEach(([code, timeRange]) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <input type="text" class="form-control" value="${code}" readonly>
            </td>
            <td>
                <input type="text" class="form-control time-range" 
                       value="${timeRange}" 
                       pattern="([0-1][0-9]|2[0-3])[0-5][0-9]-([0-1][0-9]|2[0-3])[0-5][0-9]|OFF"
                       title="Use format HHMM-HHMM (e.g., 0900-1700) or OFF">
            </td>
        `;
        tbody.appendChild(row);
    });
}

// Add new shift mapping
document.getElementById('add-shift').addEventListener('click', () => {
    const tbody = shiftMappingsTable.querySelector('tbody');
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>
            <input type="text" class="form-control" placeholder="Shift Code">
        </td>
        <td>
            <input type="text" class="form-control time-range" 
                   placeholder="HHMM-HHMM or OFF"
                   pattern="([0-1][0-9]|2[0-3])[0-5][0-9]-([0-1][0-9]|2[0-3])[0-5][0-9]|OFF"
                   title="Use format HHMM-HHMM (e.g., 0900-1700) or OFF">
        </td>
    `;
    tbody.appendChild(row);
});

// Save shift mappings
saveShiftsButton.addEventListener('click', () => {
    const newMappings = {};
    const rows = shiftMappingsTable.querySelectorAll('tbody tr');

    rows.forEach(row => {
        const code = row.querySelector('td:first-child input').value.trim();
        const timeRange = row.querySelector('td:last-child input').value.trim();

        if (code && timeRange) {
            if (timeRange !== 'OFF' && !/^([0-1][0-9]|2[0-3])[0-5][0-9]-([0-1][0-9]|2[0-3])[0-5][0-9]$/.test(timeRange)) {
                showAlert('Invalid time format. Use HHMM-HHMM (e.g., 0900-1700) or OFF');
                return;
            }
            newMappings[code] = timeRange;
        }
    });

    fetch('/shifts', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(newMappings)
    })
    .then(response => response.json())
    .then(data => {
        if (data.error) {
            throw new Error(data.error);
        }
        shiftMappings = newMappings;
        shiftsModal.hide();
        showAlert('Shift mappings updated successfully!', 'success');
    })
    .catch(error => {
        showAlert(error.message);
    });
});

function showAlert(message, type = 'danger') {
    const alertContainer = document.getElementById('alert-container');
    const alert = document.createElement('div');
    alert.className = `alert alert-${type} alert-dismissible fade show`;
    alert.innerHTML = `
        <div class="d-flex align-items-center">
            <i class="bi ${type === 'danger' ? 'bi-exclamation-circle' : 'bi-check-circle'} me-2"></i>
            <div>${message}</div>
        </div>
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    alertContainer.innerHTML = '';
    alertContainer.appendChild(alert);
}

function setLoading(loading) {
    loadingIndicator.classList.toggle('d-none', !loading);
    uploadArea.classList.toggle('d-none', loading);
    if (!loading) {
        fileInput.value = ''; // Reset file input when loading completes
    }
}

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    handleFile(file);
});

uploadArea.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    handleFile(file);
});

function handleFile(file) {
    if (!file) return;

    if (!file.name.endsWith('.xlsx')) {
        showAlert('Please upload an Excel (.xlsx) file');
        return;
    }

    setLoading(true);
    employeeSection.classList.add('d-none');
    const formData = new FormData();
    formData.append('file', file);

    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        setLoading(false);
        if (data.error) {
            throw new Error(data.error);
        }

        // Populate employee dropdown
        employeeSelect.innerHTML = '<option value="">Choose an employee...</option>';
        data.employees.forEach(employee => {
            const option = document.createElement('option');
            option.value = employee;
            option.textContent = employee;
            employeeSelect.appendChild(option);
        });

        // Show employee selection section
        employeeSection.classList.remove('d-none');
        showAlert('File uploaded successfully! Please select an employee.', 'success');
    })
    .catch(error => {
        setLoading(false);
        showAlert(error.message);
    });
}

employeeSelect.addEventListener('change', () => {
    generateButton.disabled = !employeeSelect.value;
});

generateButton.addEventListener('click', () => {
    const employee = employeeSelect.value;
    if (!employee) return;

    setLoading(true);
    fetch('/generate-calendar', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            employee: employee
        })
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(data => {
                throw new Error(data.error || 'Failed to generate calendar');
            });
        }
        return response.blob();
    })
    .then(blob => {
        setLoading(false);
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${employee}_schedule.ics`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        showAlert('Calendar generated and downloaded successfully!', 'success');
    })
    .catch(error => {
        setLoading(false);
        showAlert(error.message);
    });
});