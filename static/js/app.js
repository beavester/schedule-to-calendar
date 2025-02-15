let currentFilePath = null;

// Setup drag and drop handlers
const uploadArea = document.querySelector('.upload-area');
const fileInput = document.getElementById('file-input');
const employeeSection = document.getElementById('employee-section');
const employeeSelect = document.getElementById('employee-select');
const generateButton = document.getElementById('generate-calendar');
const loadingIndicator = document.getElementById('loading');

function showAlert(message, type = 'danger') {
    const alertContainer = document.getElementById('alert-container');
    const alert = document.createElement('div');
    alert.className = `alert alert-${type} alert-dismissible fade show`;
    alert.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    alertContainer.innerHTML = '';
    alertContainer.appendChild(alert);
}

function setLoading(loading) {
    loadingIndicator.classList.toggle('d-none', !loading);
    uploadArea.classList.toggle('d-none', loading);
}

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('border-primary');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('border-primary');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('border-primary');
    const file = e.dataTransfer.files[0];
    handleFile(file);
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
        showAlert('File uploaded successfully!', 'success');
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
        showAlert('Calendar file generated successfully!', 'success');
    })
    .catch(error => {
        setLoading(false);
        showAlert(error.message);
    });
});