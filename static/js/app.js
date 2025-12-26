// Simple Schedule to Calendar App
document.addEventListener('DOMContentLoaded', () => {
    const uploadBox = document.getElementById('upload-box');
    const fileInput = document.getElementById('file-input');
    const stepUpload = document.getElementById('step-upload');
    const stepSelect = document.getElementById('step-select');
    const loading = document.getElementById('loading');
    const employeeSelect = document.getElementById('employee-select');
    const btnDownload = document.getElementById('btn-download');
    const btnRestart = document.getElementById('btn-restart');
    const message = document.getElementById('message');

    // Tap upload box to open file picker
    uploadBox.addEventListener('click', () => fileInput.click());

    // Handle file selection
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            uploadFile(e.target.files[0]);
        }
    });

    // Enable download button when name selected
    employeeSelect.addEventListener('change', () => {
        btnDownload.disabled = !employeeSelect.value;
    });

    // Download calendar
    btnDownload.addEventListener('click', downloadCalendar);

    // Start over
    btnRestart.addEventListener('click', () => {
        fileInput.value = '';
        employeeSelect.innerHTML = '<option value="">Choose...</option>';
        btnDownload.disabled = true;
        showStep('upload');
        hideMessage();
    });

    function showStep(step) {
        stepUpload.classList.add('hidden');
        stepSelect.classList.add('hidden');
        loading.classList.add('hidden');

        if (step === 'upload') stepUpload.classList.remove('hidden');
        if (step === 'select') stepSelect.classList.remove('hidden');
        if (step === 'loading') loading.classList.remove('hidden');
    }

    function showMessage(text, isError = false) {
        message.textContent = text;
        message.className = 'message ' + (isError ? 'error' : 'success');
        message.classList.remove('hidden');
    }

    function hideMessage() {
        message.classList.add('hidden');
    }

    async function uploadFile(file) {
        if (!file.name.match(/\.xlsx?$/i)) {
            showMessage('Please select an Excel file (.xlsx)', true);
            return;
        }

        showStep('loading');
        hideMessage();

        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok || data.error) {
                throw new Error(data.error || 'Upload failed');
            }

            // Populate dropdown
            employeeSelect.innerHTML = '<option value="">Choose...</option>';
            data.employees.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                employeeSelect.appendChild(opt);
            });

            showStep('select');
            showMessage('File uploaded! Select your name below.');

        } catch (err) {
            showStep('upload');
            showMessage(err.message || 'Failed to process file', true);
        }
    }

    async function downloadCalendar() {
        const employee = employeeSelect.value;
        if (!employee) return;

        showStep('loading');

        try {
            const response = await fetch('/generate-calendar', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ employee })
            });

            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'Failed to generate calendar');
            }

            // Download the file
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${employee.replace(/[^a-z0-9]/gi, '_')}_schedule.ics`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);

            showStep('select');
            showMessage('Calendar downloaded! Add it to your calendar app.');

        } catch (err) {
            showStep('select');
            showMessage(err.message || 'Failed to download calendar', true);
        }
    }
});
