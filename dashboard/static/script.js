const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileNameDisplay = document.getElementById('file-name');
const processBtn = document.getElementById('process-btn');
const apiKeyInput = document.getElementById('api-key');
const resultSection = document.getElementById('result-section');
const resultTableBody = document.querySelector('#result-table tbody');
const downloadBtn = document.getElementById('download-btn');
const toast = document.getElementById('toast');

let selectedFile = null;

// Drag & Drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
});

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

function handleFiles(files) {
    if (files.length > 0) {
        const file = files[0];
        if (file.type === 'application/pdf') {
            selectedFile = file;
            fileNameDisplay.textContent = `Selected: ${file.name}`;
            validateForm();
        } else {
            showToast('Please upload a PDF file', 'error');
        }
    }
}

apiKeyInput.addEventListener('input', validateForm);

function validateForm() {
    processBtn.disabled = !(selectedFile && apiKeyInput.value.trim());
}

processBtn.addEventListener('click', async () => {
    if (!selectedFile || !apiKeyInput.value.trim()) return;

    setLoading(true);
    resultSection.classList.add('hidden');

    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('api_key', apiKeyInput.value.trim());

    try {
        const response = await fetch('/process', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok) {
            displayResults(result.data);
            downloadBtn.href = result.download_url;
            resultSection.classList.remove('hidden');
            showToast('Extraction complete!', 'success');
        } else {
            showToast(result.error || 'Extraction failed', 'error');
        }
    } catch (error) {
        showToast('Network error occurred', 'error');
        console.error(error);
    } finally {
        setLoading(false);
    }
});

function displayResults(data) {
    resultTableBody.innerHTML = '';
    data.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item['#']}</td>
            <td>${item.key}</td>
            <td>${item.value}</td>
            <td>${item.comments}</td>
        `;
        resultTableBody.appendChild(row);
    });
}

function setLoading(isLoading) {
    if (isLoading) {
        processBtn.classList.add('loading');
        processBtn.disabled = true;
    } else {
        processBtn.classList.remove('loading');
        processBtn.disabled = false;
    }
}

function showToast(message, type = 'success') {
    toast.textContent = message;
    toast.className = `toast show ${type}`;
    setTimeout(() => {
        toast.className = 'toast';
    }, 3000);
}
