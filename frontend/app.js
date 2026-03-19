document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');

    // Sections
    const uploadSection = document.getElementById('upload-card');
    const processingSection = document.getElementById('processing-section');
    const successSection = document.getElementById('success-section');
    const errorSection = document.getElementById('error-section');

    // Dynamic content
    const processedFilename = document.getElementById('processed-filename');
    const errorMessage = document.getElementById('error-message');

    // Action buttons
    const downloadBtn = document.getElementById('download-btn');
    const resetBtn = document.getElementById('reset-btn');
    const errorResetBtn = document.getElementById('error-reset-btn');

    // State
    let currentBlobUrl = null;
    let currentOriginalFilename = null;

    // --- Drag and Drop Events ---

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop zone when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    dropZone.addEventListener('drop', handleDrop, false);

    // Click to browse
    dropZone.addEventListener('click', (e) => {
        // Prevent triggering twice if they click the button itself
        if (e.target !== browseBtn) {
            fileInput.click();
        }
    });

    browseBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    fileInput.addEventListener('change', handleFileSelect, false);

    // --- Action Button Events ---

    downloadBtn.addEventListener('click', () => {
        if (currentBlobUrl && currentOriginalFilename) {
            const a = document.createElement('a');
            a.href = currentBlobUrl;
            a.download = `Basin_Electric_Costs_${currentOriginalFilename}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
    });

    resetBtn.addEventListener('click', resetApp);
    errorResetBtn.addEventListener('click', resetApp);


    // --- Core Functions ---

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight(e) {
        dropZone.classList.add('drag-active');
    }

    function unhighlight(e) {
        dropZone.classList.remove('drag-active');
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }

    function handleFileSelect(e) {
        const files = e.target.files;
        handleFiles(files);
    }

    function handleFiles(files) {
        if (files.length === 0) return;

        const file = files[0];
        // Validate file type
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            showError("Please upload a valid Excel file (.xlsx or .xls)");
            return;
        }

        currentOriginalFilename = file.name;
        uploadFile(file);
    }

    function uploadFile(file) {
        // Show processing state
        uploadSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        processingSection.classList.remove('hidden');

        const formData = new FormData();
        formData.append('file', file);

        fetch('/api/process', {
            method: 'POST',
            body: formData
        })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(errData => {
                        throw new Error(errData.detail || "Server error while processing");
                    }).catch(() => {
                        throw new Error(`Server returned ${response.status} ${response.statusText}`);
                    });
                }
                return response.blob();
            })
            .then(blob => {
                // Create a URL for the downloaded file
                if (currentBlobUrl) {
                    URL.revokeObjectURL(currentBlobUrl);
                }
                currentBlobUrl = URL.createObjectURL(blob);

                // Show success state
                processingSection.classList.add('hidden');
                processedFilename.textContent = file.name;
                successSection.classList.remove('hidden');
            })
            .catch(error => {
                showError(error.message);
            });
    }

    function showError(msg) {
        uploadSection.classList.add('hidden');
        processingSection.classList.add('hidden');
        successSection.classList.add('hidden');

        errorMessage.textContent = msg;
        errorSection.classList.remove('hidden');
    }

    function resetApp() {
        // Clear file input
        fileInput.value = '';

        // Cleanup old blob if it exists
        if (currentBlobUrl) {
            URL.revokeObjectURL(currentBlobUrl);
            currentBlobUrl = null;
        }

        // Show upload section
        successSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        processingSection.classList.add('hidden');
        uploadSection.classList.remove('hidden');
    }
});
