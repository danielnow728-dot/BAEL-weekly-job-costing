document.addEventListener('DOMContentLoaded', () => {
    const loader = document.getElementById('history-loader');
    const libraryContent = document.getElementById('library-content');
    const listProcessed = document.getElementById('history-list');
    const listRaw = document.getElementById('raw-list');
    const empty = document.getElementById('history-empty');

    fetch('/api/history')
        .then(res => {
            if (!res.ok) throw new Error("Failed to fetch history");
            return res.json();
        })
        .then(data => {
            loader.classList.add('hidden');
            if (data.files && data.files.length > 0) {
                libraryContent.classList.remove('hidden');
                
                data.files.forEach(file => {
                    const li = document.createElement('li');
                    li.className = 'history-item';
                    
                    const date = new Date(file.date).toLocaleString();
                    const size = (file.size / 1024).toFixed(1) + ' KB';

                    li.innerHTML = `
                        <a href="/api/history/${encodeURIComponent(file.name)}" class="history-item-link" download>
                            <div class="history-info">
                                <i data-feather="file-text" class="file-icon"></i>
                                <div class="history-details">
                                    <h4>${file.name}</h4>
                                    <span>${date} &bull; ${size}</span>
                                </div>
                            </div>
                            <div class="download-action">
                                <i data-feather="download"></i> <span>Download</span>
                            </div>
                        </a>
                    `;
                    
                    if (file.type === 'raw') {
                        listRaw.appendChild(li);
                    } else {
                        listProcessed.appendChild(li);
                    }
                });
                
                // Hide lists if empty
                if (listRaw.children.length === 0) listRaw.innerHTML = "<li><span style='color: var(--clr-text-muted); font-size: 0.85rem'>No raw files found.</span></li>";
                if (listProcessed.children.length === 0) listProcessed.innerHTML = "<li><span style='color: var(--clr-text-muted); font-size: 0.85rem'>No processed files found.</span></li>";

                // Initialize icons for the newly injected HTML
                if (typeof feather !== 'undefined') {
                    feather.replace();
                }
            } else {
                empty.classList.remove('hidden');
            }
        })
        .catch(err => {
            loader.innerHTML = `<p style="color:var(--clr-error)">Error loading history: ${err.message}</p>`;
        });
});
