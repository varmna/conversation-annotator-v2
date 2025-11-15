// Initialize main variables
const tool = {
    currentIndex: 0,
    conversations: [],
    annotations: {},
    buckets: [
        "Bot Response",
        "HVA",
        "AB feature/HVA Related Query",
        "Personalized/Account-Specific Queries",
        "Promo & Freebie Related Queries",
        "Help-page/Direct Customer Service",
        "BP for Non-Profit Organisation Related Query",
        "Personal Prime Related Query",
        "Customer Behavior",
        "Other Queries",
        "Overall Observations"
    ]
};

// Initialize UI elements
const elements = {
    uploadScreen: document.getElementById('upload-screen'),
    mainInterface: document.getElementById('main-interface'),
    fileInput: document.getElementById('excel-upload'),
    uploadBox: document.getElementById('upload-box'),
    conversationDisplay: document.getElementById('conversation-display'),
    conversationInfo: document.getElementById('conversation-info'),
    bucketArea: document.getElementById('bucket-area'),
    prevBtn: document.getElementById('prev-btn'),
    nextBtn: document.getElementById('next-btn'),
    saveBtn: document.getElementById('save-btn'),
    downloadBtn: document.getElementById('download-btn'),
    progress: document.getElementById('progress'),
    progressText: document.getElementById('progress-text'),
    statusMessage: document.getElementById('status-message'),
    loadingSpinner: document.getElementById('loading-spinner')
};

// Create bucket UI
function createBucketUI() {
    tool.buckets.forEach(bucket => {
        const bucketHTML = `
            <div class="bucket">
                <label class="bucket-label">
                    <input type="checkbox" name="${bucket}">
                    ${bucket}
                </label>
                <textarea 
                    placeholder="Add comments for ${bucket}" 
                    name="${bucket}"
                    rows="3"
                ></textarea>
            </div>
        `;
        elements.bucketArea.insertAdjacentHTML('beforeend', bucketHTML);
    });
}

// File upload handler
elements.fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    if (!file.name.endsWith('.xlsx')) {
        showStatus('❌ Please select an Excel (.xlsx) file', 'error');
        return;
    }

    try {
        showLoading(true);
        showStatus('📂 Loading file...', 'info');

        const data = await readExcelFile(file);
        
        if (!data || data.length === 0) {
            throw new Error('No data found in file');
        }

        processExcelData(data);
        elements.uploadScreen.style.display = 'none';
        elements.mainInterface.style.display = 'flex';
        showStatus('✅ File loaded successfully!', 'success');
    } catch (error) {
        console.error('Error:', error);
        showStatus('❌ Error loading file: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
});

// Read Excel file
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'a 'array'});
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                resolve(XLSX.utils.sheet_to_json(sheet));
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process Excel data
function processExcelData(rawData) {
    // Group conversations by Id
    const groupedData = {};
    rawData.forEach(row => {
        if (!groupedData[row.Id]) {
            groupedData[row.Id] = [];
        }
        groupedData[row.Id].push(row);
    });
    
    tool.conversations = Object.values(groupedData);
    tool.currentIndex = 0;
    tool.annotations = {};
    
    updateProgressBar();
    displayConversation();
}

// Display conversation
function displayConversation() {
    const conv = tool.conversations[tool.currentIndex];
    const lastMessage = conv[conv.length - 1];

    // Update conversation info
    elements.conversationInfo.innerHTML = `
        <div class="info-item">
            <strong>ID:</strong> ${conv[0].Id}
        </div>
        <div class="info-item">
            <strong>Feedback:</strong> 
            <span class="badge ${lastMessage['Customer Feedback']?.toLowerCase() === 'negative' ? 'bg-danger' : 'bg-success'}">
                ${lastMessage['Customer Feedback'] || 'N/A'}
            </span>
        </div>
    `;

    // Display messages
    let html = '<div class="messages">';
    conv.forEach(message => {
        if (message.llmGeneratedUserMessage) {
            html += `
                <div class="message customer">
                    <div class="message-header">👤 Customer</div>
                    ${message.llmGeneratedUserMessage}
                </div>
            `;
        }
        if (message.botMessage) {
            html += `
                <div class="message bot">
                    <div class="message-header">🤖 Bot</div>
                    ${message.botMessage}
                </div>
            `;
        }
    });
    html += '</div>';

    elements.conversationDisplay.innerHTML = html;
    updateProgressBar();
    loadAnnotations();
}

// Update progress bar
function updateProgressBar() {
    const progress = ((tool.currentIndex + 1) / tool.conversations.length) * 100;
    elements.progress.style.width = `${progress}%`;
    elements.progressText.textContent = 
        `${tool.currentIndex + 1}/${tool.conversations.length} Conversations`;
}

// Save annotations
function saveCurrentAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const hasAnnotations = tool.buckets.some(bucket => 
        document.querySelector(`input[name="${bucket}"]`).checked
    );

    if (!hasAnnotations) {
        showStatus('⚠️ Please select at least one bucket', 'warning');
        return;
    }

    tool.annotations[convId] = {};
    
    tool.buckets.forEach(bucket => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        if (checkbox.checked) {
            tool.annotations[convId][bucket] = textarea.value.trim();
        }
    });

    showStatus('✅ Annotations saved!', 'success');
}

// Load annotations
function loadAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const savedAnnotations = tool.annotations[convId] || {};
    
    tool.buckets.forEach(bucket => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        checkbox.checked = false;
        textarea.value = '';
        textarea.disabled = true;
    });

    Object.entries(savedAnnotations).forEach(([bucket, comment]) => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        const textarea = document.querySelector(`textarea[name="${bucket}"]`);
        if (checkbox && textarea) {
            checkbox.checked = true;
            textarea.value = comment;
            textarea.disabled = false;
        }
    });
}

// Enable/disable textarea based on checkbox
elements.bucketArea.addEventListener('change', (e) => {
    if (e.target.type === 'checkbox') {
        const textarea = e.target.closest('.bucket').querySelector('textarea');
        textarea.disabled = !e.target.checked;
        if (e.target.checked) {
            textarea.focus();
        }
    }
});

// Show status message
function showStatus(message, type) {
    elements.statusMessage.textContent = message;
    elements.statusMessage.className = `status-message alert alert-${type}`;
    elements.statusMessage.style.display = 'block';
    
    setTimeout(() => {
        elements.statusMessage.style.display = 'none';
    }, 3000);
}

// Show/hide loading spinner
function showLoading(show) {
    elements.loadingSpinner.style.display = show ? 'flex' : 'none';
}

// Navigation handlers
elements.prevBtn.addEventListener('click', () => {
    if (tool.currentIndex > 0) {
        tool.currentIndex--;
        displayConversation();
    } else {
        showStatus('⚠️ This is the first conversation', 'warning');
    }
});

elements.nextBtn.addEventListener('click', () => {
    if (tool.currentIndex < tool.conversations.length - 1) {
        tool.currentIndex++;
        displayConversation();
    } else {
        showStatus('⚠️ This is the last conversation', 'warning');
    }
});

elements.saveBtn.addEventListener('click', saveCurrentAnnotations);

// Download handler
elements.downloadBtn.addEventListener('click', () => {
    try {
        if (Object.keys(tool.annotations).length === 0) {
            showStatus('⚠️ No annotations to download', 'warning');
            return;
        }

        showLoading(true);
        showStatus('💾 Preparing download...', 'info');
        
        const annotatedData = [];
        
        tool.conversations.forEach(conv => {
            const convId = conv[0].Id;
            const savedAnnotations = tool.annotations[convId];
            
            if (savedAnnotations && Object.keys(savedAnnotations).length > 0) {
                conv.forEach((message, index) => {
                    const isFirstMessage = index === 0;
                    const isLastMessage = index === conv.length - 1;
                    
                    const row = {
                        'Id': message.Id,
                        'llmGeneratedUserMessage': message.llmGeneratedUserMessage || '',
                        'botMessage': message.botMessage || '',
                        'Customer Feedback': isLastMessage ? message['Customer Feedback'] || '' : ''
                    };

                    if (isFirstMessage) {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = savedAnnotations[bucket] || '';
                        });
                    } else {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = '';
                        });
                    }
                    
                    annotatedData.push(row);
                });
            }
        });

        // Create Excel file
        const ws = XLSX.utils.json_to_sheet(annotatedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Annotations");

        // Convert to binary string
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);

        // Trigger download
        const a = document.createElement('a');
        a.href = url;
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        a.download = `annotated_conversations_${timestamp}.xlsx`;
        document.body.appendChild(a);
        a.click();
        
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        const annotatedCount = new Set(annotatedData.map(row => row.Id)).size;
        showStatus(`✅ Downloaded ${annotatedCount} conversation(s)!`, 'success');
    } catch (error) {
        console.error('Download error:', error);
        showStatus('❌ Error downloading file', 'error');
    } finally {
        showLoading(false);
    }
});

// Helper function for Excel binary conversion
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

// Initialize
createBucketUI();

// Drag and drop support
elements.uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.uploadBox.classList.add('dragover');
});

elements.uploadBox.addEventListener('dragleave', () => {
    elements.uploadBox.classList.remove('dragover');
});

elements.uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.uploadBox.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        elements.fileInput.files = e.dataTransfer.files;
        elements.fileInput.dispatchEvent(new Event('change'));
    } else {
        showStatus('❌ Please select an Excel (.xlsx) file', 'error');
    }
});
