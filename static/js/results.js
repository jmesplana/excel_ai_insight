import { renderJevReport } from './jev-report-ui.js';
import { escapeHtml, renderMarkdown } from './rendering.js';

/** Owns result tables and chat controls. Dependencies are explicit. */
export function initResults({ getFileData, getChatDataset, streamChat, getLLMConfig, showAlert }) {
    // =========================================================================
    // NEW RESULTS VIEWER FUNCTIONALITY
    // =========================================================================

    // Global variables for results data
    let currentAnalyzedData = null;
    let currentAnalyzedFilename = null;

    // Load analyzed data into the results viewer from an in-memory object
    // of shape { sheets: { <name>: { columns, data } } }.
    function loadAnalyzedDataFromBackend(data) {
        try {
            if (!data || !data.sheets) {
                throw new Error('No analyzed data available');
            }
            currentAnalyzedData = data;
            currentAnalyzedFilename = `analyzed_${(getFileData().filename || 'data')}`;

            // Populate all tabs with data
            displayQuickInsights(data);
            renderJevReport(Object.values(data.sheets)[0], {streamChat, getLLMConfig});
            displayInteractiveDataTable(data);
            setupChatInterface(currentAnalyzedFilename);

            return data;
        } catch (error) {
            console.error('Error loading analyzed data:', error);
            showAlert('result-message', `Error loading data: ${escapeHtml(error.message)}`, 'danger');
        }
    }

    // Display quick insights (stats cards and column analysis)
    function displayQuickInsights(data) {
        const statsCardsContainer = document.getElementById('stats-cards');
        const columnInsightsContainer = document.getElementById('column-insights');

        if (!data || !data.sheets) {
            statsCardsContainer.innerHTML = '<div class="alert alert-warning">No data available</div>';
            return;
        }

        // Get the first sheet's data
        const sheetName = Object.keys(data.sheets)[0];
        const sheetData = data.sheets[sheetName];
        const columns = sheetData.columns;
        const rows = sheetData.data;

        // Create stats cards
        const totalRows = rows.length;
        const totalColumns = columns.length;

        // Count analysis columns (columns with _analysis or custom output names)
        const analysisColumns = (sheetData.outputColumns || []).length;

        statsCardsContainer.innerHTML = `
            <div class="col-md-3">
                <div class="card text-center">
                    <div class="card-body">
                        <h3 class="card-title text-primary">${totalRows}</h3>
                        <p class="card-text text-muted">Total Rows</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-center">
                    <div class="card-body">
                        <h3 class="card-title text-success">${totalColumns}</h3>
                        <p class="card-text text-muted">Total Columns</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-center">
                    <div class="card-body">
                        <h3 class="card-title text-info">${analysisColumns}</h3>
                        <p class="card-text text-muted">Analysis Columns</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-center">
                    <div class="card-body">
                        <h3 class="card-title text-warning">${Object.keys(data.sheets).length}</h3>
                        <p class="card-text text-muted">Sheets</p>
                    </div>
                </div>
            </div>
        `;

        // Create column analysis
        let columnAnalysisHTML = '<div class="row">';
        columns.forEach((col, index) => {
            // Get sample values and stats for this column
            const values = rows.map(row => row[col]).filter(v => v !== null && v !== undefined && v !== '');
            const uniqueValues = [...new Set(values)];
            const nullCount = rows.length - values.length;

            columnAnalysisHTML += `
                <div class="col-md-6 mb-3">
                    <div class="card">
                        <div class="card-body">
                            <h6 class="card-title"><i class="bi bi-bar-chart"></i> ${escapeHtml(col)}</h6>
                            <ul class="list-unstyled mb-0">
                                <li><strong>Non-null values:</strong> ${values.length}</li>
                                <li><strong>Null values:</strong> ${nullCount}</li>
                                <li><strong>Unique values:</strong> ${uniqueValues.length}</li>
                                <li><strong>Sample:</strong> ${escapeHtml(values.slice(0, 3).join(', '))}${values.length > 3 ? '...' : ''}</li>
                            </ul>
                        </div>
                    </div>
                </div>
            `;
        });
        columnAnalysisHTML += '</div>';
        columnInsightsContainer.innerHTML = columnAnalysisHTML;
    }

    // Display interactive data table with search and sorting
    function displayInteractiveDataTable(data) {
        const sheet = Object.values(data.sheets)[0];
        const container = document.getElementById('result-preview');
        const info = document.getElementById('table-info');
        const search = document.getElementById('table-search');
        let page = 0, column = null, direction = 1;
        function render() {
            const term = search.value.toLowerCase();
            let rows = sheet.data.filter(row => !term || sheet.columns.some(c => String(row[c] ?? '').toLowerCase().includes(term)));
            if (column) rows = rows.slice().sort((a, b) => String(a[column] ?? '').localeCompare(String(b[column] ?? ''), undefined, {numeric: true}) * direction);
            const pages = Math.max(1, Math.ceil(rows.length / 200));
            page = Math.min(page, pages - 1);
            info.textContent = `Showing ${Math.min(200, rows.length - page * 200)} of ${rows.length} rows (${sheet.data.length} total; matching search)`;
            container.innerHTML = `<table class="table table-striped table-bordered" id="results-data-table"><thead><tr>${sheet.columns.map(c => `<th data-column="${escapeHtml(c)}" style="cursor:pointer">${escapeHtml(c)} ↕</th>`).join('')}</tr></thead><tbody id="table-body">${rows.slice(page * 200, (page + 1) * 200).map(row => `<tr>${sheet.columns.map(c => `<td>${escapeHtml(row[c] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table><div class="d-flex gap-2"><button id="results-prev" class="btn btn-sm btn-outline-secondary" ${page === 0 ? 'disabled' : ''}>Previous</button><span>Page ${page + 1} of ${pages}</span><button id="results-next" class="btn btn-sm btn-outline-secondary" ${page + 1 === pages ? 'disabled' : ''}>Next</button></div>`;
            container.querySelectorAll('th').forEach(th => th.onclick = () => {
                direction = column === th.dataset.column ? -direction : 1; column = th.dataset.column; render();
            });
            container.querySelector('#results-prev').onclick = () => { page--; render(); };
            container.querySelector('#results-next').onclick = () => { page++; render(); };
        }
        search.value = '';
        search.oninput = () => { page = 0; render(); };
        render();
    }

    // Setup chat interface
    function setupChatInterface(filename) {
        const chatInput = document.getElementById('chat-input');
        const chatSendBtn = document.getElementById('chat-send-btn');
        const chatMessages = document.getElementById('chat-messages');

        // Handle send button click
        if (chatSendBtn) {
            chatSendBtn.onclick = async function() {
                await sendChatMessage(filename);
            };
        }

        // Handle Enter key in input
        if (chatInput) {
            chatInput.onkeypress = async function(e) {
                if (e.key === 'Enter') {
                    await sendChatMessage(filename);
                }
            };
        }
    }

    // Send chat message to backend
    async function sendChatMessage(filename) {
        const chatInput = document.getElementById('chat-input');
        const chatMessages = document.getElementById('chat-messages');
        const question = chatInput.value.trim();

        if (!question) {
            return;
        }

        // Add user message to chat
        const userMessageHTML = `
            <div class="mb-3 text-end">
                <div class="d-inline-block bg-primary text-white rounded p-2" style="max-width: 70%;">
                    <strong>You:</strong> ${escapeHtml(question)}
                </div>
            </div>
        `;
        chatMessages.innerHTML += userMessageHTML;

        // Clear input
        chatInput.value = '';

        // Scroll to bottom
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;

        // Show loading indicator
        const loadingHTML = `
            <div class="mb-3" id="chat-loading">
                <div class="d-inline-block bg-light rounded p-2">
                    <span class="spinner-border spinner-border-sm"></span> AI is thinking...
                </div>
            </div>
        `;
        chatMessages.innerHTML += loadingHTML;
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;

        try {
            // Credentials may come from the server environment, so we do
            // not hard-block here; the server returns a clear error if not.

            const dataset = getChatDataset();
            if (!dataset) throw new Error('No data loaded. Please upload a file first.');

            // Remove loading indicator and prepare a streaming target.
            document.getElementById('chat-loading')?.remove();
            const aiMessageDiv = document.createElement('div');
            aiMessageDiv.className = 'mb-3';
            aiMessageDiv.innerHTML = `
                <div class="d-inline-block bg-light rounded p-2" style="max-width: 70%;">
                    <strong>AI:</strong>
                    <div class="markdown-content"></div>
                </div>
            `;
            chatMessages.appendChild(aiMessageDiv);
            const contentEl = aiMessageDiv.querySelector('.markdown-content');

            await streamChat(
                { ...dataset, question: question, ...getLLMConfig() },
                (full) => {
                    contentEl.innerHTML = renderMarkdown(full);
                    chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;
                }
            );

        } catch (error) {
            // Remove loading indicator
            document.getElementById('chat-loading')?.remove();

            // Show error message
            const errorHTML = `
                <div class="mb-3">
                    <div class="d-inline-block bg-danger text-white rounded p-2" style="max-width: 70%;">
                        <strong>Error:</strong> ${escapeHtml(error.message)}
                    </div>
                </div>
            `;
            chatMessages.innerHTML += errorHTML;
        }

        // Scroll to bottom
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;
    }

    // Make loadAnalyzedDataFromBackend available globally


    // =========================================================================
    // PRE-ANALYSIS CHAT FUNCTIONALITY (Step 3)
    // =========================================================================

    // Setup preview chat when user uploads a file
    function setupPreviewChat() {
        const previewChatInput = document.getElementById('preview-chat-input');
        const previewChatSendBtn = document.getElementById('preview-chat-send-btn');

        if (previewChatSendBtn) {
            previewChatSendBtn.onclick = async function() {
                await sendPreviewChatMessage();
            };
        }

        if (previewChatInput) {
            previewChatInput.addEventListener('keypress', async function(e) {
                if (e.key === 'Enter') {
                    await sendPreviewChatMessage();
                }
            });
        }
    }

    // Send chat message about uploaded (un-analyzed) data
    async function sendPreviewChatMessage() {
        const chatInput = document.getElementById('preview-chat-input');
        const chatMessages = document.getElementById('preview-chat-messages');
        const question = chatInput.value.trim();

        if (!question) {
            return;
        }

        // Check if we have uploaded file data
        if (!getFileData() || !getFileData().filename) {
            const errorHTML = `
                <div class="mb-3">
                    <div class="d-inline-block bg-danger text-white rounded p-2">
                        <strong>Error:</strong> No file uploaded. Please upload a file first.
                    </div>
                </div>
            `;
            chatMessages.innerHTML += errorHTML;
            return;
        }

        // Add user message to chat
        const userMessageHTML = `
            <div class="mb-3 text-end">
                <div class="d-inline-block bg-primary text-white rounded p-2" style="max-width: 70%;">
                    <strong>You:</strong> ${escapeHtml(question)}
                </div>
            </div>
        `;
        chatMessages.innerHTML += userMessageHTML;

        // Clear input
        chatInput.value = '';

        // Scroll to bottom
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;

        // Show loading indicator
        const loadingHTML = `
            <div class="mb-3" id="preview-chat-loading">
                <div class="d-inline-block bg-light rounded p-2">
                    <span class="spinner-border spinner-border-sm"></span> AI is thinking...
                </div>
            </div>
        `;
        chatMessages.innerHTML += loadingHTML;
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;

        try {
            // Credentials may come from the server environment, so we do
            // not hard-block here; the server returns a clear error if not.

            const dataset = getChatDataset();
            if (!dataset) throw new Error('No data loaded. Please upload a file first.');

            // Remove loading indicator and prepare a streaming target.
            document.getElementById('preview-chat-loading')?.remove();
            const aiMessageDiv = document.createElement('div');
            aiMessageDiv.className = 'mb-3';
            aiMessageDiv.innerHTML = `
                <div class="d-inline-block bg-light rounded p-2" style="max-width: 70%;">
                    <strong>AI:</strong>
                    <div class="markdown-content"></div>
                </div>
            `;
            chatMessages.appendChild(aiMessageDiv);
            const contentEl = aiMessageDiv.querySelector('.markdown-content');

            await streamChat(
                { ...dataset, question: question, ...getLLMConfig() },
                (full) => {
                    contentEl.innerHTML = renderMarkdown(full);
                    chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;
                }
            );

        } catch (error) {
            // Remove loading indicator
            document.getElementById('preview-chat-loading')?.remove();

            // Show error message
            const errorHTML = `
                <div class="mb-3">
                    <div class="d-inline-block bg-danger text-white rounded p-2" style="max-width: 70%;">
                        <strong>Error:</strong> ${escapeHtml(error.message)}
                    </div>
                </div>
            `;
            chatMessages.innerHTML += errorHTML;
        }

        // Scroll to bottom
        chatMessages.parentElement.scrollTop = chatMessages.parentElement.scrollHeight;
    }

    // Initialize preview chat when file is uploaded
    setupPreviewChat();

    // =========================================================================
    // FLOATING CHAT MODAL FUNCTIONALITY
    // =========================================================================

    // Helper function to escape HTML


    const floatingChatBtn = document.getElementById('floating-chat-btn');
    const chatModal = new bootstrap.Modal(document.getElementById('chat-modal'));
    const modalChatInput = document.getElementById('modal-chat-input');
    const modalChatSendBtn = document.getElementById('modal-chat-send-btn');
    const modalChatMessages = document.getElementById('modal-chat-messages');
    const quickQuestionsContainer = document.getElementById('quick-questions');

    // Track current file for chat
    let currentChatFilename = null;
    let isAnalyzedData = false;

    // Show floating button when file is uploaded
    function showFloatingChatButton(filename, analyzed = false) {
        currentChatFilename = filename;
        isAnalyzedData = analyzed;
        floatingChatBtn.classList.remove('hidden');

        // Update quick questions based on context
        updateQuickQuestions(analyzed);
    }

    // Quick questions based on context
    function updateQuickQuestions(analyzed) {
        const questions = analyzed ? [
            "What are the main trends in the analyzed data?",
            "Summarize the key findings",
            "What insights can you provide?",
            "How many rows were analyzed?",
            "What columns contain analysis results?"
        ] : [
            "What columns are in this dataset?",
            "How many rows and columns are there?",
            "Show me a summary of the data",
            "What data types are present?",
            "Are there any null values?"
        ];

        quickQuestionsContainer.innerHTML = questions.map(q =>
            `<button class="btn btn-sm btn-outline-primary quick-question-btn" data-question="${escapeHtml(q)}">${escapeHtml(q)}</button>`
        ).join('');

        // Add click handlers to quick question buttons
        document.querySelectorAll('.quick-question-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                modalChatInput.value = this.dataset.question;
                sendModalChatMessage();
            });
        });
    }

    // Open chat modal
    floatingChatBtn.addEventListener('click', function() {
        chatModal.show();
    });

    // Handle modal chat send
    modalChatSendBtn.addEventListener('click', sendModalChatMessage);
    modalChatInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            sendModalChatMessage();
        }
    });

    // Send message in modal
    async function sendModalChatMessage() {
        const question = modalChatInput.value.trim();

        if (!question) {
            return;
        }

        if (!currentChatFilename) {
            alert('No file loaded. Please upload a file first.');
            return;
        }

        // Add user message
        const userMessageDiv = document.createElement('div');
        userMessageDiv.className = 'mb-3 text-end';
        userMessageDiv.innerHTML = `
            <div class="d-inline-block bg-primary text-white rounded p-2" style="max-width: 70%;">
                <strong>You:</strong> ${escapeHtml(question)}
            </div>
        `;
        modalChatMessages.appendChild(userMessageDiv);

        // Clear input
        modalChatInput.value = '';

        // Scroll to bottom
        const container = document.getElementById('modal-chat-container');
        container.scrollTop = container.scrollHeight;

        // Credentials may come from the server environment, so we do not
        // hard-block here; the server returns a clear error if not.

        // Remove any existing streaming response
        const existingStreamingResponse = document.getElementById('streaming-response');
        if (existingStreamingResponse) {
            existingStreamingResponse.remove();
        }

        // Create AI response container (for streaming) with unique ID
        const responseId = `streaming-response-${Date.now()}`;
        const aiMessageDiv = document.createElement('div');
        aiMessageDiv.className = 'mb-3';
        aiMessageDiv.id = responseId;
        aiMessageDiv.innerHTML = `
            <div class="d-inline-block bg-light rounded p-3" style="max-width: 80%;">
                <strong>AI:</strong>
                <div class="markdown-content mt-2" id="streaming-content-${Date.now()}">
                    <span class="spinner-border spinner-border-sm"></span> Analyzing...
                </div>
            </div>
        `;
        modalChatMessages.appendChild(aiMessageDiv);
        container.scrollTop = container.scrollHeight;

        // Store the content div ID for this message
        const contentDivId = aiMessageDiv.querySelector('.markdown-content').id;

        try {
            const dataset = getChatDataset();
            if (!dataset) throw new Error('No data loaded. Please upload a file first.');

            await streamChat({...dataset, question, ...getLLMConfig()}, full => {
                const target = document.getElementById(contentDivId);
                if (target) target.innerHTML = renderMarkdown(full);
                container.scrollTop = container.scrollHeight;
            });

        } catch (error) {
            console.error('Chat error:', error);

            // Remove the streaming response div using the unique ID
            document.getElementById(responseId)?.remove();

            // Show error
            const errorDiv = document.createElement('div');
            errorDiv.className = 'mb-3';
            errorDiv.innerHTML = `
                <div class="d-inline-block bg-danger text-white rounded p-2" style="max-width: 70%;">
                    <strong>Error:</strong> ${escapeHtml(error.message)}
                </div>
            `;
            modalChatMessages.appendChild(errorDiv);
        }

        // Scroll to bottom
        container.scrollTop = container.scrollHeight;
    }


    return {
        load: loadAnalyzedDataFromBackend,
        onStep(step) {
    if (step === 3 && getFileData()?.filename) showFloatingChatButton(getFileData().filename, false);
    if (step === 5 && currentAnalyzedFilename) showFloatingChatButton(currentAnalyzedFilename, true);
        }
    };
}
