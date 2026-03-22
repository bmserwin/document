let fileBuffer = null;
let variables = []; // { id, name, originalText }
let rowData = [{}]; // Data for each student
let lastGeneratedZipBlob = null;
let lastGeneratedDocs = []; // Array of { name, buffer }
let lastGeneratedFile = null; // Pre-calculated File object for sharing

// Selectors
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const docPreview = document.getElementById('doc-preview');
const varModal = document.getElementById('var-modal');
const overlay = document.getElementById('overlay');
const varNameInput = document.getElementById('var-name-input');
const varsListDisplay = document.getElementById('vars-list');

// Phase Switching
function go(phaseNum) {
    document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
    document.getElementById(`p${phaseNum}`).classList.add('active');

    document.querySelectorAll('.step').forEach((s, idx) => {
        s.classList.remove('active', 'completed');
        if (idx + 1 < phaseNum) s.classList.add('completed');
        if (idx + 1 === phaseNum) s.classList.add('active');
    });

    if (phaseNum === 3) buildDataGrid();
}

// Phase 1: Upload
dropZone.onclick = () => fileInput.click();
dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('hover'); };
dropZone.ondragleave = () => dropZone.classList.remove('hover');
dropZone.ondrop = (e) => {
    e.preventDefault();
    handleFile(e.dataTransfer.files[0]);
};

fileInput.onchange = (e) => handleFile(e.target.files[0]);

async function handleFile(file) {
    if (!file || !file.name.endsWith('.docx')) {
        alert("Please upload a .docx file.");
        return;
    }
    const reader = new FileReader();
    reader.onload = async (e) => {
        fileBuffer = e.target.result;
        await renderDocPreview();
        go(2);
    };
    reader.readAsArrayBuffer(file);
}

async function renderDocPreview() {
    const result = await mammoth.convertToHtml({ arrayBuffer: fileBuffer });
    docPreview.innerHTML = result.value;
}

// Phase 2: Selection
let selectionTimeout;
const selectionBar = document.getElementById('selection-bar');
const selectionInfo = document.getElementById('selection-text-info');
let lastSelectedText = "";

function handleSelection() {
    const selection = window.getSelection();
    const text = selection.toString().trim();

    // Check if selection is inside docPreview
    if (text.length > 0 && docPreview.contains(selection.anchorNode)) {
        lastSelectedText = text;
        const isMobile = window.innerWidth <= 768;

        if (isMobile) {
            selectionInfo.textContent = `"${text.substring(0, 15)}${text.length > 15 ? '...' : ''}"`;
            selectionBar.style.display = 'flex';
        } else {
            document.getElementById('selected-text-preview').textContent = `"${text}"`;
            showModal(text);
        }
    } else {
        selectionBar.style.display = 'none';
    }
}

// Explicit trigger from Selection Bar
function triggerVarModal() {
    document.getElementById('selected-text-preview').textContent = `"${lastSelectedText}"`;
    showModal(lastSelectedText);
    selectionBar.style.display = 'none';
}

/**
 * AUTO-SCAN METHOD:
 * Automatically finds markers like {{variable}} or $$variable$$ in the document text
 * This is the easiest method for mobile users.
 */
function scanForMarkers() {
    const text = docPreview.innerText;
    // Regex matches {{name}} or $$name$$
    const markers = text.match(/\{\{([^}]+)\}\}|\$\$([^$]+)\$\$/g);

    if (!markers) {
        alert("No markers like {{name}} or $$name$$ found in the document.");
        return;
    }

    let foundAny = false;
    markers.forEach(m => {
        // Extract name and clean up
        const name = m.replace(/\{\{|\}\}|\$\$/g, '').trim().replace(/\s+/g, '_');
        if (!variables.find(v => v.name === name)) {
            variables.push({ id: Date.now() + Math.random(), name, originalText: m });
            highlightPreview(m, name);
            foundAny = true;
        }
    });

    if (foundAny) {
        updateVarsSidebar();
        alert(`Successfully found and added ${markers.length} variables!`);
    } else {
        alert("All markers found are already added to your list.");
    }
}

// Global selection change handling
document.onselectionchange = () => {
    if (document.getElementById('p2').classList.contains('active')) {
        clearTimeout(selectionTimeout);
        selectionTimeout = setTimeout(() => {
            const selection = window.getSelection();
            if (selection.toString().trim().length === 0) {
                selectionBar.style.display = 'none';
            } else if (window.innerWidth <= 768) {
                handleSelection();
            }
        }, 400);
    }
};

docPreview.onmouseup = () => {
    if (window.innerWidth > 768) handleSelection();
};

// Also trigger on touch end for mobile to catch selection adjustments
docPreview.ontouchend = () => {
    if (window.innerWidth <= 768) {
        clearTimeout(selectionTimeout);
        selectionTimeout = setTimeout(handleSelection, 200);
    }
};

function showModal(text) {
    varModal.style.display = 'block';
    overlay.style.display = 'block';
    varNameInput.value = '';
    varNameInput.focus();
    varModal.dataset.text = text;
}

function hideModal() {
    varModal.style.display = 'none';
    overlay.style.display = 'none';
    window.getSelection().removeAllRanges();
}

document.getElementById('cancel-var-btn').onclick = hideModal;

document.getElementById('save-var-btn').onclick = () => {
    const name = varNameInput.value.trim().replace(/\s+/g, '_');
    const originalText = varModal.dataset.text;

    if (!name) return alert("Enter a variable name");
    if (variables.find(v => v.name === name)) return alert("Variable name already exists");

    variables.push({ id: Date.now(), name, originalText });
    updateVarsSidebar();
    highlightPreview(originalText, name);
    hideModal();
};

function updateVarsSidebar() {
    if (variables.length === 0) {
        varsListDisplay.innerHTML = `<div style="text-align: center; padding: 40px 10px; color: var(--text-muted); font-style: italic;">No variables yet. Highlight some text!</div>`;
        return;
    }
    varsListDisplay.innerHTML = variables.map(v => `
        <div class="variable-item">
            <div class="variable-info">
                <div class="variable-name">{${v.name}}</div>
                <span>Matches: "${v.originalText}"</span>
            </div>
            <button class="btn-ghost" onclick="removeVar(${v.id})"><i data-lucide="trash-2" style="width:14px"></i></button>
        </div>
    `).join('');
    lucide.createIcons();
}

function highlightPreview(text, varName) {
    const html = docPreview.innerHTML;
    const escaped = text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, 'g');
    docPreview.innerHTML = html.replace(regex, `<mark title="Variable: ${varName}">${text}</mark>`);
}

function removeVar(id) {
    variables = variables.filter(v => v.id !== id);
    updateVarsSidebar();
    renderDocPreview().then(() => {
        variables.forEach(v => highlightPreview(v.originalText, v.name));
    });
}

// Phase 3: Data Entry
function buildDataGrid() {
    refreshTableBody();
    setupGridDelegation();
}

// Use event delegation so inputs don't cause full re-render
function setupGridDelegation() {
    const container = document.getElementById('grid-container');
    container.addEventListener('input', (e) => {
        const input = e.target;
        if (input.tagName !== 'INPUT') return;
        const idx = parseInt(input.dataset.rowIdx);
        const key = input.dataset.key;
        if (!isNaN(idx) && key) {
            rowData[idx][key] = input.value;
        }
    });
}

function refreshTableBody() {
    const container = document.getElementById('grid-container');
    const isMobile = window.innerWidth <= 768;

    if (!isMobile) {
        // Desktop: Table View — inputs use data attributes, NO oninput re-render
        container.innerHTML = `
            <table id="data-table">
                <thead>
                    <tr>
                        <th style="min-width:40px;">#</th>
                        ${variables.map(v => `<th style="min-width:160px;">${v.name}</th>`).join('')}
                        <th style="min-width:160px;">File Name</th>
                        <th style="min-width:48px;"></th>
                    </tr>
                </thead>
                <tbody>
                    ${rowData.map((row, idx) => `
                        <tr>
                            <td data-label="Row">${idx + 1}</td>
                            ${variables.map(v => `
                                <td data-label="${v.name}">
                                    <input type="text"
                                        data-row-idx="${idx}"
                                        data-key="${v.name}"
                                        value="${escHtml(row[v.name] || '')}"
                                        placeholder="${escHtml(v.name)}...">
                                </td>
                            `).join('')}
                            <td data-label="File Name">
                                <input type="text"
                                    data-row-idx="${idx}"
                                    data-key="_fn"
                                    value="${escHtml(row._fn || '')}"
                                    placeholder="output_name">
                            </td>
                            <td>
                                <button class="btn-ghost" onclick="delRow(${idx})"><i data-lucide="trash-2"></i></button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;
    } else {
        // Mobile: Card/Form View
        container.innerHTML = rowData.map((row, idx) => `
            <div class="entry-card">
                <div class="card-header">
                    <span>Entry #${idx + 1}</span>
                    <button class="btn-ghost" onclick="delRow(${idx})"><i data-lucide="trash-2"></i></button>
                </div>
                <div class="card-body">
                    ${variables.map(v => `
                        <div class="form-group">
                            <label>${escHtml(v.name)}</label>
                            <input type="text"
                                data-row-idx="${idx}"
                                data-key="${v.name}"
                                value="${escHtml(row[v.name] || '')}"
                                placeholder="Enter ${escHtml(v.name)}...">
                        </div>
                    `).join('')}
                    <div class="form-group">
                        <label>Output File Name</label>
                        <input type="text"
                            data-row-idx="${idx}"
                            data-key="_fn"
                            value="${escHtml(row._fn || '')}"
                            placeholder="e.g. Assignment_1">
                    </div>
                </div>
            </div>
        `).join('');
    }
    lucide.createIcons();
}

// HTML-escape helper to prevent attribute injection
function escHtml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function updateRow(idx, key, val) { rowData[idx][key] = val; }
function addRow() {
    // Sync any in-progress input values before adding row
    syncInputsToData();
    rowData.push({});
    refreshTableBody();
}
function delRow(idx) {
    syncInputsToData();
    if (rowData.length > 1) { rowData.splice(idx, 1); refreshTableBody(); }
}

// Sync all visible inputs back to rowData before re-render
function syncInputsToData() {
    const inputs = document.querySelectorAll('#grid-container input[data-row-idx]');
    inputs.forEach(input => {
        const idx = parseInt(input.dataset.rowIdx);
        const key = input.dataset.key;
        if (!isNaN(idx) && key) {
            rowData[idx][key] = input.value;
        }
    });
}

function proceedToData() {
    if (variables.length === 0) return alert("Select at least one text to edit first!");
    go(3);
}

// Phase 4: Generation Logic

/**
 * Robust XML text replacement for DOCX files.
 * Correctly handles variables split across multiple <w:t> tags in the XML.
 */
function robustXmlReplace(xmlContent, searchText, replacementText) {
    const xmlSearchText = searchText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const xmlReplacementText = replacementText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    let modified = false;
    let xml = xmlContent;
    let safety = 0;

    while (safety++ < 100) {
        const wtRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
        const segments = [];
        let match;
        while ((match = wtRegex.exec(xml)) !== null) {
            segments.push({
                xmlStart: match.index,
                xmlEnd: match.index + match[0].length,
                textContent: match[1]
            });
        }
        if (segments.length === 0) break;

        let plainText = '';
        const charMap = [];
        for (let si = 0; si < segments.length; si++) {
            for (let ci = 0; ci < segments[si].textContent.length; ci++) {
                charMap.push({ segIdx: si, offsetInSeg: ci });
                plainText += segments[si].textContent[ci];
            }
        }

        const foundIdx = plainText.indexOf(xmlSearchText);
        if (foundIdx === -1) break;

        modified = true;
        const matchEnd = foundIdx + xmlSearchText.length - 1;
        const startSeg = charMap[foundIdx].segIdx;
        const startOffset = charMap[foundIdx].offsetInSeg;
        const endSeg = charMap[matchEnd].segIdx;
        const endOffset = charMap[matchEnd].offsetInSeg;

        const ops = [];

        // Replacement Logic: 
        // 1. Put replacement into the first matching segment
        // 2. Clear out middle and end segments of the match
        // 3. Preserve trailing text in the end segment correctly

        if (startSeg === endSeg) {
            const seg = segments[startSeg];
            const textBefore = seg.textContent.substring(0, startOffset);
            const textAfter = seg.textContent.substring(endOffset + 1);
            ops.push({
                start: seg.xmlStart,
                end: seg.xmlEnd,
                replacement: `<w:t xml:space="preserve">${textBefore}${xmlReplacementText}${textAfter}</w:t>`
            });
        } else {
            // First segment of match
            const fSeg = segments[startSeg];
            const textBefore = fSeg.textContent.substring(0, startOffset);
            ops.push({
                start: fSeg.xmlStart,
                end: fSeg.xmlEnd,
                replacement: `<w:t xml:space="preserve">${textBefore}${xmlReplacementText}</w:t>`
            });

            // Middle segments (fully cleared)
            for (let si = startSeg + 1; si < endSeg; si++) {
                ops.push({ start: segments[si].xmlStart, end: segments[si].xmlEnd, replacement: '<w:t></w:t>' });
            }

            // End segment of match (preserve text after the match)
            const lSeg = segments[endSeg];
            const textAfter = lSeg.textContent.substring(endOffset + 1);
            ops.push({
                start: lSeg.xmlStart,
                end: lSeg.xmlEnd,
                replacement: `<w:t xml:space="preserve">${textAfter}</w:t>`
            });
        }

        // Apply operations in reverse order
        ops.sort((a, b) => b.start - a.start);
        for (const op of ops) {
            xml = xml.substring(0, op.start) + op.replacement + xml.substring(op.end);
        }

        // CLEANUP: Remove runs that became truly empty (no text content and no tabs)
        // This is critical to prevent ghost formatting runs from forcing line breaks
        xml = xml.replace(/<w:r[ >][\s\S]*?<\/w:r>/g, (match) => {
            const hasText = /<w:t[^>]*>[^<]+<\/w:t>/.test(match);
            const hasTab = match.includes('<w:tab/>');
            const hasDrawing = match.includes('<w:drawing') || match.includes('<v:shape') || match.includes('<w:pict');

            if (hasText || hasTab || hasDrawing) return match;
            return '';
        });
    }
    return { xml, modified };
}

async function runGeneration() {
    const btn = document.getElementById('generate-btn');
    const loader = document.getElementById('gen-loader');
    const text = document.getElementById('gen-text');

    if (variables.length === 0) return alert("Please select variables first.");

    const activeRows = rowData.filter(r => Object.keys(r).some(k => k !== '_fn' && r[k]));
    if (activeRows.length === 0) return alert("Please add at least one row of data.");

    btn.disabled = true;
    loader.style.display = 'block';
    text.textContent = 'Generating...';
    lastGeneratedDocs = []; // Reset docs array

    try {
        const zip = new JSZip();
        let templatePiz = new PizZip(fileBuffer);
        const files = Object.keys(templatePiz.files).filter(f => f.startsWith('word/') && f.endsWith('.xml'));
        const sortedVars = [...variables].sort((a, b) => b.originalText.length - a.originalText.length);

        let anyModified = false;
        files.forEach(fName => {
            let content = templatePiz.file(fName).asText();
            sortedVars.forEach(v => {
                const result = robustXmlReplace(content, v.originalText, `{${v.name}}`);
                if (result.modified) {
                    content = result.xml;
                    anyModified = true;
                }
            });
            templatePiz.file(fName, content);
        });

        const masterTemplateBuffer = templatePiz.generate({ type: 'arraybuffer' });

        for (let i = 0; i < activeRows.length; i++) {
            const rawData = activeRows[i];
            const data = {};
            Object.keys(rawData).forEach(key => {
                // Keep values as plain text — let Word handle word-wrap naturally
                // within table cells. Non-breaking spaces previously used here
                // caused long values (e.g. "Christopher Amsan") to overflow table
                // cell boundaries and visually shift adjacent column text.
                const val = typeof rawData[key] === 'string' ? rawData[key].trim() : (rawData[key] ?? '');
                data[key] = val;
            });

            const p = new PizZip(masterTemplateBuffer);
            // Do NOT post-process w:right / w:ind on paragraphs.
            // Removing those table-cell margins caused text ("Prof. Shalu") to
            // jump visually into the adjacent "Submitted by:" column.

            const DocxTemplaterRef = window.docxtemplater || window.Docxtemplater;
            const doc = new DocxTemplaterRef(p, { paragraphLoop: true, linebreaks: true });
            doc.render(data);

            const out = doc.getZip().generate({
                type: "uint8array",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            });

            const fileName = (rawData._fn || `Assignment_${i + 1}`).replace(/\.[^/.]+$/, "") + ".docx";
            zip.file(fileName, out);
            lastGeneratedDocs.push({ name: fileName, buffer: out });
        }

        const zipBlob = await zip.generateAsync({ type: "blob", mimeType: "application/zip" });
        lastGeneratedZipBlob = zipBlob;

        // Show share section on all devices
        const shareSection = document.getElementById('mobile-share-section');
        const shareBtn = document.getElementById('share-btn');
        lastGeneratedFile = new File([zipBlob], "Assignments.zip", { type: "application/zip" });
        shareSection.style.display = 'block';
        shareBtn.onclick = shareFiles;

        function triggerDownload() {
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = "Assignments_Bundle.zip";
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 500);
        }

        document.getElementById('final-download-btn').onclick = triggerDownload;
        go(4);

    } catch (err) {
        console.error("DocMorph Error:", err);
        alert("Generation stopped: " + err.message);
    } finally {
        btn.disabled = false;
        loader.style.display = 'none';
        text.innerHTML = '<i data-lucide="zap"></i> Generate ZIP';
        lucide.createIcons();
    }
}

async function shareFiles() {
    if (!lastGeneratedZipBlob) return;

    // Try native Web Share API first (works on mobile + some desktop browsers)
    if (navigator.share && lastGeneratedFile) {
        // Check if file sharing is supported
        const canShareFile = navigator.canShare && navigator.canShare({ files: [lastGeneratedFile] });

        if (canShareFile) {
            if (!window.isSecureContext && window.location.hostname !== 'localhost') {
                alert("🔒 File sharing requires HTTPS. Please download the ZIP manually.");
                return;
            }
            try {
                await navigator.share({
                    files: [lastGeneratedFile],
                    title: 'DocMorph Assignments',
                    text: 'Here are your customized documents!'
                });
                return; // Success — done
            } catch (err) {
                if (err.name === 'AbortError') return; // User cancelled
                // Fall through to download fallback
            }
        } else {
            // Try sharing just title+text (no file) — asks user to share link
            try {
                await navigator.share({
                    title: 'DocMorph Assignments',
                    text: 'Here are your customized DocMorph assignments!'
                });
                return;
            } catch (err) {
                if (err.name === 'AbortError') return;
            }
        }
    }

    // Fallback: just download the ZIP
    const url = URL.createObjectURL(lastGeneratedZipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Assignments_Bundle.zip';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
}

/**
 * Downloads all generated documents in a specific format.
 * DOCX: direct download. PDF: uses server-side conversion if available,
 * otherwise prints each doc via a hidden iframe (browser print -> Save as PDF).
 */
function downloadAllFormats(format) {
    if (lastGeneratedDocs.length === 0) return alert("Generate files first!");

    if (format === 'docx') {
        if (lastGeneratedDocs.length === 1) {
            const doc = lastGeneratedDocs[0];
            const blob = new Blob([doc.buffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = doc.name;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
        } else {
            // Download all as individual DOCX files
            lastGeneratedDocs.forEach((doc, i) => {
                setTimeout(() => {
                    const blob = new Blob([doc.buffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = doc.name;
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
                }, i * 400); // stagger downloads
            });
        }
        return;
    }

    if (format === 'pdf') {
        // Try server-side PDF conversion first
        downloadAsPdf();
    }
}

async function downloadAsPdf() {
    if (lastGeneratedDocs.length === 0) return;

    const btn = document.querySelector('button[onclick="downloadAllFormats(\'pdf\')"]');
    const origText = btn ? btn.innerHTML : '';
    if (btn) btn.innerHTML = '<div class="loader" style="display:block;width:16px;height:16px;"></div> Converting...';

    try {
        // Use server-side conversion
        const formData = new FormData();
        lastGeneratedDocs.forEach(doc => {
            const blob = new Blob([doc.buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            formData.append('files', blob, doc.name);
        });

        const response = await fetch('/convert-pdf', {
            method: 'POST',
            body: formData,
            signal: AbortSignal.timeout(30000)
        });

        if (!response.ok) throw new Error('Server conversion failed');

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = lastGeneratedDocs.length === 1
            ? lastGeneratedDocs[0].name.replace('.docx', '.pdf')
            : 'Assignments_PDF.zip';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);

    } catch (err) {
        console.warn('Server PDF conversion unavailable, using print fallback:', err.message);
        // Fallback: open each DOCX in an iframe and trigger browser print
        printDocsAsPdf();
    } finally {
        if (btn) btn.innerHTML = origText;
        lucide.createIcons();
    }
}

function printDocsAsPdf() {
    // Opens each doc as a blob URL, then triggers print dialog
    // User saves as PDF from the print dialog
    lastGeneratedDocs.forEach((doc, i) => {
        setTimeout(() => {
            const blob = new Blob([doc.buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            const url = URL.createObjectURL(blob);

            // Open in new tab for user to print/save as PDF
            const win = window.open(url, '_blank');
            if (!win) {
                // Popup blocked — fallback to download
                const a = document.createElement('a');
                a.href = url;
                a.download = doc.name.replace('.docx', '.pdf');
                document.body.appendChild(a);
                a.click();
                setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
            } else {
                // Try auto-printing after a brief delay for the tab to load
                setTimeout(() => {
                    try { win.print(); } catch (e) { }
                    setTimeout(() => URL.revokeObjectURL(url), 5000);
                }, 1000);
            }
        }, i * 800);
    });

    if (lastGeneratedDocs.length > 0) {
        showToast('📄 Opening docs for PDF — use Ctrl+P / File → Print → Save as PDF');
    }
}

function showToast(msg, duration = 5000) {
    let toast = document.getElementById('docmorph-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'docmorph-toast';
        toast.style.cssText = [
            'position:fixed', 'bottom:90px', 'left:50%', 'transform:translateX(-50%)',
            'background:#1f2937', 'border:1px solid var(--primary)', 'color:var(--text)',
            'padding:14px 24px', 'border-radius:20px', 'z-index:9999',
            'font-size:0.9rem', 'max-width:calc(100% - 40px)', 'text-align:center',
            'box-shadow:0 10px 40px rgba(0,0,0,0.6)', 'white-space:pre-wrap'
        ].join(';');
        document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.style.display = 'block';
    clearTimeout(toast._timeout);
    toast._timeout = setTimeout(() => { toast.style.display = 'none'; }, duration);
}
