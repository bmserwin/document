let fileBuffer = null;
let variables = []; // { id, name, originalText }
let rowData = [{}]; // Data for each student

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
docPreview.onmouseup = () => {
    const selection = window.getSelection();
    const text = selection.toString().trim();
    if (text.length > 0) {
        document.getElementById('selected-text-preview').textContent = `"${text}"`;
        showModal(text);
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
    const head = document.getElementById('table-head');
    const body = document.getElementById('table-body');

    let headHtml = '<th>#</th>';
    variables.forEach(v => { headHtml += `<th>${v.name}</th>`; });
    headHtml += '<th>File Name</th><th></th>';
    head.innerHTML = headHtml;

    refreshTableBody();
}

function refreshTableBody() {
    const body = document.getElementById('table-body');
    body.innerHTML = rowData.map((row, idx) => `
        <tr>
            <td>${idx + 1}</td>
            ${variables.map(v => `
                <td><input type="text" value="${row[v.name] || ''}" oninput="updateRow(${idx}, '${v.name}', this.value)" placeholder="${v.name}..."></td>
            `).join('')}
            <td><input type="text" value="${row._fn || ''}" oninput="updateRow(${idx}, '_fn', this.value)" placeholder="output_name"></td>
            <td><button class="btn-ghost" onclick="delRow(${idx})"><i data-lucide="x"></i></button></td>
        </tr>
    `).join('');
    lucide.createIcons();
}

function updateRow(idx, key, val) { rowData[idx][key] = val; }
function addRow() { rowData.push({}); refreshTableBody(); }
function delRow(idx) { if (rowData.length > 1) { rowData.splice(idx, 1); refreshTableBody(); } }

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
                let val = typeof rawData[key] === 'string' ? rawData[key].trim() : rawData[key];

                if (typeof val === 'string') {
                    // CRITICAL: Replace ALL potential wrap points with non-breaking versions
                    val = val.replace(/ /g, '\u00A0')  // Non-breaking space
                        .replace(/-/g, '\u2011') // Non-breaking hyphen
                        .replace(/:/g, ':\u00A0'); // Colon followed by non-breaking space
                }
                data[key] = val;
            });

            const p = new PizZip(masterTemplateBuffer);

            // POST-PROCESSING: Remove right indents from paragraphs containing our variables
            // This prevents the "right margin" from forcing a wrap too early
            let docXml = p.file('word/document.xml').asText();
            docXml = docXml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (para) => {
                const hasVariable = Object.keys(data).some(k => para.includes(`{${k}}`));
                if (hasVariable) {
                    // Strip any right indent that might be squeezing the text
                    return para.replace(/<w:ind[^>]*w:right="[^"]*"[^>]*\/>/g, '')
                        .replace(/w:right="[^"]*"/g, 'w:right="0"');
                }
                return para;
            });
            p.file('word/document.xml', docXml);

            const DocxTemplaterRef = window.docxtemplater || window.Docxtemplater;
            const doc = new DocxTemplaterRef(p, { paragraphLoop: true, linebreaks: true });
            doc.render(data);

            const out = doc.getZip().generate({
                type: "uint8array",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            });

            const fileName = (rawData._fn || `Assignment_${i + 1}`).replace(/\.[^/.]+$/, "") + ".docx";
            zip.file(fileName, out);
        }

        const zipBlob = await zip.generateAsync({ type: "blob", mimeType: "application/zip" });

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
