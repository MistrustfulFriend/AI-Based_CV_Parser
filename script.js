let pdfFile = null;

// PDF.js worker configuration
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

document.getElementById('pdf-upload').addEventListener('change', function(e) {
    pdfFile = e.target.files[0];
    if (pdfFile) {
        document.getElementById('file-name').textContent = `Selected: ${pdfFile.name}`;
        checkFormReady();
    }
});

document.getElementById('api-key').addEventListener('input', checkFormReady);

function checkFormReady() {
    const apiKey = document.getElementById('api-key').value.trim();
    document.getElementById('process-btn').disabled = !(pdfFile && apiKey);
}

document.getElementById('process-btn').addEventListener('click', processResume);

async function extractTextFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
    let fullText = '';
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ');
        fullText += pageText + '\n';
    }
    
    return fullText;
}

async function processResume() {
    const statusDiv = document.getElementById('status');
    const previewDiv = document.getElementById('preview');
    const apiKey = document.getElementById('api-key').value.trim();
    
    statusDiv.className = 'loading';
    statusDiv.textContent = 'Extracting text from PDF...';
    previewDiv.classList.remove('show');
    
    try {
        const pdfText = await extractTextFromPDF(pdfFile);
        
        statusDiv.textContent = 'Analyzing resume with AI...';
        
        const response = await fetch('/api/parse', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                pdf_text: pdfText,
                api_key: apiKey
            })
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || `Server error: ${response.statusText}`);
        }
        
        const result = await response.json();
        
        if (result.error) {
            throw new Error(result.error);
        }
        
        statusDiv.className = 'success';
        statusDiv.textContent = 'Resume processed successfully! Generating Word document...';
        
        // Download the Word document
        const downloadResponse = await fetch('/api/download', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(result.data)
        });
        
        if (!downloadResponse.ok) {
            throw new Error('Failed to generate Word document');
        }
        
        const blob = await downloadResponse.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${result.data.expert.last_name}_${result.data.expert.first_name}_Profile.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        statusDiv.textContent = 'Document downloaded successfully!';
        
        // Show preview
        previewDiv.innerHTML = `<h3>Parsed Data Preview:</h3><pre>${JSON.stringify(result.data, null, 2)}</pre>`;
        previewDiv.classList.add('show');
        
    } catch (error) {
        statusDiv.className = 'error';
        statusDiv.textContent = `Error: ${error.message}`;
        console.error('Full error:', error);
    }
}