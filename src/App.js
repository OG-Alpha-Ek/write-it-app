import React, { useState, useRef, useEffect } from 'react';
import { 
  UploadCloud, 
  FileText, 
  Download, 
  Copy, 
  CheckCircle2, 
  RefreshCcw, 
  AlertCircle,
  FileImage,
  File as FileIcon,
  Table as TableIcon,
  FileSpreadsheet,
  Eye,
  Edit3,
  Loader2
} from 'lucide-react';

const App = () => {
  const [files, setFiles] = useState([]);
  const [extractedText, setExtractedText] = useState('');
  const [status, setStatus] = useState('idle'); // idle, processing, success, error
  const [errorMessage, setErrorMessage] = useState('');
  const [dragActive, setDragActive] = useState(false);
  const [copied, setCopied] = useState(false);
  const [viewMode, setViewMode] = useState('edit'); // 'edit' or 'preview'
  const [downloading, setDownloading] = useState(false);
  const [previewHtml, setPreviewHtml] = useState('');
  
  const fileInputRef = useRef(null);

  // Load marked dynamically when needed for preview
  useEffect(() => {
    if (viewMode === 'preview' && extractedText) {
      import('marked').then((module) => {
        setPreviewHtml(module.marked(extractedText));
      }).catch(err => {
        console.error("Failed to load marked:", err);
        setPreviewHtml('<p>Error loading preview. Please check console.</p>');
      });
    }
  }, [viewMode, extractedText]);

  // Helper to fetch with exponential backoff
  const fetchWithRetry = async (url, options, retries = 5) => {
    const delays = [1000, 2000, 4000, 8000, 16000];
    for (let i = 0; i < retries; i++) {
      try {
        const response = await fetch(url, options);
        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
        return await response.json();
      } catch (err) {
        if (i === retries - 1) throw err;
        await new Promise(resolve => setTimeout(resolve, delays[i]));
      }
    }
  };

  const handleFiles = async (selectedFiles) => {
    const validFiles = Array.from(selectedFiles).filter(f => 
      f.type.startsWith('image/') || f.type === 'application/pdf'
    );

    if (validFiles.length === 0) {
      setErrorMessage('Please select valid images or PDF files.');
      setStatus('error');
      return;
    }

    // Map files to local state objects
    const newFiles = validFiles.map(file => ({
      file,
      id: Math.random().toString(36).substring(7),
      name: file.name,
      type: file.type,
      status: 'pending', // pending, processing, done, error
      preview: file.type.startsWith('image/') ? URL.createObjectURL(file) : null
    }));

    setFiles(newFiles);
    setStatus('processing');
    setErrorMessage('');
    
    let combinedText = extractedText;

    // Process files sequentially
    for (let i = 0; i < newFiles.length; i++) {
      const currentFile = newFiles[i];
      
      // Update UI to show which file is processing
      setFiles(prev => prev.map(f => f.id === currentFile.id ? { ...f, status: 'processing' } : f));

      try {
        const base64Data = await readFileAsBase64(currentFile.file);
        const mimeType = currentFile.type;
        const text = await analyzeDocument(base64Data, mimeType);
        
        combinedText = combinedText ? `${combinedText}\n\n---\n\n${text}` : text;
        setExtractedText(combinedText);
        
        setFiles(prev => prev.map(f => f.id === currentFile.id ? { ...f, status: 'done' } : f));
      } catch (error) {
        console.error(`Error processing ${currentFile.name}:`, error);
        setFiles(prev => prev.map(f => f.id === currentFile.id ? { ...f, status: 'error' } : f));
        // Continue with other files even if one fails
      }
    }

    setStatus('success');
  };

  const readFileAsBase64 = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result.split(',')[1]);
      reader.onerror = (e) => reject(e);
      reader.readAsDataURL(file);
    });
  };

  const analyzeDocument = async (base64Data, mimeType) => {
    // Pulls API key from the local .env file instead of hardcoding it
    const apiKey = process.env.REACT_APP_GEMINI_API_KEY || ""; 
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`;
    
    const prompt = `You are a professional document digitization assistant. 
    Extract all text, tables, and data from the provided document perfectly. 
    - Preserve paragraphs, lists, and general formatting using standard Markdown.
    - If you encounter ANY tables, output them perfectly formatted as standard Markdown tables.
    - If you encounter any GRAPHS, CHARTS, or PLOTS, carefully extract the underlying data and represent it as a Markdown table.
    - Output ONLY the extracted content. Do not include conversational filler like "Here is the text".`;

    const payload = {
      contents: [{
        role: "user",
        parts: [
          { text: prompt },
          { inlineData: { mimeType: mimeType, data: base64Data } }
        ]
      }]
    };

    const data = await fetchWithRetry(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) throw new Error("No text returned from the model.");
    return text.trim();
  };

  // Drag and Drop
  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") setDragActive(true);
    else if (e.type === "dragleave") setDragActive(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  };

  const handleChange = (e) => {
    e.preventDefault();
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(e.target.files);
    }
  };

  const handleCopy = () => {
    const textArea = document.createElement("textarea");
    textArea.value = extractedText;
    document.body.appendChild(textArea);
    textArea.select();
    try {
      document.execCommand('copy');
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {}
    document.body.removeChild(textArea);
  };

  // ---------------- EXPORT LOGIC ---------------- //

  // 1. Export Word (With Native Tables)
  const handleDownloadWord = async () => {
    setDownloading(true);
    try {
      const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle } = await import('docx');
      
      const lines = extractedText.split('\n');
      const children = [];
      let tableBuffer = [];

      const flushTable = () => {
        if (tableBuffer.length > 0) {
          const tableRows = tableBuffer.map((row, rowIndex) => {
            return new TableRow({
              children: row.map(cellText => new TableCell({
                children: [new Paragraph({ 
                  children: [new TextRun({ text: cellText, font: "Calibri", size: 22, bold: rowIndex === 0 })],
                  spacing: { before: 100, after: 100 }
                })],
                width: { size: 100 / row.length, type: WidthType.PERCENTAGE },
                margins: { top: 100, bottom: 100, left: 100, right: 100 }
              }))
            });
          });

          children.push(new Table({
            rows: tableRows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            }
          }));
          children.push(new Paragraph({ text: "" })); // spacing after table
          tableBuffer = [];
        }
      };

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        // Super basic markdown table detection
        if (line.startsWith('|') && line.endsWith('|')) {
          const cells = line.split('|').slice(1, -1).map(c => c.trim());
          if (!cells.every(c => /^[-: ]+$/.test(c))) { // Skip markdown separator row
            tableBuffer.push(cells);
          }
        } else {
          flushTable();
          if (line) {
            let textRunOptions = { text: line, font: "Calibri", size: 24 };
            // Basic bold detection for headings
            if (line.startsWith('#')) {
              textRunOptions = { text: line.replace(/^#+\s*/, ''), font: "Calibri", size: 32, bold: true };
            }
            children.push(new Paragraph({ children: [new TextRun(textRunOptions)], spacing: { after: 200 } }));
          }
        }
      }
      flushTable(); // Flush if ends with table

      const doc = new Document({ sections: [{ properties: {}, children }] });
      const blob = await Packer.toBlob(doc);
      triggerDownload(blob, `write_It_${new Date().getTime()}.docx`);
    } catch (error) {
      console.error("Word export error:", error);
      alert("Failed to export Word document.");
    }
    setDownloading(false);
  };

  // 2. Export Excel (Extracts tables to sheets)
  const handleDownloadExcel = async () => {
    setDownloading(true);
    try {
      const XLSX = await import('xlsx');
      const lines = extractedText.split('\n');
      const tables = [];
      let currentTable = null;

      // Extract tables
      for (const line of lines) {
        if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
          if (!currentTable) currentTable = [];
          const cells = line.split('|').slice(1, -1).map(c => c.trim());
          if (!cells.every(c => /^[-: ]+$/.test(c))) {
            currentTable.push(cells);
          }
        } else {
          if (currentTable) {
            tables.push(currentTable);
            currentTable = null;
          }
        }
      }
      if (currentTable) tables.push(currentTable);

      const wb = XLSX.utils.book_new();
      
      if (tables.length > 0) {
        tables.forEach((tableData, idx) => {
          const ws = XLSX.utils.aoa_to_sheet(tableData);
          XLSX.utils.book_append_sheet(wb, ws, `Table_${idx + 1}`);
        });
      } else {
        // If no tables, just put the text in cell A1
        const ws = XLSX.utils.aoa_to_sheet([[extractedText]]);
        XLSX.utils.book_append_sheet(wb, ws, "Document Text");
      }

      XLSX.writeFile(wb, `write_It_Data_${new Date().getTime()}.xlsx`);
    } catch (error) {
      console.error("Excel export error:", error);
      alert("Failed to export Excel document.");
    }
    setDownloading(false);
  };

  // 3. Export PDF (Uses html2pdf to render markdown)
  const handleDownloadPDF = async () => {
    setDownloading(true);
    try {
      const module = await import('marked');
      const htmlContent = module.marked(extractedText);
      
      // Load html2pdf dynamically
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
      document.body.appendChild(script);

      script.onload = () => {
        const element = document.createElement('div');
        element.innerHTML = htmlContent;
        
        // Add basic styling for PDF
        element.style.padding = '20px';
        element.style.fontFamily = 'Helvetica, Arial, sans-serif';
        element.style.color = '#000';
        element.style.backgroundColor = '#fff';
        
        // Style tables for PDF
        const tables = element.querySelectorAll('table');
        tables.forEach(table => {
          table.style.width = '100%';
          table.style.borderCollapse = 'collapse';
          table.style.marginBottom = '20px';
          const cells = table.querySelectorAll('th, td');
          cells.forEach(cell => {
            cell.style.border = '1px solid #ccc';
            cell.style.padding = '8px';
          });
        });

        const opt = {
          margin:       1,
          filename:     `write_It_${new Date().getTime()}.pdf`,
          image:        { type: 'jpeg', quality: 0.98 },
          html2canvas:  { scale: 2 },
          jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
        };

        window.html2pdf().set(opt).from(element).save().then(() => {
          setDownloading(false);
        });
      };
    } catch (error) {
      console.error("PDF export error:", error);
      alert("Failed to export PDF.");
      setDownloading(false);
    }
  };

  const triggerDownload = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const resetApp = () => {
    setFiles([]);
    setExtractedText('');
    setStatus('idle');
    setErrorMessage('');
  };

  return (
    <div className="min-h-screen bg-[#121212] text-gray-100 font-sans selection:bg-[#1DB954] selection:text-white flex flex-col">
      {/* Header */}
      <header className="border-b border-[#282828] bg-[#000000] sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-8 h-8 rounded-full bg-[#1DB954] flex items-center justify-center">
              <Edit3 className="w-4 h-4 text-black" />
            </div>
            <span className="text-2xl font-bold tracking-tight text-white">write It</span>
          </div>
          {status !== 'idle' && (
            <button 
              onClick={resetApp}
              className="flex items-center space-x-2 text-sm font-medium text-gray-400 hover:text-white transition-colors"
            >
              <RefreshCcw className="w-4 h-4" />
              <span>Start Over</span>
            </button>
          )}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex-1 w-full flex flex-col">
        
        {/* Upload Area (Idle) */}
        {status === 'idle' && (
          <div className="flex-1 flex flex-col items-center justify-center">
            <div className="max-w-3xl w-full text-center space-y-8">
              <div>
                <h1 className="text-4xl sm:text-6xl font-extrabold tracking-tight mb-6 text-white">
                  Turn Pixels into <span className="text-[#1DB954]">Perfect Documents.</span>
                </h1>
                <p className="text-lg text-gray-400 max-w-2xl mx-auto">
                  Upload images, notes, or PDFs. We extract text, format tables, and pull data from graphs—ready to export to Word, Excel, or PDF.
                </p>
              </div>

              <div 
                className={`relative border-2 border-dashed rounded-2xl p-14 transition-all duration-300 ease-in-out cursor-pointer group
                  ${dragActive ? 'border-[#1DB954] bg-[#1DB954]/10 scale-[1.02]' : 'border-[#282828] bg-[#181818] hover:border-[#1DB954]/50 hover:bg-[#282828]'}
                `}
                onDragEnter={handleDrag}
                onDragLeave={handleDrag}
                onDragOver={handleDrag}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/*, application/pdf"
                  multiple
                  className="hidden"
                  onChange={handleChange}
                />
                
                <div className="flex flex-col items-center justify-center space-y-6">
                  <div className="w-20 h-20 rounded-full bg-[#282828] flex items-center justify-center group-hover:scale-110 transition-transform duration-300">
                    <UploadCloud className={`w-10 h-10 ${dragActive ? 'text-[#1DB954]' : 'text-gray-400'}`} />
                  </div>
                  <div className="text-center">
                    <p className="text-2xl font-bold text-white mb-2">
                      Drop files here
                    </p>
                    <p className="text-gray-400">
                      Supports Images (PNG, JPG) and PDFs. Select multiple files at once.
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Processing & Results View */}
        {status !== 'idle' && (
          <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-6 min-h-0">
            
            {/* Left Column: File Queue */}
            <div className="lg:col-span-3 bg-[#181818] border border-[#282828] rounded-xl p-4 flex flex-col h-[300px] lg:h-full overflow-hidden">
              <h3 className="text-sm font-bold text-gray-400 uppercase tracking-wider mb-4 flex items-center space-x-2">
                <FileIcon className="w-4 h-4" />
                <span>Queue ({files.length})</span>
              </h3>
              
              <div className="flex-1 overflow-y-auto space-y-2 pr-2 custom-scrollbar">
                {files.map((f, i) => (
                  <div key={f.id} className="p-3 bg-[#282828] rounded-lg flex items-center space-x-3">
                    {f.preview ? (
                      <img src={f.preview} alt="preview" className="w-10 h-10 rounded object-cover" />
                    ) : (
                      <div className="w-10 h-10 rounded bg-[#121212] flex items-center justify-center">
                        <FileText className="w-5 h-5 text-[#1DB954]" />
                      </div>
                    )}
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-white truncate">{f.name}</p>
                      <div className="flex items-center mt-1">
                        {f.status === 'pending' && <span className="text-xs text-gray-500">Waiting...</span>}
                        {f.status === 'processing' && <span className="text-xs text-[#1DB954] flex items-center"><Loader2 className="w-3 h-3 mr-1 animate-spin" /> Processing</span>}
                        {f.status === 'done' && <span className="text-xs text-gray-400 flex items-center"><CheckCircle2 className="w-3 h-3 mr-1" /> Done</span>}
                        {f.status === 'error' && <span className="text-xs text-red-400">Error</span>}
                      </div>
                    </div>
                  </div>
                ))}
                
                {status === 'success' && (
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    className="w-full mt-4 p-3 border-2 border-dashed border-[#282828] rounded-lg text-sm text-gray-400 hover:text-white hover:border-[#1DB954] transition-colors flex items-center justify-center space-x-2"
                  >
                    <UploadCloud className="w-4 h-4" />
                    <span>Add more files</span>
                  </button>
                )}
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/*, application/pdf"
                  multiple
                  className="hidden"
                  onChange={handleChange}
                />
              </div>
            </div>

            {/* Right Column: Editor & Actions */}
            <div className="lg:col-span-9 bg-[#181818] border border-[#282828] rounded-xl flex flex-col h-[600px] lg:h-full overflow-hidden">
              
              {/* Toolbar */}
              <div className="p-4 border-b border-[#282828] flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-[#121212]/50">
                
                {/* View Toggles */}
                <div className="flex bg-[#282828] p-1 rounded-lg w-max">
                  <button
                    onClick={() => setViewMode('edit')}
                    className={`flex items-center space-x-2 px-4 py-2 rounded-md text-sm font-medium transition-all ${viewMode === 'edit' ? 'bg-[#1DB954] text-black shadow-sm' : 'text-gray-400 hover:text-white'}`}
                  >
                    <Edit3 className="w-4 h-4" />
                    <span>Edit Source</span>
                  </button>
                  <button
                    onClick={() => setViewMode('preview')}
                    className={`flex items-center space-x-2 px-4 py-2 rounded-md text-sm font-medium transition-all ${viewMode === 'preview' ? 'bg-[#1DB954] text-black shadow-sm' : 'text-gray-400 hover:text-white'}`}
                  >
                    <Eye className="w-4 h-4" />
                    <span>Preview Tables</span>
                  </button>
                </div>

                {/* Export Actions */}
                {status === 'success' && (
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={handleCopy}
                      className="p-2.5 rounded-lg bg-[#282828] hover:bg-[#3E3E3E] text-gray-300 transition-colors group"
                      title="Copy to clipboard"
                    >
                      {copied ? <CheckCircle2 className="w-4 h-4 text-[#1DB954]" /> : <Copy className="w-4 h-4 group-hover:scale-110 transition-transform" />}
                    </button>
                    
                    <button 
                      onClick={handleDownloadWord}
                      disabled={downloading}
                      className="px-3 py-2.5 rounded-lg bg-[#2B579A] hover:bg-[#3668b5] text-white text-sm font-medium transition-all flex items-center space-x-2 disabled:opacity-50"
                      title="Download as Word Doc"
                    >
                      <FileText className="w-4 h-4" />
                      <span className="hidden sm:inline">Word</span>
                    </button>

                    <button 
                      onClick={handleDownloadExcel}
                      disabled={downloading}
                      className="px-3 py-2.5 rounded-lg bg-[#217346] hover:bg-[#2c8d58] text-white text-sm font-medium transition-all flex items-center space-x-2 disabled:opacity-50"
                      title="Extract tables to Excel"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      <span className="hidden sm:inline">Excel</span>
                    </button>

                    <button 
                      onClick={handleDownloadPDF}
                      disabled={downloading}
                      className="px-3 py-2.5 rounded-lg bg-[#D24726] hover:bg-[#e85836] text-white text-sm font-medium transition-all flex items-center space-x-2 disabled:opacity-50"
                      title="Download as PDF"
                    >
                      <Download className="w-4 h-4" />
                      <span className="hidden sm:inline">PDF</span>
                    </button>
                  </div>
                )}
              </div>

              {/* Editor/Preview Area */}
              <div className="flex-1 relative bg-[#181818] overflow-hidden flex flex-col">
                {status === 'processing' && !extractedText ? (
                  <div className="absolute inset-0 flex flex-col items-center justify-center p-6 bg-[#181818] z-10">
                    <Loader2 className="w-10 h-10 text-[#1DB954] animate-spin mb-4" />
                    <p className="text-gray-400 font-medium">Extracting documents perfectly...</p>
                  </div>
                ) : status === 'error' && !extractedText ? (
                  <div className="absolute inset-0 flex flex-col items-center justify-center p-6 text-center">
                    <AlertCircle className="w-12 h-12 text-red-500 mb-4" />
                    <p className="text-red-400 mb-4">{errorMessage}</p>
                  </div>
                ) : (
                  <>
                    {/* Raw Editor */}
                    <textarea
                      value={extractedText}
                      onChange={(e) => setExtractedText(e.target.value)}
                      className={`w-full h-full p-6 bg-transparent text-gray-200 font-mono text-sm resize-none outline-none focus:ring-0 leading-relaxed custom-scrollbar ${viewMode === 'edit' ? 'block' : 'hidden'}`}
                      placeholder="Extracted text will appear here. Markdown tables are supported."
                    />
                    
                    {/* Rich Preview */}
                    <div 
                      className={`w-full h-full p-8 overflow-y-auto bg-white text-black prose prose-sm max-w-none custom-scrollbar ${viewMode === 'preview' ? 'block' : 'hidden'}`}
                      dangerouslySetInnerHTML={{ __html: previewHtml || '<p>*No text extracted yet.*</p>' }}
                    >
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="py-6 border-t border-[#282828] bg-[#000000]">
        <p className="text-center text-sm font-medium text-gray-500 tracking-wide">
          by <span className="text-[#1DB954]">Ephraim kwaedza</span>
        </p>
      </footer>

      {/* CSS for custom scrollbars and basic prose styling (since Tailwind prose isn't imported) */}
      <style dangerouslySetInnerHTML={{__html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #3E3E3E;
          border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #1DB954;
        }
        
        /* Basic styling for the markdown preview */
        .prose table { width: 100%; border-collapse: collapse; margin-bottom: 1em; }
        .prose th, .prose td { border: 1px solid #d1d5db; padding: 0.5rem; text-align: left; }
        .prose th { background-color: #f3f4f6; font-weight: 600; }
        .prose h1, .prose h2, .prose h3 { margin-top: 1.5em; margin-bottom: 0.5em; font-weight: bold; }
        .prose p { margin-bottom: 1em; }
      `}} />
    </div>
  );
};

export default App;