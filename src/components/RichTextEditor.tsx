import React, { useRef, useState, useEffect } from 'react';
import mammoth from 'mammoth';
import { AlignLeft, AlignCenter, AlignRight, FileText, Upload } from 'lucide-react';

interface RichTextEditorProps {
  value: string;
  onChange: (val: string) => void;
  readOnly?: boolean;
}

export function RichTextEditor({ value, onChange, readOnly }: RichTextEditorProps) {
  const editorRef = useRef<HTMLDivElement>(null);
  const [fileName, setFileName] = useState<string>('');
  const isUpdating = useRef(false);

  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value && !isUpdating.current) {
      editorRef.current.innerHTML = value;
    }
  }, [value]);

  const handleInput = () => {
    if (editorRef.current) {
      isUpdating.current = true;
      onChange(editorRef.current.innerHTML);
      // reset the updating flag in the next tick
      setTimeout(() => {
        isUpdating.current = false;
      }, 0);
    }
  };

  const formatText = (command: string, value: string | undefined = undefined) => {
    document.execCommand(command, false, value);
    editorRef.current?.focus();
    handleInput();
  };

  const applyFontName = (fontName: string) => {
    const tempFontName = `__TEMP_FONT_${Math.random()}__`;
    document.execCommand('fontName', false, tempFontName);
    if (editorRef.current) {
      const fonts = editorRef.current.querySelectorAll(`font[face="${tempFontName}"]`);
      fonts.forEach(font => {
        const span = document.createElement('span');
        span.style.fontFamily = fontName;
        span.innerHTML = font.innerHTML;
        font.replaceWith(span);
      });
    }
    editorRef.current?.focus();
    handleInput();
  };

  const applyColor = (color: string) => {
    const tempFontName = `__TEMP_FONT_${Math.random()}__`;
    document.execCommand('fontName', false, tempFontName);
    if (editorRef.current) {
      const fonts = editorRef.current.querySelectorAll(`font[face="${tempFontName}"]`);
      fonts.forEach(font => {
        const span = document.createElement('span');
        span.style.color = color;
        span.innerHTML = font.innerHTML;
        font.replaceWith(span);
      });
    }
    editorRef.current?.focus();
    handleInput();
  };

  const applyFontSize = (size: string) => {
    const tempFontName = `__TEMP_FONT_${Math.random()}__`;
    document.execCommand('fontName', false, tempFontName);
    if (editorRef.current) {
      const fonts = editorRef.current.querySelectorAll(`font[face="${tempFontName}"]`);
      fonts.forEach(font => {
        const span = document.createElement('span');
        span.style.fontSize = size;
        span.innerHTML = font.innerHTML;
        font.replaceWith(span);
      });
    }
    editorRef.current?.focus();
    handleInput();
  };

  const applyLineHeight = (height: string) => {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;
    
    let currentNode = selection.anchorNode;
    while (currentNode && currentNode !== editorRef.current) {
      if (currentNode instanceof HTMLElement && (currentNode.tagName === 'P' || currentNode.tagName === 'DIV' || currentNode.tagName === 'LI')) {
        currentNode.style.lineHeight = height;
        handleInput();
        return;
      }
      currentNode = currentNode.parentNode;
    }
    
    // If no block found, just set it globally or wrap
    document.execCommand('insertHTML', false, `<div style="line-height: ${height}">${selection.toString()}</div>`);
    handleInput();
  };

  const fonts = [
    'Arial', 'Times New Roman', 'Helvetica', 'Tahoma', 'Verdana', 'Georgia'
  ];

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = async (e) => {
      const arrayBuffer = e.target?.result as ArrayBuffer;
      try {
        const options = {
          styleMap: [
            "table => table.custom-word-table:fresh",
            "p[style-name='Heading 1'] => h1:fresh",
            "p[style-name='Heading 2'] => h2:fresh",
            "p[style-name='Heading 3'] => h3:fresh",
          ]
        };
        const result = await mammoth.convertToHtml({ arrayBuffer }, options);
        if (editorRef.current) {
          let html = result.value;
          
          const parser = new DOMParser();
          const doc = parser.parseFromString(html, 'text/html');

          // Remove empty rows from tables
          doc.querySelectorAll('tr').forEach(tr => {
            const textContent = tr.textContent?.trim().replace(/\u00a0/g, '') || '';
            if (!textContent) {
              tr.remove();
            }
          });

          // Style elements
          doc.querySelectorAll('table').forEach(table => {
            table.className = "w-full border-collapse border border-black my-2";
          });
          doc.querySelectorAll('th').forEach(th => {
            const td = doc.createElement('td');
            td.innerHTML = th.innerHTML;
            td.className = "border border-black p-0";
            th.replaceWith(td);
          });
          doc.querySelectorAll('td').forEach(td => {
            td.className = "border border-black p-0";
          });

          html = doc.body.innerHTML;
          
          editorRef.current.innerHTML = html;
          handleInput();
        }
      } catch (err) {
        console.error('Error parsing Word document', err);
        alert('Không thể đọc file Word. Vui lòng kiểm tra lại.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="w-full border border-gray-300 rounded-md overflow-hidden bg-white shadow-sm focus-within:ring-2 focus-within:ring-blue-500 focus-within:border-blue-500 transition-all">
      {/* Toolbar */}
      {!readOnly && (
        <div className="flex items-center gap-2 p-2 bg-gray-50 border-b border-gray-300 flex-wrap">
          <button
            type="button"
            onClick={() => formatText('bold')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors font-bold w-8 h-8 flex items-center justify-center"
            title="In đậm"
          >
            B
          </button>
          <button
            type="button"
            onClick={() => formatText('italic')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors italic w-8 h-8 flex items-center justify-center"
            title="In nghiêng"
          >
            I
          </button>
          <button
            type="button"
            onClick={() => formatText('underline')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors underline w-8 h-8 flex items-center justify-center"
            title="Gạch chân"
          >
            U
          </button>
          <div className="w-px h-6 bg-gray-300 mx-1"></div>
          <button
            type="button"
            onClick={() => formatText('justifyLeft')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors"
            title="Căn trái"
          >
            <AlignLeft size={18} />
          </button>
          <button
            type="button"
            onClick={() => formatText('justifyCenter')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors"
            title="Căn giữa"
          >
            <AlignCenter size={18} />
          </button>
          <button
            type="button"
            onClick={() => formatText('justifyRight')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors"
            title="Căn phải"
          >
            <AlignRight size={18} />
          </button>
          <button
            type="button"
            onClick={() => formatText('justifyFull')}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition-colors w-8 h-8 flex items-center justify-center font-serif"
            title="Căn đều 2 bên"
          >
            ≡
          </button>
          
          <div className="w-px h-6 bg-gray-300 mx-1"></div>
          
          <input 
            type="color" 
            onChange={(e) => applyColor(e.target.value)} 
            className="w-8 h-8 p-0 border-0 rounded cursor-pointer"
            title="Màu chữ" 
          />
          
          <select 
            onChange={(e) => applyFontName(e.target.value)}
            className="p-1.5 text-sm border border-gray-300 rounded bg-white"
            title="Font chữ"
          >
            <option value="">Font chữ...</option>
            {fonts.map(font => <option key={font} value={font}>{font}</option>)}
          </select>
          
          <select 
            onChange={(e) => applyFontSize(e.target.value)}
            className="p-1.5 text-sm border border-gray-300 rounded bg-white"
            title="Cỡ chữ"
          >
            <option value="">Cỡ chữ...</option>
            <option value="10pt">10pt</option>
            <option value="11pt">11pt</option>
            <option value="12pt">12pt</option>
            <option value="13pt">13pt</option>
            <option value="14pt">14pt</option>
            <option value="16pt">16pt</option>
            <option value="18pt">18pt</option>
            <option value="20pt">20pt</option>
            <option value="24pt">24pt</option>
          </select>

          <select 
            onChange={(e) => applyLineHeight(e.target.value)}
            className="p-1.5 text-sm border border-gray-300 rounded bg-white"
            title="Giãn cách dòng"
          >
            <option value="">Giãn dòng...</option>
            <option value="1.0">1.0</option>
            <option value="1.15">1.15</option>
            <option value="1.5">1.5</option>
            <option value="2.0">2.0</option>
          </select>
          
          <div className="w-px h-6 bg-gray-300 mx-1"></div>
          
          <div className="relative flex items-center group">
            <input
              type="file"
              accept=".docx"
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              title="Nhập từ file Word"
            />
            <button
              type="button"
              className="flex items-center gap-1 px-3 py-1.5 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors"
            >
              <Upload size={16} />
              <span>Nhập từ Word</span>
            </button>
          </div>
          {fileName && <span className="text-xs text-gray-500 ml-2 truncate max-w-[200px] flex items-center gap-1"><FileText size={12}/> {fileName}</span>}
          
          <div className="w-full text-sm text-amber-800 bg-amber-50 p-2 mt-2 rounded border border-amber-200">
            <strong>* Lưu ý:</strong> Nút "Nhập từ Word" sẽ lấy định dạng mặc định (có thể mất màu sắc). Để giữ nguyên định dạng, <strong>VUI LÒNG MỞ FILE WORD, COPY VÀ DÁN TRỰC TIẾP (Ctrl+V)</strong> vào khung soạn thảo.
          </div>
        </div>
      )}

      {/* Editor Content */}
      <div className="bg-white overflow-y-auto overflow-x-auto p-4 flex justify-center h-[400px] min-h-[200px] resize-y relative will-change-transform border-t border-white">
        <div
          ref={editorRef}
          contentEditable={!readOnly}
          onInput={handleInput}
          onBlur={handleInput}
          onPaste={(e) => {
            // Delay 1 tick to let the browser paste HTML/Images first
            setTimeout(() => {
              if (editorRef.current) {
                let changed = false;
                editorRef.current.querySelectorAll('img').forEach(img => {
                  // If it's a local file path, it will appear as broken
                  if (img.src.startsWith('file://') || img.src.startsWith('webkit-fake-url://') || img.src === '' || img.src.indexOf('http') !== 0 && img.src.indexOf('data:') !== 0 && img.src.indexOf('blob:') !== 0) {
                    
                    const wAttr = img.getAttribute('width') || '';
                    const hAttr = img.getAttribute('height') || '';
                    const wStyle = img.style.width || '';
                    const hStyle = img.style.height || '';
                    
                    const width = parseInt(wAttr || wStyle || '0', 10);
                    const height = parseInt(hAttr || hStyle || '0', 10);
                    
                    // Look for typical horizontal line dimensions from Word (wide and very thin)
                    // e.g. height: 1.5pt, 2px, etc.
                    if ((width > 50 || wStyle.includes('%')) && (height <= 20 || hStyle.includes('pt') || hStyle.includes('px'))) {
                      const hr = document.createElement('hr');
                      hr.style.border = '0';
                      hr.style.borderTop = '1pt solid black';
                      hr.style.width = '40%';
                      hr.style.margin = '5px auto';
                      img.replaceWith(hr);
                    } else if (height <= 30 && width >= 50) {
                      const hr = document.createElement('hr');
                      hr.style.border = '0';
                      hr.style.borderTop = '1pt solid black';
                      hr.style.width = '40%';
                      hr.style.margin = '5px auto';
                      img.replaceWith(hr);
                    } else {
                      // Remove other broken local images from Word
                      img.remove();
                    }
                    changed = true;
                  }
                });
                
                if (changed) {
                  handleInput();
                }
              }
            }, 50);
          }}
          className="bg-white outline-none word-editor-content shrink-0 w-full"
          style={{
            minHeight: '100%',
            lineHeight: '1.2'
          }}
          placeholder="Nhập yêu cầu chung..."
        ></div>
      </div>
    </div>
  );
}
