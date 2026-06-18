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

  const formatText = (command: string) => {
    document.execCommand(command, false, undefined);
    editorRef.current?.focus();
    handleInput();
  };

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
            td.style.backgroundColor = '';
          });
          
          doc.querySelectorAll('*').forEach(el => {
              if (el instanceof HTMLElement) {
                 el.style.backgroundColor = '';
              }
          });

          // Unpack lists if they cause layout jumps. We just preserve text indentation
          doc.querySelectorAll('li').forEach(li => {
              const div = document.createElement('div');
              div.innerHTML = `- ${li.innerHTML}`;
              div.style.marginLeft = '1rem';
              li.replaceWith(div);
          });
          doc.querySelectorAll('ul, ol').forEach(list => {
              const frag = document.createDocumentFragment();
              while (list.firstChild) {
                  frag.appendChild(list.firstChild);
              }
              list.replaceWith(frag);
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
          
          <div className="w-px h-6 bg-gray-300 mx-1"></div>
          
          <div className="relative flex items-center">
            <input
              type="file"
              accept=".docx"
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
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
        </div>
      )}

      {/* Editor Content */}
      <div className="bg-gray-100 overflow-y-auto overflow-x-auto p-4 sm:p-8 flex justify-center h-[400px] min-h-[200px] resize-y relative will-change-transform border-b border-white">
        <div
          ref={editorRef}
          contentEditable={!readOnly}
          onInput={handleInput}
          onBlur={handleInput}
          className="bg-white shadow-md outline-none word-editor-content shrink-0"
          style={{
            width: '210mm',
            minHeight: '297mm',
            padding: '20mm 20mm', // standard word margins roughly
            lineHeight: '1.2'
          }}
          placeholder="Nhập yêu cầu chung..."
        ></div>
      </div>
    </div>
  );
}
