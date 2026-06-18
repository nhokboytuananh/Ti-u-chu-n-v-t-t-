import React, { useState } from 'react';
import * as xlsx from 'xlsx';
import { Upload, Plus, Trash2 } from 'lucide-react';

interface ExcelTableProps {
  tableData: string[][];
  setTableData: React.Dispatch<React.SetStateAction<string[][]>>;
  merges: xlsx.Range[];
  setMerges: React.Dispatch<React.SetStateAction<xlsx.Range[]>>;
  rowTags?: Record<number, string>;
  setRowTags?: React.Dispatch<React.SetStateAction<Record<number, string>>>;
  readOnly?: boolean;
}

export function ExcelTable({ tableData, setTableData, merges, setMerges, rowTags = {}, setRowTags, readOnly }: ExcelTableProps) {
  const [fileName, setFileName] = useState<string>('');

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = xlsx.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Store merges if they exist
      const sheetMerges: xlsx.Range[] = worksheet['!merges'] || [];
      // we need to shift merges if we find empty rows at the top to skip
      
      // Convert to 2D array, header: 1 means array of arrays
      const jsonData = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '' });
      
      // Find the first non-empty row to determine the exact number of columns
      let columnCount = 0;
      let firstDataRowIndex = -1;
      
      for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row.some(cell => String(cell).trim() !== '')) {
           columnCount = Math.max(row.length, columnCount); // Ensure column count is accurate for all rows
           if (firstDataRowIndex === -1) {
               firstDataRowIndex = i;
           }
        }
      }

      if (firstDataRowIndex !== -1) {
        // Adjust merges to account for sliced items
        const adjustedMerges = sheetMerges
            .filter(merge => merge.e.r >= firstDataRowIndex) // keep merges that are present in our slice
            .map(merge => ({
                s: { r: Math.max(0, merge.s.r - firstDataRowIndex), c: merge.s.c },
                e: { r: merge.e.r - firstDataRowIndex, c: merge.e.c }
            }));

        // Map data ensuring every row has exactly `columnCount` elements
        const cleanedData = jsonData
          .slice(firstDataRowIndex)
          // Don't filter out entirely empty rows deeply inside because it breaks merges if rows disappear!
          // We just filter trailing empty rows at the bottom.
          .map(row => {
             const newRow = Array(columnCount).fill('');
             for (let i = 0; i < columnCount; i++) {
               newRow[i] = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
             }
             return newRow;
          });
          
        // Trim trailing empty rows carefully so we don't break merge indexes
        let lastValidRow = cleanedData.length - 1;
        while (lastValidRow >= 0 && cleanedData[lastValidRow].every(cell => cell.trim() === '')) {
            lastValidRow--;
        }
        
        const finalData = cleanedData.slice(0, lastValidRow + 1);

        if (finalData.length > 0) {
          setTableData(finalData);
          setMerges(adjustedMerges);
        }
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = ''; // allow uploading same file again
  };

  const handleCellChange = (rowIndex: number, colIndex: number, value: string) => {
    const newData = [...tableData];
    newData[rowIndex] = [...newData[rowIndex]];
    newData[rowIndex][colIndex] = value;
    setTableData(newData);
  };

  const handleTagChange = (rowIndex: number, val: string) => {
    if (!setRowTags) return;
    setRowTags(prev => ({
      ...prev,
      [rowIndex]: val
    }));
  };

  const addRow = () => {
    const colsCount = tableData[0]?.length || 4;
    setTableData([...tableData, Array(colsCount).fill('')]);
  };

  const addColumn = () => {
    setTableData(tableData.map(row => [...row, '']));
  };

  const deleteRow = (index: number) => {
    if (tableData.length <= 1) return;
    setTableData(tableData.filter((_, i) => i !== index));
    // Complex logic to adjust merges when deleting a row could be added here, 
    // but typically for viewing extracted tables it's edge-case. Let's just clear merges to be safe if structure changes manually
    setMerges([]); 
  };

  const deleteColumn = (index: number) => {
    if (tableData[0]?.length <= 1) return;
    setTableData(tableData.map(row => row.filter((_, i) => i !== index)));
    setMerges([]);
  };

  // Helper to check if cell is merged and should be displayed
  const getCellRenderingArgs = (rowIndex: number, colIndex: number) => {
    let rowSpan = 1;
    let colSpan = 1;
    let hideCell = false;

    for (const merge of merges) {
      // If cell is top-left of a merge
      if (merge.s.r === rowIndex && merge.s.c === colIndex) {
        rowSpan = (merge.e.r - merge.s.r) + 1;
        colSpan = (merge.e.c - merge.s.c) + 1;
        return { hideCell, rowSpan, colSpan };
      }
      
      // If cell is contained within a merge bounds (but not the starting cell)
      if (
        rowIndex >= merge.s.r && rowIndex <= merge.e.r &&
        colIndex >= merge.s.c && colIndex <= merge.e.c
      ) {
        hideCell = true;
        return { hideCell, rowSpan, colSpan };
      }
    }

    return { hideCell, rowSpan, colSpan };
  };

  // Init empty table if needed
  if (tableData.length === 0) {
      if (readOnly) {
         return <div className="p-8 text-center text-gray-500 border border-gray-200 rounded-md bg-gray-50">Không có dữ liệu bảng.</div>;
      }
      return (
        <div className="w-full bg-white rounded-md border border-gray-300 shadow-sm overflow-hidden">
             <div className="flex items-center justify-between p-3 border-b border-gray-200 bg-gray-50">
               <div className="flex items-center gap-3">
                 <div className="relative">
                   <input
                     type="file"
                     accept=".xlsx,.xls,.csv"
                     onChange={handleFileUpload}
                     className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                     title="Nhập từ file Excel"
                   />
                   <button
                     type="button"
                     className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-emerald-700 bg-emerald-50 hover:bg-emerald-100 rounded transition-colors border border-emerald-200"
                   >
                     <Upload size={16} />
                     <span>Trích xuất từ Excel</span>
                   </button>
                 </div>
               </div>
               <div className="flex gap-2">
                 <button type="button" onClick={() => setTableData([['Cột 1', 'Cột 2']])} className="text-sm text-blue-600 hover:underline">Tạo bảng trống</button>
               </div>
             </div>
             <div className="p-8 text-center text-gray-500">
               Chưa có dữ liệu bảng.
             </div>
        </div>
      );
  }

  return (
    <div className="w-full bg-white rounded-md border border-gray-300 shadow-sm overflow-hidden">
      {!readOnly && (
        <div className="flex items-center justify-between p-3 border-b border-gray-200 bg-gray-50 flex-wrap gap-2">
          <div className="flex items-center gap-3">
            <div className="relative">
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                title="Nhập từ file Excel"
              />
              <button
                type="button"
                className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-emerald-700 bg-emerald-50 hover:bg-emerald-100 rounded transition-colors border border-emerald-200"
              >
                <Upload size={16} />
                <span>Trích xuất từ Excel</span>
              </button>
            </div>
            {fileName && <span className="text-sm text-gray-600 font-medium truncate max-w-[200px]">{fileName}</span>}
          </div>
          
          <div className="flex gap-2">
            <button
              type="button"
              onClick={addRow}
              className="flex items-center gap-1 px-2 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded transition-colors"
            >
              <Plus size={16} /> Thêm Hàng
            </button>
            <button
              type="button"
              onClick={addColumn}
              className="flex items-center gap-1 px-2 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded transition-colors"
            >
              <Plus size={16} /> Thêm Cột
            </button>
          </div>
        </div>
      )}

      <div className="overflow-x-auto p-1">
        <table className="w-full text-left border-collapse min-w-max">
          <thead>
            {tableData.slice(0, 1).map((row, rowIndex) => (
              <tr key={rowIndex} className="bg-gray-100/50">
                {row.map((cell, colIndex) => {
                  const { hideCell, rowSpan, colSpan } = getCellRenderingArgs(rowIndex, colIndex);
                  if (hideCell) return null;
                  
                  return (
                  <th key={colIndex} rowSpan={rowSpan} colSpan={colSpan} className="p-0 border border-gray-300 font-semibold text-gray-700 min-w-[150px] relative group bg-gray-50">
                    <textarea
                      value={cell}
                      readOnly={readOnly}
                      onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
                      className={`w-full min-h-[40px] p-2 bg-transparent outline-none font-semibold text-gray-800 transition-shadow resize ${readOnly ? 'resize-none' : 'focus:bg-white focus:ring-2 focus:ring-blue-500'}`}
                      placeholder={readOnly ? '' : `Cột ${colIndex + 1}`}
                      style={{ fieldSizing: 'content' } as any} // modern css feature, fallbacks to manual resize
                    />
                    {!readOnly && tableData[0]?.length > 1 && (
                       <button 
                         type="button"
                         onClick={() => deleteColumn(colIndex)}
                         className="absolute -top-3 -right-2 bg-red-100 text-red-600 p-1 rounded-full opacity-0 group-hover:opacity-100 transition-opacity shadow-sm hover:bg-red-200 z-10"
                         title="Xóa cột này"
                       >
                         <Trash2 size={12}/>
                       </button>
                    )}
                  </th>
                )})}
                {!readOnly && (
                  <>
                    <th className="w-[40px] border border-gray-300 bg-gray-50"></th>
                    <th className="w-[150px] border border-gray-300 bg-red-50 p-2 text-center text-sm font-semibold text-red-700 whitespace-nowrap" title="Nhập tham biến để ẩn/hiện hàng khi xuất">
                      Tham biến
                    </th>
                  </>
                )}
              </tr>
            ))}
          </thead>
          <tbody>
            {tableData.slice(1).map((row, rowIndex) => (
              <tr key={rowIndex + 1} className="hover:bg-gray-50/50 transition-colors group">
                {row.map((cell, colIndex) => {
                  const { hideCell, rowSpan, colSpan } = getCellRenderingArgs(rowIndex + 1, colIndex);
                  if (hideCell) return null;

                  return (
                  <td key={colIndex} rowSpan={rowSpan} colSpan={colSpan} className="p-0 border border-gray-300 relative">
                    <textarea
                      value={cell}
                      readOnly={readOnly}
                      onChange={(e) => handleCellChange(rowIndex + 1, colIndex, e.target.value)}
                      className={`w-full min-h-[40px] h-full p-2 bg-transparent outline-none transition-shadow ${readOnly ? 'resize-none' : 'focus:bg-white focus:ring-2 focus:ring-blue-500 resize'}`}
                    />
                  </td>
                )})}
                {!readOnly && (
                  <>
                    <td className="w-[40px] border border-gray-300 bg-white text-center">
                      <button 
                         type="button"
                         onClick={() => deleteRow(rowIndex + 1)}
                         className="text-gray-400 hover:text-red-600 p-1 rounded-full hover:bg-red-50 transition-colors opacity-0 group-[&:hover]:opacity-100 mx-auto"
                         title="Xóa hàng này"
                       >
                         <Trash2 size={16}/>
                       </button>
                    </td>
                    <td className="border border-gray-300 bg-red-50/30 w-[150px] max-w-[150px]">
                      <input
                        type="text"
                        value={rowTags[rowIndex + 1] || ''}
                        onChange={(e) => handleTagChange(rowIndex + 1, e.target.value)}
                        placeholder="VD: ACSR 50/8"
                        className="w-full h-full min-h-[40px] text-sm px-2 outline-none bg-transparent"
                      />
                    </td>
                  </>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
