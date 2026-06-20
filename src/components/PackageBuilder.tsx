import React, { useState } from 'react';
import { FileDown, GripVertical, CheckSquare, Square, Trash2, Save, FileSpreadsheet, Search, Loader2 } from 'lucide-react';
import * as xlsx from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Material, BiddingPackage } from '../types';

interface PackageBuilderProps {
  savedMaterials: Material[];
  savedPackages: BiddingPackage[];
  setSavedPackages: React.Dispatch<React.SetStateAction<BiddingPackage[]>>;
  isAuthenticated: boolean;
  requireAuth: (callback: () => void) => void;
}

export function PackageBuilder({ savedMaterials, savedPackages, setSavedPackages, isAuthenticated, requireAuth }: PackageBuilderProps) {
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [packageHiddenTags, setPackageHiddenTags] = useState<Record<string, string[]>>({});
  const [packageName, setPackageName] = useState('');
  const [currentPackageId, setCurrentPackageId] = useState<string | null>(null);
  const [searchMaterialCode, setSearchMaterialCode] = useState('');
  const [searchMaterialName, setSearchMaterialName] = useState('');
  const [searchPackage, setSearchPackage] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  
  // Modals
  const [deletePackageId, setDeletePackageId] = useState<string | null>(null);

  // Toggle selection
  const toggleSelection = (id: string) => {
    if (selectedIds.includes(id)) {
      setSelectedIds(selectedIds.filter(selectedId => selectedId !== id));
    } else {
      setSelectedIds([...selectedIds, id]);
    }
  };

  const moveItem = (index: number, direction: 'up' | 'down') => {
    if ((direction === 'up' && index === 0) || (direction === 'down' && index === selectedIds.length - 1)) return;
    
    const newSelectedIds = [...selectedIds];
    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    [newSelectedIds[index], newSelectedIds[targetIndex]] = [newSelectedIds[targetIndex], newSelectedIds[index]];
    setSelectedIds(newSelectedIds);
  };

  const moveItemTo = (currentIndex: number, newIndexStr: string) => {
    let newIndex = parseInt(newIndexStr, 10);
    if (isNaN(newIndex)) return;
    newIndex = newIndex - 1; // 1-based to 0-based
    if (newIndex < 0) newIndex = 0;
    if (newIndex >= selectedIds.length) newIndex = selectedIds.length - 1;
    
    if (currentIndex === newIndex) return;
    
    const newSelectedIds = [...selectedIds];
    const [item] = newSelectedIds.splice(currentIndex, 1);
    newSelectedIds.splice(newIndex, 0, item);
    setSelectedIds(newSelectedIds);
  };

  const removeSelected = (id: string) => {
    setSelectedIds(selectedIds.filter(selectedId => selectedId !== id));
  };

  const handleSavePackage = () => {
    requireAuth(async () => {
      if (!packageName.trim()) {
        alert('Vui lòng nhập tên gói thầu để lưu!');
        return;
      }
      if (selectedIds.length === 0) {
        alert('Vui lòng chọn ít nhất 1 vật tư!');
        return;
      }

      const pkg: BiddingPackage = {
        id: currentPackageId || crypto.randomUUID(),
        name: packageName,
        materialIds: selectedIds,
        hiddenTags: packageHiddenTags,
      };

      setIsSaving(true);
      try {
        const { auth } = await import('../lib/firebase.ts');
        const token = await auth.currentUser?.getIdToken();
        const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || '';
        const res = await fetch(`${API_BASE_URL}/api/packages`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
          },
          body: JSON.stringify(pkg)
        });
        if (!res.ok) throw new Error('Network error');

        setSavedPackages(prev => {
          const existing = prev.findIndex(p => p.id === pkg.id);
          if (existing >= 0) {
            const copy = [...prev];
            copy[existing] = pkg;
            return copy;
          }
          return [...prev, pkg];
        });

        setCurrentPackageId(pkg.id);
        alert('Đã lưu tiêu chuẩn gói thầu thành công!');
      } catch (err) {
        console.error(err);
        alert('Lỗi lưu gói thầu');
      } finally {
        setIsSaving(false);
      }
    });
  };

  const loadPackage = (pkg: BiddingPackage) => {
    setCurrentPackageId(pkg.id);
    setPackageName(pkg.name);
    // filter out deleted materials from the package
    const validIds = pkg.materialIds.filter(id => savedMaterials.some(m => m.id === id));
    setSelectedIds(validIds);
    setPackageHiddenTags(pkg.hiddenTags || {});
  };

  const createNewPackage = () => {
    requireAuth(() => {
      setCurrentPackageId(null);
      setPackageName('');
      setSelectedIds([]);
      setPackageHiddenTags({});
    });
  };

  const deletePackage = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    requireAuth(() => {
      setDeletePackageId(id);
    });
  };
  
  const confirmDeletePackage = async () => {
    if (deletePackageId) {
      try {
        const { auth } = await import('../lib/firebase.ts');
        const token = await auth.currentUser?.getIdToken();
        const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || '';
        const res = await fetch(`${API_BASE_URL}/api/packages/${deletePackageId}`, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${token}`
          }
        });
        if (!res.ok) throw new Error('Network error');

        setSavedPackages(prev => prev.filter(p => p.id !== deletePackageId));
        if (currentPackageId === deletePackageId) {
          createNewPackage();
        }
        setDeletePackageId(null);
      } catch (err) {
        console.error(err);
        alert('Lỗi xóa gói thầu');
      }
    }
  };

  const getFilteredExcelData = (mat: Material) => {
    if (!mat.excelData || mat.excelData.length === 0) return { data: [], merges: [] };

    const hiddenTagsForMat = packageHiddenTags[mat.id] || [];
    const isRowHidden = (rIdx: number) => {
       if (!mat.rowTags || !mat.rowTags[rIdx]) return false;
       const tags = mat.rowTags[rIdx].split(',').map(t => t.trim()).filter(Boolean);
       return tags.some(tag => hiddenTagsForMat.includes(tag));
    };

    let newMergeRefs: xlsx.Range[] = [];
    let rMapping: number[] = []; 
    let newExcelData: string[][] = [];
    let newRowIndex = 0;
    
    for (let rIdx = 0; rIdx < mat.excelData.length; rIdx++) {
        if (isRowHidden(rIdx)) {
           rMapping.push(-1); // dropped
        } else {
           rMapping.push(newRowIndex);
           newExcelData.push(mat.excelData[rIdx]);
           newRowIndex++;
        }
    }
    
    for (const merge of (mat.excelMerges || [])) {
        let startR = merge.s.r;
        while(startR <= merge.e.r && rMapping[startR] === -1) startR++; 
        let endR = merge.e.r;
        while(endR >= merge.s.r && rMapping[endR] === -1) endR--; 
        
        if (startR <= endR) {
            newMergeRefs.push({
               s: { r: rMapping[startR], c: merge.s.c },
               e: { r: rMapping[endR], c: merge.e.c }
            });
        }
    }
    return { data: newExcelData, merges: newMergeRefs };
  };

  const handleExportWord = () => {
    if (selectedIds.length === 0) {
      alert("Vui lòng chọn ít nhất 1 vật tư để xuất file!");
      return;
    }

    const selectedMaterials = selectedIds.map(id => savedMaterials.find(m => m.id === id)).filter(Boolean) as Material[];

    let contentHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${packageName || 'Tieu_Chuan_Goi_Thau'}</title>
        <style>
          body { font-family: 'Times New Roman', Times, serif; font-size: 13pt; line-height: 1.15; margin: 0; padding: 0; }
          h2 { font-size: 16pt; font-weight: bold; margin: 0pt 0pt 6pt 0pt; text-align: left; }
          p { margin: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt; }
          .rich-text p, .rich-text div, .rich-text li { line-height: 1.15; }
          ul, ol { margin-top: 0pt; margin-bottom: 0pt; }
          li { line-height: 1.15; }
          table { width: 100%; border-collapse: collapse; margin: 0pt; mso-table-lspace: 0pt; mso-table-rspace: 0pt; border: 1px solid black; }
          table, th, td { border: 1px solid black; }
          th, td { padding: 2pt 4pt; mso-padding-alt: 2pt 4pt 2pt 4pt; text-align: left; vertical-align: middle; }
          th p, td p, th div, td div { margin: 0pt !important; padding: 0pt !important; mso-margin-top-alt: 0pt !important; mso-margin-bottom-alt: 0pt !important; line-height: 1.0 !important; text-align: left; }
          th { font-weight: bold; background-color: transparent; }
          .material-section { margin-bottom: 20pt; page-break-inside: avoid; }
        </style>
      </head>
      <body>
    `;

    selectedMaterials.forEach((mat, index) => {
      contentHtml += `<div class="material-section">`;
      contentHtml += `<h2>II.${index + 1}. ${mat.name || 'Vật tư chưa có tên'}</h2>`;
      
      if (mat.richText) {
        contentHtml += `<p style="margin: 0pt; padding: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;"><strong>II.${index + 1}.1. Yêu cầu chung:</strong></p>`;
        const cleanedRichText = mat.richText;
        contentHtml += `<div class="rich-text" style="margin: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;">${cleanedRichText}</div>`;
      }

      const { data: filteredData, merges: filteredMerges } = getFilteredExcelData(mat);

      if (filteredData && filteredData.length > 0) {
        contentHtml += `<p style="margin: 0pt; padding: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;"><strong>II.${index + 1}.2. Bảng thông số kỹ thuật:</strong></p>`;
        contentHtml += `<table border="1" cellpadding="0" cellspacing="0" style="margin: 0pt; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">`;
        filteredData.forEach((row, rIdx) => {
          contentHtml += `<tr>`;
          row.forEach((cell, cIdx) => {
            let rowSpan = 1;
            let colSpan = 1;
            let skip = false;

            if (filteredMerges) {
                for (const merge of filteredMerges) {
                    if (merge.s.r === rIdx && merge.s.c === cIdx) {
                        rowSpan = (merge.e.r - merge.s.r) + 1;
                        colSpan = (merge.e.c - merge.s.c) + 1;
                    } else if (
                        rIdx >= merge.s.r && rIdx <= merge.e.r &&
                        cIdx >= merge.s.c && cIdx <= merge.e.c
                    ) {
                        skip = true;
                    }
                }
            }

            if (!skip) {
              const cellTag = (rIdx === 0 && filteredData.length > 1) ? 'th' : 'td';
              const rawContent = (cell || '').toString().trim();
              
              let displayContent = rawContent ? rawContent.replace(/\r\n/g, '<br/>').replace(/[\r\n]/g, '<br/>') : '&nbsp;';
              
              if (rawContent.startsWith('[IMG:data:image/') && rawContent.endsWith(']')) {
                const base64Src = rawContent.slice(5, -1);
                displayContent = `<img src="${base64Src}" style="max-width: 150px; max-height: 150px;" />`;
              }
              
              const cellStyle = rIdx === 0 ? 'white-space: nowrap;' : '';
              contentHtml += `<${cellTag} rowspan="${rowSpan}" colspan="${colSpan}" style="${cellStyle}"><p style="margin: 0pt; padding: 0pt; line-height: 1.1;">${displayContent}</p></${cellTag}>`;
            }
          });
          contentHtml += `</tr>`;
        });
        contentHtml += `</table>`;
      }
      
      contentHtml += `</div>`;
    });

    contentHtml += `</body></html>`;
    const blob = new Blob(['\ufeff', contentHtml], { type: 'application/msword' });
    saveAs(blob, `${packageName || 'Tieu_Chuan_Goi_Thau'}.doc`);
  };

  const handleExportExcel = async () => {
    if (selectedIds.length === 0) {
      alert("Vui lòng chọn ít nhất 1 vật tư để xuất file!");
      return;
    }
    
    const selectedMaterials = selectedIds.map(id => savedMaterials.find(m => m.id === id)).filter(Boolean) as Material[];
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ThongSoKyThuat");

    let currentRow = 1;

    selectedMaterials.forEach((mat, idx) => {
      // Add Material Name as header
      const headerCell = worksheet.getCell(currentRow, 1);
      headerCell.value = `${idx + 1}. ${mat.name}`;
      headerCell.font = { bold: true, size: 14 };
      worksheet.mergeCells(currentRow, 1, currentRow, 5); // arbitrarily merge 5 cols to give space
      currentRow += 1;

      const { data: filteredData, merges: filteredMerges } = getFilteredExcelData(mat);

      if (filteredData && filteredData.length > 0) {
         // Insert rows one by one
         filteredData.forEach((row, rowIdx) => {
             row.forEach((cellValue, colIdx) => {
                 const cell = worksheet.getCell(currentRow + rowIdx, colIdx + 1);
                 
                 if (typeof cellValue === 'string' && cellValue.startsWith('[IMG:data:image/') && cellValue.endsWith(']')) {
                    cell.value = '';
                    const base64 = cellValue.slice(5, -1);
                    const extMatch = base64.match(/data:image\/(.+?);base64,/);
                    const ext = extMatch ? extMatch[1] : 'png';
                    
                    try {
                      const imageId = workbook.addImage({
                        base64: base64,
                        extension: ext as any,
                      });
                      
                      worksheet.addImage(imageId, {
                        tl: { col: colIdx + 0.05, row: currentRow + rowIdx - 1 + 0.05 },
                        br: { col: colIdx + 1 - 0.05, row: currentRow + rowIdx - 0.05 },
                        editAs: 'oneCell'
                      });

                      const worksheetRow = worksheet.getRow(currentRow + rowIdx);
                      if ((worksheetRow.height || 15) < 100) {
                          worksheetRow.height = 100;
                      }
                    } catch (err) {
                      console.error('Lỗi khi chèn ảnh vào Excel:', err);
                      // In case of error
                    }
                 } else {
                    cell.value = cellValue;
                 }
                 
                 cell.alignment = { wrapText: true, vertical: 'middle' };
                 // Adding borders
                 cell.border = {
                     top: {style:'thin'},
                     left: {style:'thin'},
                     bottom: {style:'thin'},
                     right: {style:'thin'}
                 };
             });
         });
         
         // Apply merges
         if (filteredMerges && filteredMerges.length > 0) {
           filteredMerges.forEach(merge => {
              const top = currentRow + merge.s.r;
              const left = merge.s.c + 1;
              const bottom = currentRow + merge.e.r;
              const right = merge.e.c + 1;
              worksheet.mergeCells(top, left, bottom, right);
           });
         }
         
         currentRow += filteredData.length;
      } else {
         const cell = worksheet.getCell(currentRow, 1);
         cell.value = "Không có bảng thông số";
         cell.font = { italic: true };
         currentRow += 1;
      }
      
      // Blank row between items
      currentRow += 1;
    });

    // Auto-fit columns (best effort, up to max width)
    worksheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell?.({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? cell.value.toString().length : 10;
            if (columnLength > maxLength) {
                maxLength = columnLength;
            }
        });
        column.width = Math.min(maxLength < 10 ? 10 : maxLength, 50) + 2; // limit to 50 for wrapping
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const data = new Blob([buffer], { type: 'application/octet-stream' });
    saveAs(data, `${packageName || 'Bang_Thong_So_Goi_Thau'}.xlsx`);
  };



  const filteredMaterials = savedMaterials
    .filter(m => 
      (m.code || '').toLowerCase().includes(searchMaterialCode.toLowerCase()) &&
      (m.name || '').toLowerCase().includes(searchMaterialName.toLowerCase())
    )
    .sort((a, b) => (a.name || '').localeCompare(b.name || ''));

  const filteredPackages = savedPackages
    .filter(p => p.name.toLowerCase().includes(searchPackage.toLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name));

  return (
    <div className="flex flex-col gap-6 w-full min-h-[calc(100vh-8rem)]">
      
      {/* Top Banner: Saved Packages List */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-4">
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-semibold text-gray-800">Các gói thầu đã lưu</h3>
          <div className="relative w-64 text-sm">
            <Search className="absolute left-2.5 top-2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Tìm gói thầu..."
              value={searchPackage}
              onChange={e => setSearchPackage(e.target.value)}
              className="w-full pl-8 pr-4 py-1.5 rounded-md border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
            />
          </div>
        </div>
        <div className="flex gap-2 overflow-x-auto pb-2">
          {isAuthenticated && (
            <button
               onClick={createNewPackage}
               className="px-4 py-2 border border-dashed border-gray-400 text-gray-600 rounded-lg whitespace-nowrap hover:bg-gray-50 font-medium"
            >
               + Tạo gói mới
            </button>
          )}
          {filteredPackages.map(pkg => (
            <div
              key={pkg.id}
              className={`flex items-center gap-1 px-3 py-1.5 border rounded-lg whitespace-nowrap transition-colors group ${
                currentPackageId === pkg.id 
                  ? 'border-blue-500 bg-blue-50 text-blue-700'
                  : 'border-gray-200 hover:border-blue-300 text-gray-700 hover:bg-gray-50'
              }`}
            >
              <button
                onClick={() => loadPackage(pkg)}
                className="font-medium mr-1"
              >
                {pkg.name} ({pkg.materialIds.length} vật tư)
              </button>
              {isAuthenticated && (
                <button
                  onClick={(e) => deletePackage(pkg.id, e)}
                  className={`p-1 rounded-md transition-colors ${
                    currentPackageId === pkg.id ? 'text-blue-400 hover:text-red-500 hover:bg-red-50' : 'text-gray-300 group-hover:text-gray-400 hover:!text-red-500 hover:!bg-red-50'
                  }`}
                  title="Xóa gói thầu"
                >
                  <Trash2 size={14} />
                </button>
              )}
            </div>
          ))}
          {filteredPackages.length === 0 && savedPackages.length > 0 && (
            <span className="text-gray-400 text-sm py-2">Không tìm thấy gói thầu.</span>
          )}
        </div>
      </div>

      <div className="flex flex-col md:flex-row gap-6 w-full flex-1">
        {/* Cột trái: Danh sách vật tư để chọn */}
        <div className="w-full md:w-1/2 bg-white rounded-xl shadow-sm border border-gray-200 flex flex-col overflow-hidden max-h-[600px]">
          <div className="px-5 py-4 border-b border-gray-100 bg-gray-50 space-y-3">
            <div>
              <h3 className="font-semibold text-gray-800">Danh sách vật tư hiện có</h3>
              <p className="text-xs text-gray-500 mt-1">Chọn các vật tư muốn thêm vào gói thầu</p>
            </div>
            <div className="flex flex-col gap-2">
              <div className="relative text-sm">
                <Search className="absolute left-2.5 top-2.5 text-gray-400" size={16} />
                <input
                  type="text"
                  placeholder="Tìm mã vật tư..."
                  value={searchMaterialCode}
                  onChange={e => setSearchMaterialCode(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 rounded-md border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none block"
                />
              </div>
              <div className="relative text-sm">
                <Search className="absolute left-2.5 top-2.5 text-gray-400" size={16} />
                <input
                  type="text"
                  placeholder="Tìm tên vật tư..."
                  value={searchMaterialName}
                  onChange={e => setSearchMaterialName(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 rounded-md border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none block"
                />
              </div>
            </div>
          </div>
          <div className="p-3 flex-1 overflow-y-auto space-y-2 relative bg-gray-50/30">
            {filteredMaterials.length === 0 ? (
              <div className="text-center text-gray-500 text-sm mt-10">Không tìm thấy vật tư nào.</div>
            ) : (
              filteredMaterials.map(mat => {
                const isSelected = selectedIds.includes(mat.id);
                return (
                  <div 
                    key={mat.id}
                    onClick={() => toggleSelection(mat.id)}
                    className={`flex items-center gap-3 p-3 rounded-lg border cursor-pointer transition-all ${
                      isSelected ? 'border-blue-500 bg-blue-50' : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50 bg-white'
                    }`}
                  >
                    <div className={`text-${isSelected ? 'blue-600' : 'gray-400'}`}>
                      {isSelected ? <CheckSquare size={20} /> : <Square size={20} />}
                    </div>
                    <div className="overflow-hidden flex-1">
                      <h4 className="font-medium text-gray-800 text-sm truncate">{mat.name}</h4>
                      <p className="text-xs text-gray-500 truncate">{mat.code}</p>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>

        {/* Cột phải: Thứ tự vật tư trong gói thầu & Nút xuất */}
        <div className="w-full md:w-1/2 bg-white rounded-xl shadow-sm border border-gray-200 flex flex-col overflow-hidden max-h-[600px]">
          <div className="px-5 py-3 border-b border-gray-100 bg-gray-50 flex flex-col gap-3">
            
            <div className="flex items-center justify-between">
              <div>
                <h3 className="font-semibold text-gray-800">Cấu trúc hiện tại ({selectedIds.length})</h3>
                <p className="text-xs text-gray-500 mt-1">Sắp xếp bằng nút mũi tên hoặc nhập trực tiếp số thứ tự</p>
              </div>
              {isAuthenticated && (
                <button
                  onClick={handleSavePackage}
                  disabled={isSaving}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white rounded-lg text-sm font-semibold transition-all shadow"
                >
                  {isSaving ? (
                    <><Loader2 size={16} className="animate-spin" /> Đang lưu...</>
                  ) : (
                    <><Save size={16} /> Lưu gói thầu</>
                  )}
                </button>
              )}
            </div>
            
            <input
              type="text"
              placeholder="Nhập tên gói thầu (bắt buộc để lưu)..."
              value={packageName}
              onChange={e => setPackageName(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md outline-none focus:ring-2 focus:border-blue-500 focus:ring-blue-500 text-sm font-medium"
            />
            
            <div className="flex items-center gap-2">
              <button
                onClick={handleExportWord}
                disabled={selectedIds.length === 0}
                className={`flex-1 flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all ${
                  selectedIds.length > 0 
                    ? 'bg-blue-50 border border-blue-200 text-blue-700 hover:bg-blue-100'
                    : 'bg-gray-100 text-gray-400 cursor-not-allowed border border-transparent'
                }`}
              >
                <FileDown size={16} /> Xuất Word (Tiêu chuẩn)
              </button>
              
              <button
                onClick={handleExportExcel}
                disabled={selectedIds.length === 0}
                className={`flex-1 flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all ${
                  selectedIds.length > 0 
                    ? 'bg-emerald-50 border border-emerald-200 text-emerald-700 hover:bg-emerald-100'
                    : 'bg-gray-100 text-gray-400 cursor-not-allowed border border-transparent'
                }`}
              >
                <FileSpreadsheet size={16} /> Xuất Excel (Bảng TS)
              </button>
            </div>
          </div>
          
          <div className="p-3 flex-1 overflow-y-auto space-y-2 bg-gray-50/50">
             {selectedIds.length === 0 ? (
              <div className="text-center text-gray-400 text-sm mt-10">
                Chọn vật tư từ danh sách bên trái để thêm vào gói thầu.
              </div>
             ) : (
               selectedIds.map((id, index) => {
                 const mat = savedMaterials.find(m => m.id === id);
                 if (!mat) return null;
                 
                 const tagsSet = new Set<string>();
                 if (mat.rowTags) {
                   Object.values(mat.rowTags).forEach(tagStr => {
                     const tags = tagStr.split(',').map(t => t.trim()).filter(Boolean);
                     tags.forEach(t => tagsSet.add(t));
                   });
                 }
                 const uniqueTags = Array.from(tagsSet);
                 
                 const hiddenTagsForMat = packageHiddenTags[mat.id] || [];

                 const toggleTag = (matId: string, tag: string) => {
                   setPackageHiddenTags(prev => {
                     const currentHidden = prev[matId] || [];
                     if (currentHidden.includes(tag)) {
                       return { ...prev, [matId]: currentHidden.filter(t => t !== tag) };
                     } else {
                       return { ...prev, [matId]: [...currentHidden, tag] };
                     }
                   });
                 };
                 
                 return (
                   <div key={`${id}-${index}`} className="flex flex-col p-3 bg-white rounded-lg border border-gray-200 shadow-sm gap-2">
                     <div className="flex items-center gap-2">
                       <div className="flex flex-col gap-1 items-center justify-center shrink-0 text-gray-400">
                         <button 
                           onClick={() => moveItem(index, 'up')} 
                           disabled={index === 0}
                           className="p-0.5 hover:text-blue-600 disabled:opacity-30 disabled:hover:text-gray-400"
                         >
                           ▲
                         </button>
                         <button 
                           onClick={() => moveItem(index, 'down')} 
                           disabled={index === selectedIds.length - 1}
                           className="p-0.5 hover:text-blue-600 disabled:opacity-30 disabled:hover:text-gray-400"
                         >
                           ▼
                         </button>
                       </div>
                       
                       <input
                         key={`order-${id}-${index}`}
                         type="number"
                         min="1"
                         max={selectedIds.length}
                         className="flex bg-gray-100 border border-transparent hover:bg-gray-200 focus:bg-white focus:border-blue-400 focus:ring-1 focus:ring-blue-400 rounded px-1 py-1 shrink-0 text-gray-600 font-mono text-sm mr-2 w-10 text-center outline-none transition-colors"
                         defaultValue={index + 1}
                         onBlur={(e) => {
                           moveItemTo(index, e.target.value);
                           e.target.value = (index + 1).toString(); // reset to current if no change
                         }}
                         onKeyDown={(e) => {
                           if (e.key === 'Enter') {
                             moveItemTo(index, (e.target as HTMLInputElement).value);
                             (e.target as HTMLInputElement).blur();
                           }
                         }}
                         title="Nhập số thứ tự để di chuyển"
                       />
                       
                       <div className="overflow-hidden flex-1">
                         <h4 className="font-medium text-gray-800 text-sm truncate">{mat.name}</h4>
                         <p className="text-xs text-gray-500 truncate">{mat.code || '---'}</p>
                       </div>

                       <button
                         onClick={() => removeSelected(id)}
                         className="p-1.5 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded transition-colors shrink-0"
                         title="Bỏ chọn"
                       >
                         <Trash2 size={16} />
                       </button>
                     </div>
                     
                     {uniqueTags.length > 0 && (
                       <div className="ml-14 border-t pt-2 mt-1">
                         <span className="text-xs text-gray-500 font-medium mb-1 block">Tham biến hiển thị:</span>
                         <div className="flex flex-wrap gap-2">
                           {uniqueTags.map(tag => {
                              const isHidden = hiddenTagsForMat.includes(tag);
                              return (
                                <label key={tag} className={`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border ${isHidden ? 'bg-gray-50 text-gray-400 border-gray-200' : 'bg-blue-50 text-blue-700 border-blue-200'}`}>
                                  <input 
                                    type="checkbox" 
                                    className="hidden"
                                    checked={!isHidden} 
                                    onChange={() => toggleTag(mat.id, tag)} 
                                  />
                                  <div className="w-3 h-3 flex items-center justify-center shrink-0">
                                    {isHidden ? <Square size={12} /> : <CheckSquare size={12} />}
                                  </div>
                                  <span>{tag}</span>
                                </label>
                              );
                           })}
                         </div>
                       </div>
                     )}
                   </div>
                 )
               })
             )}
           </div>
        </div>
      </div>
      
      {/* Delete Confirmation Modal */}
      {deletePackageId && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-5">
              <h3 className="text-lg font-bold text-gray-900 mb-2">Xóa gói thầu</h3>
              <p className="text-sm text-gray-500">
                Bạn có chắc chắn muốn xóa gói thầu này không? Hành động này không thể hoàn tác.
              </p>
            </div>
            <div className="bg-gray-50 px-5 py-4 flex items-center justify-end gap-3 border-t border-gray-100">
              <button
                onClick={() => setDeletePackageId(null)}
                className="px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-100 rounded-lg transition-colors"
              >
                Hủy bỏ
              </button>
              <button
                onClick={confirmDeletePackage}
                className="px-4 py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg shadow-sm transition-colors"
              >
                Xác nhận xóa
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

