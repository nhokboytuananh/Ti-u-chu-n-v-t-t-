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
  const [packageHiddenTables, setPackageHiddenTables] = useState<Record<string, string[]>>({});
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

      // Fallback for crypto.randomUUID in non-secure contexts
      const generateUUID = () => {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
          return crypto.randomUUID();
        }
        return 'pkg_' + Math.random().toString(36).substring(2, 15) + '_' + Date.now().toString(36);
      };

      const pkg: BiddingPackage = {
        id: currentPackageId || generateUUID(),
        name: packageName,
        materialIds: selectedIds,
        hiddenTags: packageHiddenTags,
        hiddenTables: packageHiddenTables,
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
        if (!res.ok) {
          const errMsg = await res.text();
          throw new Error(errMsg || `HTTP error ${res.status}`);
        }

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
      } catch (err: any) {
        console.error(err);
        alert('Lỗi lưu gói thầu: ' + (err?.message || 'Có lỗi xảy ra khi kết nối máy chủ.'));
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
    setPackageHiddenTables(pkg.hiddenTables || {});
  };

  const createNewPackage = () => {
    requireAuth(() => {
      setCurrentPackageId(null);
      setPackageName('');
      setSelectedIds([]);
      setPackageHiddenTags({});
      setPackageHiddenTables({});
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

  const getFilteredTableData = (table: import('../types').ExcelTableConfig, matId: string) => {
    if (!table.data || table.data.length === 0) return { data: [], merges: [] };

    const hiddenTagsForMat = packageHiddenTags[matId] || [];
    const isRowHidden = (rIdx: number) => {
       if (!table.tags || !table.tags[rIdx]) return false;
       const tags = table.tags[rIdx].split(',').map(t => t.trim()).filter(Boolean);
       return tags.some(tag => hiddenTagsForMat.includes(tag));
    };

    let newMergeRefs: xlsx.Range[] = [];
    let rMapping: number[] = []; 
    let newExcelData: string[][] = [];
    let newRowIndex = 0;
    
    for (let rIdx = 0; rIdx < table.data.length; rIdx++) {
        if (isRowHidden(rIdx)) {
           rMapping.push(-1); // dropped
        } else {
           rMapping.push(newRowIndex);
           newExcelData.push(table.data[rIdx]);
           newRowIndex++;
        }
    }
    
    for (const merge of (table.merges || [])) {
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
          .rich-text img {
             display: inline-block !important;
             vertical-align: middle;
             margin: 4.5pt 6pt;
             max-width: 100% !important;
             height: auto !important;
          }
        </style>
      </head>
      <body>
    `;

    // Add Summary Document Requirements Table
    const hasAnyDocs = selectedMaterials.some(m => {
        const reqs = m.docRequirements || {};
        return reqs.typeTest || reqs.catalog || reqs.endUser || reqs.iso;
    });

    if (hasAnyDocs) {
        contentHtml += `<h2 style="font-size: 14pt;">Danh mục các tài liệu chứng minh nguồn gốc và chất lượng hàng hóa</h2>`;
        contentHtml += `<table border="1" cellpadding="0" cellspacing="0" style="margin-bottom: 20pt; text-align: center; border-collapse: collapse; width: 100%; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">`;
        contentHtml += `
          <thead>
            <tr>
              <th style="width: 5%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">TT</p></th>
              <th style="width: 35%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">Tên vật tư - thiết bị</p></th>
              <th style="width: 15%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">Biên bản thí nghiệm điển hình</p></th>
              <th style="width: 15%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">Tài liệu kỹ thuật (bản vẽ, Catalogue, ...)</p></th>
              <th style="width: 15%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">Xác nhận của đơn vị sử dụng cuối cùng</p></th>
              <th style="width: 15%; text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">Chứng chỉ quản lý chất lượng ISO 9001 của nhà sản xuất</p></th>
            </tr>
          </thead>
          <tbody>
        `;
        selectedMaterials.forEach((mat, index) => {
            const reqs = mat.docRequirements || {};
            
            // Lấy danh sách các biến (tags) cho vật tư này
            const tagsSet = new Set<string>();
            const tagNames: Record<string, string> = {};
            if (mat.tables) {
                mat.tables.forEach(table => {
                    if (table.tags && table.data) {
                        Object.entries(table.tags).forEach(([rowStr, tagStr]) => {
                            const rowIndex = parseInt(rowStr, 10);
                            const tags = tagStr.split(',').map((t: string) => t.trim()).filter(Boolean);
                            const variantName = (table.data[rowIndex] && table.data[rowIndex].length > 1) ? table.data[rowIndex][1] : '';
                            tags.forEach((t: string) => {
                                tagsSet.add(t);
                                if (!tagNames[t] && variantName) {
                                    tagNames[t] = variantName.toString().trim();
                                }
                            });
                        });
                    }
                });
            }
            const hiddenTagsForMat = packageHiddenTags[mat.id] || [];
            const activeVariants = Array.from(tagsSet).filter(tag => !hiddenTagsForMat.includes(tag));
            const hasVariants = activeVariants.length > 0;

            contentHtml += `
              <tr>
                <td style="text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${index + 1}</p></td>
                <td style="text-align: left; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: left;">${mat.name || ''}</p></td>
                <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${hasVariants ? '' : (reqs.typeTest ? 'x' : '')}</p></td>
                <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${hasVariants ? '' : (reqs.catalog ? 'x' : '')}</p></td>
                <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${hasVariants ? '' : (reqs.endUser ? 'x' : '')}</p></td>
                <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${hasVariants ? '' : (reqs.iso ? 'x' : '')}</p></td>
              </tr>
            `;

            if (hasVariants) {
                activeVariants.forEach(variant => {
                    const displayVariantName = tagNames[variant] || variant;
                    contentHtml += `
                      <tr>
                        <td style="text-align: center; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;"></p></td>
                        <td style="text-align: left; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: left;">${displayVariantName}</p></td>
                        <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${reqs.typeTest ? 'x' : ''}</p></td>
                        <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${reqs.catalog ? 'x' : ''}</p></td>
                        <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${reqs.endUser ? 'x' : ''}</p></td>
                        <td style="text-align: center; font-family: Arial; padding: 0pt;"><p style="margin: 0pt; line-height: 1.0; text-align: center;">${reqs.iso ? 'x' : ''}</p></td>
                      </tr>
                    `;
                });
            }
        });
        contentHtml += `</tbody></table>`;
    }

    selectedMaterials.forEach((mat, index) => {
      contentHtml += `<div class="material-section">`;
      contentHtml += `<h2>II.${index + 1}. ${mat.name || 'Vật tư chưa có tên'}</h2>`;
      
      if (mat.richText) {
        contentHtml += `<p style="margin: 0pt; padding: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;"><strong>II.${index + 1}.1. Yêu cầu chung:</strong></p>`;
        let cleanedRichText = mat.richText || '';
        try {
          const parser = new DOMParser();
          const doc = parser.parseFromString(cleanedRichText, 'text/html');
          const imgs = doc.querySelectorAll('img');
          imgs.forEach(img => {
            const originalStyle = img.getAttribute('style') || '';
            const originalWidthAttr = img.getAttribute('width') || '';
            
            // Tìm kiếm chiều rộng từ style
            let styleWidth = '';
            const styleWidthMatch = originalStyle.match(/width\s*:\s*([^;]+)/i);
            if (styleWidthMatch) {
              styleWidth = styleWidthMatch[1].trim();
            }
            
            // Xác định xem ảnh này là ảnh cỡ nhỏ hay ảnh cần giãn to
            let isSmallImage = false;
            let numericWidth = 0;
            
            if (originalWidthAttr) {
              numericWidth = parseInt(originalWidthAttr, 10);
            } else if (styleWidth && styleWidth.includes('px')) {
              numericWidth = parseInt(styleWidth, 10);
            }
            
            if (numericWidth > 0 && numericWidth < 300) {
              isSmallImage = true;
            }
            
            // Xây dựng Style tương thích xuất sắc cho Word và trình duyệt
            let newStyle = 'display: inline-block; vertical-align: middle; margin: 4.5pt 6pt; ';
            
            if (isSmallImage) {
              newStyle += `width: ${numericWidth}px; max-width: 100%; height: auto; `;
              img.setAttribute('width', String(numericWidth));
              img.removeAttribute('height'); // Để Word tự động scale tỉ lệ chiều cao
            } else {
              newStyle += 'width: 100%; max-width: 100%; height: auto; ';
              img.setAttribute('width', '100%');
              img.removeAttribute('height');
            }
            
            // Gộp style mới với style gốc một cách mượt mà (bỏ qua 'width' và 'height' cũ nếu có)
            let cleanedOriginalStyle = originalStyle
              .replace(/width\s*:\s*[^;]+;?/gi, '')
              .replace(/height\s*:\s*[^;]+;?/gi, '')
              .trim();
              
            img.setAttribute('style', `${newStyle} ${cleanedOriginalStyle}`.trim());
          });
          cleanedRichText = doc.body.innerHTML;
        } catch (e) {
          console.error("DOMParser clean failed, falling back to regex:", e);
          cleanedRichText = cleanedRichText.replace(/<img\s+/gi, '<img style="max-width: 100%; height: auto; display: inline-block; vertical-align: middle; margin: 4.5pt 6pt;" ');
        }
        contentHtml += `<div class="rich-text" style="margin: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;">${cleanedRichText}</div>`;
      }

      if (mat.tables && mat.tables.length > 0) {
        let tableOffset = 0;
        mat.tables.forEach((table, tIdx) => {
          const uniqueTableId = table.id || String(tIdx);
          const isTableHidden = (packageHiddenTables[mat.id] || []).includes(uniqueTableId);
          if (isTableHidden) return;
          
          const { data: filteredData, merges: filteredMerges } = getFilteredTableData(table, mat.id);
          if (filteredData && filteredData.length > 0) {
            // II.{mat index}.2, 3, etc.
            const sectionIdx = mat.richText ? tableOffset + 2 : tableOffset + 1;
            tableOffset++;
            const displayTitle = (table.title || 'Bảng thông số').trim();
            const titleWithColon = displayTitle.endsWith(':') ? displayTitle : `${displayTitle}:`;
            contentHtml += `<p style="margin: 0pt; padding: 0pt; mso-margin-top-alt: 0pt; mso-margin-bottom-alt: 0pt;"><strong>II.${index + 1}.${sectionIdx}. ${titleWithColon}</strong></p>`;
            contentHtml += `<table border="1" cellpadding="0" cellspacing="0" style="margin: 0pt; mso-table-lspace: 0pt; mso-table-rspace: 0pt;">`;
            filteredData.forEach((row, rIdx) => {
              if (rIdx === 0) {
                contentHtml += `<thead>`;
              } else if (rIdx === 1) {
                contentHtml += `<tbody>`;
              }
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
              if (rIdx === 0) {
                contentHtml += `</thead>`;
              }
            });
            if (filteredData.length > 1) {
              contentHtml += `</tbody>`;
            }
            contentHtml += `</table>`;
          }
        });
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

      if (mat.tables && mat.tables.length > 0) {
        let hasExportedTable = false;
        mat.tables.forEach((table, tIdx) => {
          const uniqueTableId = table.id || String(tIdx);
          const isTableHidden = (packageHiddenTables[mat.id] || []).includes(uniqueTableId);
          if (isTableHidden) return;
          
          const { data: filteredData, merges: filteredMerges } = getFilteredTableData(table, mat.id);

          if (filteredData && filteredData.length > 0) {
             hasExportedTable = true;
             const tableTitleCell = worksheet.getCell(currentRow, 1);
             tableTitleCell.value = table.title;
             tableTitleCell.font = { italic: true, bold: true, size: 12 };
             currentRow += 1;

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
                            tl: { col: colIdx + 0.05, row: currentRow + rowIdx - 1 + 0.05 } as any,
                            br: { col: colIdx + 1 - 0.05, row: currentRow + rowIdx - 0.05 } as any,
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
             
             currentRow += filteredData.length + 1; // 1 blank row between tables
          }
        });
        if (!hasExportedTable) {
           const cell = worksheet.getCell(currentRow, 1);
           cell.value = "Không có bảng thông số";
           cell.font = { italic: true };
           currentRow += 2;
        }
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
                 if (mat.tables) {
                   mat.tables.forEach(table => {
                     if (table.tags) {
                       Object.values(table.tags).forEach(tagStr => {
                         const tags = tagStr.split(',').map((t: string) => t.trim()).filter(Boolean);
                         tags.forEach((t: string) => tagsSet.add(t));
                       });
                     }
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

                 const hiddenTablesForMat = packageHiddenTables[mat.id] || [];
                 const toggleTable = (matId: string, tableId: string, checked: boolean) => {
                   setPackageHiddenTables(prev => {
                     const currentHidden = prev[matId] || [];
                     if (checked) {
                       return { ...prev, [matId]: currentHidden.filter(id => id !== tableId) };
                     } else {
                       return { ...prev, [matId]: [...currentHidden, tableId] };
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
                     
                     {mat.tables?.length > 0 && (
                       <div className="ml-14 border-t border-gray-100 pt-2 mt-1">
                         <span className="text-xs text-gray-500 font-medium mb-2 block">Bảng thông số & Tham biến hiển thị:</span>
                         {mat.tables.map((table, tIdx) => {
                            const uniqueTableId = table.id || String(tIdx);
                            const isTableHidden = hiddenTablesForMat.includes(uniqueTableId);
                            
                            const tableTagsSet = new Set<string>();
                            if (table.tags) {
                               Object.values(table.tags).forEach(tagStr => {
                                   const tags = tagStr.split(',').map((t: string) => t.trim()).filter(Boolean);
                                   tags.forEach((t: string) => tableTagsSet.add(t));
                               });
                            }
                            const uniqueTableTags = Array.from(tableTagsSet);

                            return (
                              <div key={uniqueTableId} className="mb-2 pb-2 border-b border-gray-50 last:mb-0 last:pb-0 last:border-0 border-dashed">
                                <label className="flex items-center gap-2 cursor-pointer mb-1.5 w-fit hover:bg-gray-50 px-1 py-0.5 rounded transition-colors -ml-1">
                                  <input
                                     type="checkbox"
                                     className="rounded border-gray-300 text-blue-600 focus:ring-blue-500 w-3.5 h-3.5 cursor-pointer"
                                     checked={!isTableHidden}
                                     onChange={(e) => toggleTable(mat.id, uniqueTableId, e.target.checked)}
                                  />
                                  <span className={`text-sm font-medium ${isTableHidden ? 'text-gray-400 line-through opacity-70' : 'text-gray-700'}`}>
                                    {table.title || 'Bảng thông số'}
                                  </span>
                                </label>
                                
                                {uniqueTableTags.length > 0 && !isTableHidden && (
                                  <div className="flex flex-wrap gap-2 ml-5">
                                    {uniqueTableTags.map(tag => {
                                       const isHidden = hiddenTagsForMat.includes(tag);
                                       return (
                                         <label key={tag} className={`flex items-center gap-1 text-xs px-2 py-1 rounded cursor-pointer border transition-colors ${isHidden ? 'bg-gray-50 text-gray-400 border-gray-200' : 'bg-blue-50 text-blue-700 border-blue-200'}`}>
                                           <input 
                                             type="checkbox" 
                                             className="hidden"
                                             checked={!isHidden} 
                                             onChange={() => toggleTag(mat.id, tag)} 
                                           />
                                           <div className={`w-3 h-3 flex items-center justify-center shrink-0 rounded-[2px] border ${isHidden ? 'bg-white border-gray-300 text-transparent' : 'bg-blue-600 border-blue-600 text-white'}`}>
                                             <CheckSquare size={12} className={isHidden ? 'opacity-0' : 'opacity-100'} />
                                           </div>
                                           <span>{tag}</span>
                                         </label>
                                       );
                                    })}
                                  </div>
                                )}
                              </div>
                            );
                         })}
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

