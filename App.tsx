/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Save, ClipboardList, Database, Info, FileText, Image as ImageIcon, AlignLeft, Plus, Edit2, Trash2, List, FileArchive, Search, X, User, LogOut, Key, Loader2 } from 'lucide-react';
import { RichTextEditor } from './components/RichTextEditor';
import { ExcelTable } from './components/ExcelTable';
import { ImageUpload } from './components/ImageUpload';
import { PackageBuilder } from './components/PackageBuilder';
import { Material, BiddingPackage } from './types';
import * as xlsx from 'xlsx';
import { auth, googleAuthProvider } from './lib/firebase.ts';
import { signInWithPopup, signOut, onAuthStateChanged, User as FirebaseUser, GoogleAuthProvider } from 'firebase/auth';

export default function App() {
  const [savedMaterials, setSavedMaterials] = useState<Material[]>([]);
  const [savedPackages, setSavedPackages] = useState<BiddingPackage[]>([]);
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);

  const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || '';

  const fetchMaterials = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/api/materials`);
      if (res.ok) {
        const data = await res.json();
        const mapped = data.map((d: any) => ({
          id: d.id,
          code: d.code,
          name: d.name,
          richText: d.content,
          excelData: d.excelData?.data || [],
          excelMerges: d.excelData?.merges || [],
          rowTags: d.excelData?.tags || {},
          images: d.images || [],
          notes: d.notes || '',
          updatedBy: d.updatedBy,
          updatedAt: d.updatedAt
        }));
        setSavedMaterials(mapped);
      }
    } catch (e) {
      console.error(e);
    }
  };

  const fetchPackages = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/api/packages`);
      if (res.ok) {
        const data = await res.json();
        setSavedPackages(data);
      }
    } catch (e) {
      console.error(e);
    }
  };

  useEffect(() => {
    fetchMaterials();
    fetchPackages();
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      const allowedEmails = import.meta.env.VITE_ALLOWED_EMAILS;
      if (currentUser && currentUser.email && allowedEmails && allowedEmails.trim() !== '') {
        const emailList = allowedEmails.split(',').map((e: string) => e.trim().toLowerCase());
        if (!emailList.includes(currentUser.email.toLowerCase())) {
          alert('Tài khoản Google của bạn không có quyền truy cập.');
          await signOut(auth);
          setUser(null);
          setIsLoading(false);
          return;
        }
      }
      setUser(currentUser);
      setIsLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleAuthProvider);
    } catch (error: any) {
      console.error("Login failed:", error);
      alert(`Đăng nhập thất bại: ${error.message || error}`);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Logout failed:", error);
    }
  };

  const isAuthenticated = !!user;

  const [activeTab, setActiveTab] = useState<'materials' | 'package'>('materials');
  
  // Current Form State
  const [isEditing, setIsEditing] = useState(true);
  const [currentId, setCurrentId] = useState<string>(() => crypto.randomUUID());
  const [code, setCode] = useState('');
  const [name, setName] = useState('');
  const [richText, setRichText] = useState('');
  const [excelData, setExcelData] = useState<string[][]>([]);
  const [excelMerges, setExcelMerges] = useState<xlsx.Range[]>([]);
  const [rowTags, setRowTags] = useState<Record<number, string>>({});
  const [images, setImages] = useState<{ url: string; name: string }[]>([]);
  const [notes, setNotes] = useState('');

  // Validation
  const [codeError, setCodeError] = useState('');
  const [nameError, setNameError] = useState('');

  // Sidebar
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [searchSidebar, setSearchSidebar] = useState('');
  const [showLoginModal, setShowLoginModal] = useState(false);
  
  // Audit Logs
  const [showAuditLogs, setShowAuditLogs] = useState(false);
  const [auditLogsData, setAuditLogsData] = useState<any[]>([]);
  const [loadingAuditLogs, setLoadingAuditLogs] = useState(false);

  const requireAuth = (callback: () => void) => {
    if (isAuthenticated) {
      callback();
    } else {
      setShowLoginModal(true);
    }
  };


  const resetForm = () => {
    requireAuth(() => {
      setCurrentId(crypto.randomUUID());
      setCode('');
      setName('');
      setRichText('');
      setExcelData([]);
      setExcelMerges([]);
      setRowTags({});
      setImages([]);
      setNotes('');
      setIsEditing(true);
      setCodeError('');
      setNameError('');
      setActiveTab('materials');
    });
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isEditing) {
      requireAuth(() => setIsEditing(true));
      return;
    }

    if (!name.trim()) {
      alert("Vui lòng nhập tên vật tư thiết bị!");
      return;
    }

    // Validation
    let hasError = false;
    setCodeError('');
    setNameError('');

    if (code.trim()) {
      const dupCode = savedMaterials.find(m => m.code.toLowerCase() === code.trim().toLowerCase() && m.id !== currentId);
      if (dupCode) {
        setCodeError('Đã tồn tại mã vật tư này');
        hasError = true;
      }
    }
    
    if (name.trim()) {
      const dupName = savedMaterials.find(m => m.name.toLowerCase() === name.trim().toLowerCase() && m.id !== currentId);
      if (dupName) {
        setNameError('Đã tồn tại tên vật tư thiết bị này');
        hasError = true;
      }
    }

    if (hasError) return;

    if (!user) {
      setShowLoginModal(true);
      return;
    }

    const newMaterial = {
      id: currentId,
      code: code.trim(),
      name: name.trim(),
      content: richText,
      excelData: { data: excelData, merges: excelMerges, tags: rowTags },
      images,
      notes,
    };
    
    setIsSaving(true);
    try {
      const token = await user.getIdToken();
      const res = await fetch(`${API_BASE_URL}/api/materials`, {
         method: 'POST',
         headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
         },
         body: JSON.stringify(newMaterial)
      });
      if (!res.ok) {
        const errorData = await res.json().catch(() => null);
        throw new Error(errorData?.details || errorData?.error || 'Network error');
      }
      
      const resData = await res.json();
      
      setSavedMaterials(prev => {
        const existingIdx = prev.findIndex(m => m.id === currentId);
        // Note: material from backend differs, mapping to frontend properties
        const fetchedMat: Material = {
            id: resData.id,
            code: resData.code,
            name: resData.name,
            richText: resData.content,
            excelData: resData.excelData?.data || [],
            excelMerges: resData.excelData?.merges || [],
            rowTags: resData.excelData?.tags || {},
            images: resData.images || [],
            notes: resData.notes || '',
            updatedBy: resData.updatedBy,
            updatedAt: resData.updatedAt
        };
        if (existingIdx >= 0) {
          const copy = [...prev];
          copy[existingIdx] = fetchedMat;
          return copy;
        }
        return [...prev, fetchedMat];
      });

      setIsEditing(false); // Switch to view mode after saving
      alert("Dữ liệu vật tư đã được lưu!");
    } catch (e: any) {
      console.error(e);
      alert(`Đã xảy ra lỗi khi lưu vật tư: ${e.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancelEdit = () => {
    const existing = savedMaterials.find(m => m.id === currentId);
    if (existing) {
      loadMaterialForView(existing);
    } else {
      setCode('');
      setName('');
      setRichText('');
      setExcelData([]);
      setExcelMerges([]);
      setRowTags({});
      setImages([]);
      setNotes('');
      setCodeError('');
      setNameError('');
      setIsEditing(false);
    }
  };

  const loadMaterialForView = (material: Material) => {
    setCurrentId(material.id);
    setCode(material.code);
    setName(material.name);
    setRichText(material.richText);
    setExcelData(material.excelData);
    setExcelMerges(material.excelMerges || []);
    setRowTags(material.rowTags || {});
    setImages(material.images || []);
    setNotes(material.notes);
    setCodeError('');
    setNameError('');
    setIsEditing(false); // Open in view mode
    setActiveTab('materials');
  };

  const loadMaterialForEdit = (material: Material, e: React.MouseEvent) => {
    e.stopPropagation();
    requireAuth(() => {
      loadMaterialForView(material);
      setIsEditing(true);
    });
  };

  const [deleteMaterialId, setDeleteMaterialId] = useState<string | null>(null);

  const deleteMaterial = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    requireAuth(() => {
      setDeleteMaterialId(id);
    });
  };

  const confirmDeleteMaterial = async () => {
    if (deleteMaterialId && user) {
      try {
        const token = await user.getIdToken();
        const res = await fetch(`${API_BASE_URL}/api/materials/${deleteMaterialId}`, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${token}`
          }
        });
        if (!res.ok) throw new Error('Network error');
        setSavedMaterials(prev => prev.filter(m => m.id !== deleteMaterialId));
        if (currentId === deleteMaterialId) {
          resetForm();
        }
        setDeleteMaterialId(null);
      } catch (error) {
        console.error(error);
        alert('Lỗi xóa vật tư');
      }
    }
  };

  const handleOpenAuditLogs = async (matId: string, e?: React.MouseEvent) => {
     if(e) e.stopPropagation();
     setShowAuditLogs(true);
     setLoadingAuditLogs(true);
     try {
       const res = await fetch(`${API_BASE_URL}/api/materials/${matId}/audit-logs`);
       if(res.ok) {
         setAuditLogsData(await res.json());
       } else {
         setAuditLogsData([]);
       }
     } catch (err) {
       console.error("Failed to load audit logs", err);
       setAuditLogsData([]);
     } finally {
       setLoadingAuditLogs(false);
     }
  };

  // Sort and filter alphabetically
  const filteredMaterials = savedMaterials
    .filter(m => {
       const search = searchSidebar.toLowerCase();
       return (m.name || '').toLowerCase().includes(search) || (m.code || '').toLowerCase().includes(search);
    })
    .sort((a, b) => (a.name || '').localeCompare(b.name || ''));

  return (
    <div className="min-h-screen bg-gray-100 text-gray-900 font-sans selection:bg-blue-100 selection:text-blue-900 flex flex-col md:flex-row">
      {/* Sidebar - Saved Items */}
      <aside className={`${isSidebarOpen ? 'w-full md:w-80' : 'w-0'} transition-all duration-300 bg-white border-r border-gray-200 shadow-sm flex flex-col z-40 overflow-hidden shrink-0`}>
        <div className="p-4 border-b border-gray-100 bg-gray-50 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2 text-gray-800 font-semibold">
            <List size={18} className="text-blue-600" />
            <span>Đã lưu ({savedMaterials.length})</span>
          </div>
          <button 
            onClick={resetForm}
            className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-md transition-colors"
            title="Thêm Vật Tư Mới"
          >
            <Plus size={18} />
          </button>
        </div>

        <div className="p-3 border-b border-gray-200 shrink-0">
           <div className="relative text-sm">
             <Search className="absolute left-2.5 top-2.5 text-gray-400" size={16} />
             <input
               type="text"
               placeholder="Tìm mã hoặc tên..."
               value={searchSidebar}
               onChange={e => setSearchSidebar(e.target.value)}
               className="w-full pl-9 pr-8 py-2 rounded-md border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none block"
             />
             {searchSidebar && (
               <button onClick={() => setSearchSidebar('')} className="absolute right-2 top-2 text-gray-400 hover:text-gray-600 p-0.5">
                 <X size={16}/>
               </button>
             )}
           </div>
        </div>

        <div className="overflow-y-auto flex-1 p-2 space-y-2">
          {filteredMaterials.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-gray-500">
              {searchSidebar ? "Không tìm thấy vật tư." : "Chưa có vật tư nào được lưu."}
            </div>
          ) : (
            filteredMaterials.map(mat => (
              <div 
                key={mat.id}
                onClick={() => loadMaterialForView(mat)}
                className={`p-3 rounded-lg border cursor-pointer transition-all ${
                  currentId === mat.id && activeTab === 'materials'
                    ? 'border-blue-500 bg-blue-50 shadow-sm ring-1 ring-blue-500/20' 
                    : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50'
                }`}
              >
                <div className="flex justify-between items-start gap-2">
                  <div className="overflow-hidden flex-1">
                    <h3 className="font-semibold text-gray-800 text-sm truncate" title={mat.name}>{mat.name}</h3>
                    <p className="text-xs text-gray-500 mt-1 truncate" title={mat.code}>{mat.code || 'Chưa có mã'}</p>
                  </div>
                  {isAuthenticated && (
                    <div className="flex flex-col gap-1 shrink-0">
                      <button 
                        onClick={(e) => loadMaterialForEdit(mat, e)}
                        className="flex items-center justify-center gap-1 px-2 py-1 flex-1 text-xs font-medium text-emerald-600 hover:bg-emerald-100 bg-emerald-50 rounded transition-colors"
                        title="Chỉnh sửa"
                      >
                        <Edit2 size={12} />
                      </button>
                      <button 
                        onClick={(e) => deleteMaterial(mat.id, e)}
                        className="p-1 flex-1 flex items-center justify-center text-gray-400 hover:text-red-500 hover:bg-red-50 rounded transition-colors self-end"
                        title="Xóa"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  )}
                </div>
                {deleteMaterialId === mat.id && (
                  <div className="mt-2 text-xs text-red-600 bg-red-50 p-2 rounded flex items-center justify-between" onClick={e => e.stopPropagation()}>
                    <span>Bạn có chắc xóa?</span>
                    <div className="flex gap-2">
                       <button onClick={(e) => { e.stopPropagation(); setDeleteMaterialId(null); }} className="hover:underline text-gray-500">Hủy</button>
                       <button onClick={(e) => { e.stopPropagation(); confirmDeleteMaterial(); }} className="hover:underline font-bold text-red-600">Xác nhận</button>
                    </div>
                  </div>
                )}
              </div>
            ))
          )}
        </div>


      </aside>

      {/* Main Content */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header className="bg-white border-b border-gray-200 sticky top-0 z-50 shadow-sm shrink-0">
          <div className="px-4 sm:px-6 h-16 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <button 
                onClick={() => setIsSidebarOpen(!isSidebarOpen)}
                className="p-2 text-gray-500 hover:bg-gray-100 rounded-md md:hidden"
              >
                <List size={20} />
              </button>
              <div className="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center text-white shadow-inner hidden sm:flex shrink-0">
                <ClipboardList size={22} />
              </div>
              <div className="hidden sm:block">
                <h1 className="text-lg sm:text-xl font-bold tracking-tight text-gray-900 leading-none truncate pr-2">Tiêu chuẩn Kỹ thuật</h1>
              </div>
            </div>
            <div className="flex flex-col sm:flex-row items-center gap-2 shrink-0">
              
              {/* Tab Navigation */}
              <div className="flex bg-gray-100 p-1 rounded-lg">
                <button
                  onClick={() => setActiveTab('materials')}
                  className={`flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md font-medium text-sm transition-colors ${
                    activeTab === 'materials' 
                      ? 'bg-white text-blue-600 shadow-sm ring-1 ring-black/5' 
                      : 'text-gray-500 hover:text-gray-700 hover:bg-gray-200/50'
                  }`}
                >
                  <ClipboardList size={16} />
                  <span className="hidden sm:inline">Quản lý</span>
                </button>
                <button
                  onClick={() => setActiveTab('package')}
                  className={`flex items-center justify-center gap-2 px-3 sm:px-4 py-2 rounded-md font-medium text-sm transition-colors ${
                    activeTab === 'package' 
                      ? 'bg-white text-blue-600 shadow-sm ring-1 ring-black/5' 
                      : 'text-gray-500 hover:text-gray-700 hover:bg-gray-200/50'
                  }`}
                >
                  <FileArchive size={16} />
                  <span className="hidden sm:inline">Gói thầu</span>
                </button>
              </div>

            </div>
            
            {activeTab === 'materials' && isAuthenticated && (
              <div className="flex items-center gap-2 shrink-0 ml-4">
                 <button 
                    onClick={resetForm}
                    type="button"
                    className="hidden lg:flex items-center gap-2 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-lg text-sm font-semibold transition-all shadow-sm focus:ring-2 focus:ring-gray-200 focus:outline-none"
                  >
                    <Plus size={16} />
                    Thêm mới
                  </button>
                <button 
                  type="button"
                  onClick={handleSave}
                  disabled={isSaving}
                  className={`flex items-center gap-2 px-4 sm:px-5 py-2 rounded-lg text-sm font-semibold transition-all shadow hover:shadow-md ring-2 ring-transparent focus:ring-offset-2 focus:outline-none ${
                    isEditing 
                      ? 'bg-blue-600 hover:bg-blue-700 text-white focus:ring-blue-500 disabled:opacity-75 disabled:cursor-not-allowed'
                      : 'bg-emerald-600 hover:bg-emerald-700 text-white focus:ring-emerald-500'
                  }`}
                >
                  {isSaving ? (
                    <><Loader2 size={18} className="animate-spin" /> Đang lưu...</>
                  ) : isEditing ? (
                    <><Save size={18} /> Lưu cập nhật</>
                  ) : (
                    <><Edit2 size={18} /> Chỉnh sửa</>
                  )}
                </button>
                {isEditing && (
                  <button 
                    type="button"
                    onClick={handleCancelEdit}
                    className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold text-gray-700 bg-white border border-gray-300 hover:bg-gray-50 transition-all shadow-sm focus:ring-2 focus:ring-gray-200 focus:outline-none"
                  >
                    <X size={18} /> Hủy
                  </button>
                )}
              </div>
            )}
            
            <div className="flex items-center gap-2 shrink-0 ml-4 border-l border-gray-200 pl-4 h-8">
              {isAuthenticated ? (
                  <button
                    onClick={handleLogout}
                    className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-md transition-colors"
                    title="Đăng xuất"
                  >
                    <LogOut size={16} />
                  </button>
              ) : (
                <button
                  onClick={() => setShowLoginModal(true)}
                  className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-md transition-colors"
                >
                  <User size={16} />
                  <span className="hidden sm:inline">Đăng nhập</span>
                </button>
              )}
            </div>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto w-full p-4 sm:p-6 lg:p-8 bg-gray-50/50">
          {activeTab === 'package' ? (
            <PackageBuilder 
               savedMaterials={savedMaterials} 
               savedPackages={savedPackages}
               setSavedPackages={setSavedPackages}
               isAuthenticated={isAuthenticated}
               requireAuth={requireAuth}
            />
          ) : (
            <div className="w-full">
              {!isEditing && (
                 <div className="bg-yellow-50 border border-yellow-200 text-yellow-800 px-4 py-3 rounded-lg mb-6 flex items-center gap-3">
                   <Info size={20} className="text-yellow-600 shrink-0"/>
                   <div>
                     <p className="text-sm font-medium">Chế độ xem trước.</p>
                     <p className="text-xs mt-0.5">Vật tư đang bị khóa chỉnh sửa. Vui lòng bấm nút <strong>"Chỉnh sửa"</strong> phía trên cùng để thay đổi dữ liệu.</p>
                   </div>
                 </div>
              )}

              <form id="material-form" onSubmit={handleSave} className="space-y-6">
                
                {/* Thông tin cơ bản */}
                <section className={`bg-white rounded-xl shadow-sm border overflow-hidden ${!isEditing ? 'border-gray-200 opacity-95' : 'border-gray-200'}`}>
                  <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center justify-between">
                    <div className="flex items-center gap-2">
                       <Info className="text-blue-500" size={20} />
                       <h2 className="text-lg font-semibold text-gray-800">Thông tin cơ bản</h2>
                    </div>
                    {savedMaterials.find(m => m.id === currentId)?.updatedAt && (
                      <div className="flex items-center gap-2 text-sm text-gray-500">
                        <span>Cập nhật lần cuối bởi <strong>{savedMaterials.find(m => m.id === currentId)?.updatedBy}</strong> lúc {new Date(savedMaterials.find(m => m.id === currentId)!.updatedAt!).toLocaleString('vi-VN')}</span>
                        <button type="button" onClick={() => handleOpenAuditLogs(currentId)} className="text-blue-600 hover:text-blue-700 underline text-xs font-semibold ml-2">Xem lịch sử</button>
                      </div>
                    )}
                  </div>
                  <div className="p-6 grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <label htmlFor="materialCode" className="block text-sm font-semibold text-gray-700">
                        Mã vật tư
                      </label>
                      <input
                        type="text"
                        id="materialCode"
                        value={code}
                        readOnly={!isEditing}
                        onChange={e => {
                          setCode(e.target.value);
                          if (codeError) setCodeError('');
                        }}
                        placeholder="VD: VT-D-10293"
                        className={`w-full px-4 py-2.5 rounded-lg border focus:ring-2 outline-none transition-shadow text-gray-900 ${
                          codeError ? 'border-red-500 focus:ring-red-500' : 'border-gray-300 focus:ring-blue-500 focus:border-blue-500'
                        } ${!isEditing ? 'bg-gray-50 border-transparent shadow-none px-0 focus:ring-0 font-medium' : ''}`}
                      />
                      {codeError && <p className="text-sm font-medium text-red-500">{codeError}</p>}
                    </div>
                    <div className="space-y-2">
                      <label htmlFor="materialName" className="block text-sm font-semibold text-gray-700">
                        Tên vật tư thiết bị {isEditing && <span className="text-red-500">*</span>}
                      </label>
                      <input
                        type="text"
                        id="materialName"
                        required={isEditing}
                        readOnly={!isEditing}
                        value={name}
                        onChange={e => {
                          setName(e.target.value);
                          if (nameError) setNameError('');
                        }}
                        placeholder="VD: Cáp điện hạ thế CXV 4x25"
                        className={`w-full px-4 py-2.5 rounded-lg border focus:ring-2 outline-none transition-shadow text-gray-900 ${
                          nameError ? 'border-red-500 focus:ring-red-500' : 'border-gray-300 focus:ring-blue-500 focus:border-blue-500'
                        } ${!isEditing ? 'bg-gray-50 border-transparent shadow-none px-0 focus:ring-0 font-medium' : ''}`}
                      />
                      {nameError && <p className="text-sm font-medium text-red-500">{nameError}</p>}
                    </div>
                  </div>
                </section>

                {/* Hình ảnh */}
                {(isEditing || images.length > 0) && (
                  <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                    <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center gap-2">
                      <ImageIcon className="text-blue-500" size={20} />
                      <h2 className="text-lg font-semibold text-gray-800">Hình ảnh minh họa</h2>
                    </div>
                    <div className="p-6">
                      <ImageUpload images={images} setImages={setImages} readOnly={!isEditing} />
                    </div>
                  </section>
                )}

                {/* Yêu cầu chung */}
                {(isEditing || (richText && richText !== '<p><br></p>' && (richText.includes('<img') || richText.replace(/<[^>]*>?/gm, '').trim() !== ''))) && (
                  <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                    <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center gap-2">
                      <AlignLeft className="text-blue-500" size={20} />
                      <h2 className="text-lg font-semibold text-gray-800">Yêu cầu chung</h2>
                    </div>
                    <div className="p-6">
                      <RichTextEditor value={richText} onChange={setRichText} readOnly={!isEditing} />
                    </div>
                  </section>
                )}

                {/* Bảng thông số kỹ thuật */}
                <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                  <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center gap-2">
                    <Database className="text-blue-500" size={20} />
                    <h2 className="text-lg font-semibold text-gray-800">Bảng thông số kỹ thuật</h2>
                  </div>
                  <div className="p-6 overflow-hidden">
                    {isEditing && (
                      <p className="text-sm text-gray-500 mb-4 items-center">
                        Cấu hình bảng hoặc tải lên file Excel để trích xuất tự động (Giữ nguyên gộp ô và có thể kéo chiều dài/rộng ô).
                      </p>
                    )}
                    <ExcelTable 
                      tableData={excelData} 
                      setTableData={setExcelData} 
                      merges={excelMerges} 
                      setMerges={setExcelMerges}
                      rowTags={rowTags}
                      setRowTags={setRowTags}
                      readOnly={!isEditing}
                    />
                  </div>
                </section>

                {/* Ghi chú */}
                <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                  <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center gap-2">
                    <FileText className="text-blue-500" size={20} />
                    <h2 className="text-lg font-semibold text-gray-800">Ghi chú</h2>
                  </div>
                  <div className="p-6">
                    <textarea
                      id="notes"
                      rows={4}
                      readOnly={!isEditing}
                      value={notes}
                      onChange={e => setNotes(e.target.value)}
                      placeholder="Nhập các ghi chú kỹ thuật, tiêu chuẩn áp dụng, nhà sản xuất gợi ý..."
                      className={`w-full px-4 py-3 rounded-lg border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-shadow text-gray-900 resize-y ${!isEditing ? 'bg-gray-50 border-transparent shadow-none px-0 py-0 focus:ring-0 resize-none font-medium' : ''}`}
                    ></textarea>
                  </div>
                </section>

              </form>
            </div>
          )}
        </main>
      </div>

      {/* Login Modal */}
      {showLoginModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-6">
              <h3 className="text-xl font-bold text-gray-900 mb-4">Đăng nhập tài khoản</h3>
              <p className="text-sm text-gray-500 mb-4">Sử dụng Google để đăng nhập và có quyền quản lý, chỉnh sửa.</p>
              
              <div className="flex justify-center mt-6">
                <button
                  onClick={() => {
                    setShowLoginModal(false);
                    handleLogin();
                  }}
                  className="px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-sm transition-colors w-full"
                >
                  Đăng nhập bằng Google
                </button>
              </div>
              <div className="flex justify-center mt-4">
                  <button
                    onClick={() => setShowLoginModal(false)}
                    className="px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-100 rounded-lg transition-colors w-full"
                  >
                    Đóng
                  </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Remove Password Change Modal as it is handled by Google */}
      
      {/* Audit Logs Modal */}
      {showAuditLogs && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50">
              <h3 className="text-lg font-bold text-gray-900">Lịch sử thay đổi vật tư</h3>
              <button onClick={() => setShowAuditLogs(false)} className="text-gray-400 hover:text-gray-600 transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-6 flex-1 overflow-y-auto bg-gray-50/50">
              {loadingAuditLogs ? (
                 <div className="flex justify-center p-8 text-gray-500">Đang tải dữ liệu...</div>
              ) : auditLogsData.length === 0 ? (
                 <div className="text-center p-8 text-gray-500">Không có dữ liệu lịch sử nào.</div>
              ) : (
                <div className="space-y-4">
                  {auditLogsData.map((log) => (
                     <div key={log.id} className="bg-white rounded-lg border border-gray-200 p-4 shadow-sm">
                       <div className="flex justify-between items-start mb-2">
                         <div className="flex gap-2 items-center">
                            <span className={`text-xs font-bold uppercase px-2 py-1 rounded ${log.action === 'create' ? 'bg-emerald-100 text-emerald-700' : log.action === 'update' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                               {log.action === 'create' ? 'Tạo mới' : log.action === 'update' ? 'Cập nhật' : 'Xóa'}
                            </span>
                            <span className="text-sm font-semibold text-gray-800">{log.updatedBy || 'Người dùng không xác định'}</span>
                         </div>
                         <span className="text-sm text-gray-500">{new Date(log.updatedAt).toLocaleString('vi-VN')}</span>
                       </div>
                       <div className="text-sm text-gray-600 mt-2 bg-gray-50 p-2 rounded max-h-40 overflow-y-auto font-mono text-xs border border-gray-100">
                          <strong>Dữ liệu mới:</strong>
                          <pre className="mt-1 whitespace-pre-wrap">{JSON.stringify(log.newData, null, 2)}</pre>
                       </div>
                     </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

    </div>
  );
}


