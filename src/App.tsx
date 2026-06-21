/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Save, ClipboardList, Database, Info, FileText, Image as ImageIcon, AlignLeft, Plus, Edit2, Trash2, List, FileArchive, Search, X, User, LogOut, Key, Loader2, LayoutDashboard, Settings, Bell, MessageSquare, ChevronDown, Menu } from 'lucide-react';
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
        const mapped = data.map((d: any) => {
          let tables = d.excelData?.tables || [];
          if (tables.length === 0 && d.excelData?.data) {
            tables = [{
              id: 'legacy-table',
              title: 'Bảng thông số kỹ thuật',
              data: d.excelData.data || [],
              merges: d.excelData.merges || [],
              tags: d.excelData.tags || {}
            }];
          }
          return {
            id: d.id,
            code: d.code,
            name: d.name,
            richText: d.content,
            tables: tables,
            images: d.images || [],
            notes: d.notes || '',
            updatedBy: d.updatedBy,
            updatedAt: d.updatedAt
          };
        });
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
  const [tables, setTables] = useState<import('./types').ExcelTableConfig[]>([{
     id: crypto.randomUUID(),
     title: 'Bảng thông số kỹ thuật',
    data: [
      ['STT', 'Hạng mục', 'Đơn vị', 'Yêu cầu', 'Thông số chào'],
      ['1', 'Nhà sản xuất', '', 'Nêu cụ thể', ''],
      ['2', 'Nước sản xuất', '', 'Nêu cụ thể', '']
    ],
    merges: [],
    tags: {}
  }]);
  const [images, setImages] = useState<{ url: string; name: string }[]>([]);
  const [notes, setNotes] = useState('');
  const [docRequirements, setDocRequirements] = useState({
    typeTest: false,
    catalog: false,
    endUser: false,
    iso: false,
  });

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
      setTables([{
         id: crypto.randomUUID(),
         title: 'Bảng thông số kỹ thuật',
         data: [
           ['STT', 'Hạng mục', 'Đơn vị', 'Yêu cầu', 'Thông số chào'],
           ['1', 'Nhà sản xuất', '', 'Nêu cụ thể', ''],
           ['2', 'Nước sản xuất', '', 'Nêu cụ thể', '']
         ],
         merges: [],
         tags: {}
      }]);
      setImages([]);
      setNotes('');
      setDocRequirements({
        typeTest: false,
        catalog: false,
        endUser: false,
        iso: false,
      });
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
      excelData: { tables },
      images,
      notes,
      docRequirements,
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
        
        // Handle migration inside
        let fetchedTables = resData.excelData?.tables || [];
        if (fetchedTables.length === 0 && resData.excelData?.data) {
           fetchedTables = [{
              id: 'legacy-table',
              title: 'Bảng thông số kỹ thuật',
              data: resData.excelData.data || [],
              merges: resData.excelData.merges || [],
              tags: resData.excelData.tags || {}
           }];
        }

        const fetchedMat: Material = {
            id: resData.id,
            code: resData.code,
            name: resData.name,
            richText: resData.content,
            tables: fetchedTables,
            images: resData.images || [],
            notes: resData.notes || '',
            docRequirements: resData.docRequirements || {},
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
      setTables([{
         id: crypto.randomUUID(),
         title: 'Bảng thông số kỹ thuật',
         data: [
           ['STT', 'Hạng mục', 'Đơn vị', 'Yêu cầu', 'Thông số chào'],
           ['1', 'Nhà sản xuất', '', 'Nêu cụ thể', ''],
           ['2', 'Nước sản xuất', '', 'Nêu cụ thể', '']
         ],
         merges: [],
         tags: {}
      }]);
      setImages([]);
      setNotes('');
      setDocRequirements({
        typeTest: false,
        catalog: false,
        endUser: false,
        iso: false,
      });
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
    setTables(material.tables && material.tables.length > 0 ? material.tables : [{
         id: crypto.randomUUID(),
         title: 'Bảng thông số kỹ thuật',
         data: [
           ['STT', 'Hạng mục', 'Đơn vị', 'Yêu cầu', 'Thông số chào'],
           ['1', 'Nhà sản xuất', '', 'Nêu cụ thể', ''],
           ['2', 'Nước sản xuất', '', 'Nêu cụ thể', '']
         ],
         merges: [],
         tags: {}
    }]);
    setImages(material.images || []);
    setNotes(material.notes);
    setDocRequirements(material.docRequirements || {
      typeTest: false,
      catalog: false,
      endUser: false,
      iso: false,
    });
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
    <div className="h-screen w-full bg-[#f4f7fc] text-gray-800 font-sans selection:bg-indigo-100 selection:text-indigo-900 flex overflow-hidden">
        {/* NEW Main Navigation Sidebar */}
        <aside className="w-64 bg-white border-r border-gray-100 flex flex-col shrink-0 z-50 shadow-sm hidden md:flex">
           <div className="h-16 flex items-center justify-center px-6 border-b border-gray-100 bg-white">
              <div className="flex items-center gap-2">
                <div className="flex-shrink-0 w-11 h-11 rounded-full border-2 border-[#10498B] flex items-center justify-center bg-white shadow-sm overflow-hidden">
                  <div className="relative w-full h-full flex items-center justify-center">
                    <svg viewBox="0 0 100 100" className="w-[125%] h-[125%] absolute">
                      <path d="M 50 4 L 62 38 L 96 50 L 62 62 L 50 96 L 38 62 L 4 50 L 38 38 Z" fill="#10498B" />
                      <path d="M 50 22 L 57 43 L 78 50 L 57 57 L 50 78 L 43 57 L 22 50 L 43 43 Z" fill="#E31837" />
                      <path d="M 50 36 L 53.5 46.5 L 64 50 L 53.5 53.5 L 50 64 L 46.5 53.5 L 36 50 L 46.5 46.5 Z" fill="#FFED00" />
                    </svg>
                  </div>
                </div>
                <div className="flex flex-col justify-center ml-0.5">
                  <div className="flex items-baseline font-black leading-none gap-[1px]">
                    <span className="text-[#10498B] text-[24px] tracking-tight">EVN</span>
                    <span className="text-[#E31837] text-[24px] tracking-tight">CPC</span>
                  </div>
                  <span className="text-[#10498B] font-black italic text-[11px] leading-tight tracking-[0.06em] mt-[3px]">PC QUANG NGAI</span>
                </div>
              </div>
           </div>
           
           <div className="px-4 py-6">
             <div className="w-full bg-indigo-50 rounded-xl p-4 flex items-center justify-between mb-2">
                <div className="flex items-center gap-3">
                   <div className="bg-white p-2 rounded-lg shadow-sm text-indigo-600 border border-indigo-100">
                      <LayoutDashboard size={20} />
                   </div>
                   <div>
                      <p className="text-xs text-gray-500 font-medium">System</p>
                      <p className="text-sm font-bold text-indigo-900">Dashboard</p>
                   </div>
                </div>
                <ChevronDown size={16} className="text-indigo-400" />
             </div>
           </div>
           
           <nav className="flex-1 overflow-y-auto px-4 space-y-1.5 pb-6">
              <p className="text-[11px] font-bold text-gray-400 mb-3 px-2 uppercase tracking-wider">WORKSPACE</p>
              
              <button 
                 onClick={() => setActiveTab('materials')}
                 className={`w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${activeTab === 'materials' ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/20' : 'text-gray-600 hover:bg-gray-50 hover:text-gray-900'}`}
              >
                 <div className="flex items-center gap-3">
                    <Database size={18} className={activeTab === 'materials' ? 'text-white' : 'text-gray-400'} />
                    Quản lý Vật tư
                 </div>
                 {activeTab === 'materials' && <div className="w-1.5 h-1.5 rounded-full bg-white opacity-80 decoration-pulse" />}
              </button>
              
              <button 
                 onClick={() => setActiveTab('package')}
                 className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${activeTab === 'package' ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/20' : 'text-gray-600 hover:bg-gray-50 hover:text-gray-900'}`}
              >
                 <FileArchive size={18} className={activeTab === 'package' ? 'text-white' : 'text-gray-400'} />
                 Tạo Gói Thầu
              </button>
           </nav>
        </aside>

        {/* Main Interface */}
        <div className="flex-1 flex flex-col min-w-0 overflow-hidden relative">
            {/* Top Bar matching screenshot */}
            <header className="h-16 bg-gradient-to-r from-indigo-500 via-purple-500 to-pink-500 flex items-center justify-between px-4 sm:px-6 shrink-0 z-40 text-white shadow-sm">
              <div className="flex items-center gap-4">
                 <button className="p-1.5 hover:bg-white/10 rounded-lg md:hidden transition-colors">
                    <Menu size={20} />
                 </button>
                 <div className="hidden sm:block">
                    <p className="text-[11px] text-white/80 font-medium tracking-wide uppercase">Welcome back</p>
                    <h2 className="font-bold text-base leading-tight">{isAuthenticated ? auth.currentUser?.displayName || 'User' : 'Guest'}</h2>
                 </div>
              </div>
              <div className="flex items-center gap-3 lg:gap-5">
                 {isAuthenticated ? (
                     <div className="flex items-center gap-3 pl-2 lg:pl-4 border-l border-white/20">
                         <button onClick={handleLogout} className="flex items-center gap-3 tooltip-trigger hover:opacity-80 transition-opacity" title="Đăng xuất">
                           <img src={auth.currentUser?.photoURL || 'https://api.dicebear.com/7.x/avataaars/svg'} alt="avatar" className="w-9 h-9 rounded-xl bg-white/20 object-cover border border-white/20 shadow-sm" />
                           <div className="text-left hidden lg:block">
                             <p className="text-sm font-bold leading-tight line-clamp-1">{auth.currentUser?.displayName || 'User'}</p>
                             <p className="text-xs text-white/70 leading-tight">Admin</p>
                           </div>
                         </button>
                     </div>
                 ) : (
                     <div className="flex items-center gap-2 pl-2 lg:pl-4 border-l border-white/20">
                         <button
                           onClick={() => setShowLoginModal(true)}
                           className="flex items-center gap-2 px-4 py-1.5 bg-white text-indigo-600 hover:bg-white/90 rounded-full transition-colors text-sm font-bold shadow-sm"
                         >
                           <User size={16} />
                           <span>Login</span>
                         </button>
                     </div>
                 )}
              </div>
            </header>

            {/* Inner Content Workspace */}
            <main className="flex-1 flex overflow-hidden relative">
                
                {activeTab === 'materials' && (
                  <aside className={`${isSidebarOpen ? 'w-full md:w-80' : 'w-0'} transition-all duration-300 bg-white border-r border-gray-200 shadow-sm flex flex-col z-30 overflow-hidden shrink-0`}>
                    <div className="p-4 border-b border-gray-100 bg-gray-50 flex items-center justify-between shrink-0">
                      <div className="flex items-center gap-2 text-gray-800 font-semibold">
                        <Database size={18} className="text-indigo-600" />
                        <span>Danh Sách ({savedMaterials.length})</span>
                      </div>
                      <button 
                        onClick={resetForm}
                        className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded-md transition-colors"
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
                           className="w-full pl-9 pr-8 py-2 rounded-md border border-gray-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none block bg-gray-50"
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
                )}

      {/* Editor Content Workspace */}
      <div className={`flex-1 flex flex-col min-w-0 overflow-hidden bg-white ${activeTab === 'materials' ? 'md:rounded-tl-2xl border-l border-t border-gray-200 shadow-inner' : ''}`}>
        
        {activeTab === 'materials' && (
        <header className="bg-white border-b border-gray-100 z-30 shrink-0">
          <div className="px-4 sm:px-6 h-16 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="hidden sm:block">
                <h1 className="text-xl font-bold tracking-tight text-gray-800 leading-none truncate pr-2">Tiêu chuẩn Kỹ thuật</h1>
                <p className="text-xs text-gray-500 mt-1">Biên soạn và quản lý thông số vật tư</p>
              </div>
            </div>
            
            {isAuthenticated && (
              <div className="flex items-center gap-2 shrink-0">
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
                      ? 'bg-indigo-600 hover:bg-indigo-700 text-white focus:ring-indigo-500 disabled:opacity-75 disabled:cursor-not-allowed'
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
          </div>
        </header>
        )}

        <div className="flex-1 overflow-y-auto w-full p-4 sm:p-6 lg:p-8 bg-gray-50/30">
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

                {/* Danh mục tài liệu chứng minh */}
                {(() => {
                  const tagsSet = new Set<string>();
                  if (tables) {
                    tables.forEach(table => {
                      if (table.tags) {
                         Object.values(table.tags).forEach(tagStr => {
                           const tags = tagStr.split(',').map(t => t.trim()).filter(Boolean);
                           tags.forEach(t => tagsSet.add(t));
                         });
                      }
                    });
                  }
                  const activeVariants = Array.from(tagsSet);
                  const hasVariants = activeVariants.length > 0;

                  let hasAnyDocs = docRequirements.typeTest || docRequirements.catalog || docRequirements.endUser || docRequirements.iso;
                  if (docRequirements.variants) {
                    hasAnyDocs = hasAnyDocs || Object.values(docRequirements.variants).some(v => v.typeTest || v.catalog || v.endUser || v.iso);
                  }

                  if (!isEditing && !hasAnyDocs) return null;

                  return (
                    <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                      <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center gap-2">
                        <FileText className="text-blue-500" size={20} />
                        <h2 className="text-lg font-semibold text-gray-800">Danh mục các tài liệu chứng minh nguồn gốc và chất lượng hàng hóa</h2>
                      </div>
                      <div className="p-6">
                        {!hasVariants ? (
                          <div className="space-y-3">
                            <label className={`flex items-center gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                              <input 
                                type="checkbox" 
                                disabled={!isEditing}
                                checked={docRequirements.typeTest || false} 
                                onChange={(e) => setDocRequirements({...docRequirements, typeTest: e.target.checked})}
                                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                              />
                              <span className={`text-sm ${!isEditing && !docRequirements.typeTest ? 'text-gray-400 line-through' : 'text-gray-700'}`}>Biên bản thí nghiệm điển hình</span>
                            </label>
                            <label className={`flex items-center gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                              <input 
                                type="checkbox" 
                                disabled={!isEditing}
                                checked={docRequirements.catalog || false} 
                                onChange={(e) => setDocRequirements({...docRequirements, catalog: e.target.checked})}
                                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                              />
                              <span className={`text-sm ${!isEditing && !docRequirements.catalog ? 'text-gray-400 line-through' : 'text-gray-700'}`}>Tài liệu kỹ thuật (bản vẽ, Catalogue, ...)</span>
                            </label>
                            <label className={`flex items-center gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                              <input 
                                type="checkbox" 
                                disabled={!isEditing}
                                checked={docRequirements.endUser || false} 
                                onChange={(e) => setDocRequirements({...docRequirements, endUser: e.target.checked})}
                                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                              />
                              <span className={`text-sm ${!isEditing && !docRequirements.endUser ? 'text-gray-400 line-through' : 'text-gray-700'}`}>Xác nhận của đơn vị sử dụng cuối cùng</span>
                            </label>
                            <label className={`flex items-center gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                              <input 
                                type="checkbox" 
                                disabled={!isEditing}
                                checked={docRequirements.iso || false} 
                                onChange={(e) => setDocRequirements({...docRequirements, iso: e.target.checked})}
                                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                              />
                              <span className={`text-sm ${!isEditing && !docRequirements.iso ? 'text-gray-400 line-through' : 'text-gray-700'}`}>Chứng chỉ quản lý chất lượng ISO 9001 của nhà sản xuất</span>
                            </label>
                          </div>
                        ) : (
                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                            {activeVariants.map(variant => {
                              const varReqs = docRequirements.variants?.[variant] || { typeTest: false, catalog: false, endUser: false, iso: false };
                              
                              const handleVariantChange = (key: string, checked: boolean) => {
                                setDocRequirements(prev => ({
                                  ...prev,
                                  variants: {
                                    ...(prev.variants || {}),
                                    [variant]: {
                                      ...((prev.variants && prev.variants[variant]) || {}),
                                      [key]: checked
                                    }
                                  }
                                }));
                              };

                              return (
                                <div key={variant} className="border border-gray-200 rounded-lg p-4 bg-gray-50/50">
                                  <h3 className="font-semibold text-gray-800 mb-3 pb-2 border-b border-gray-200 flex items-center gap-2">
                                    <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                                    Tham biến: {variant}
                                  </h3>
                                  <div className="space-y-2.5">
                                    <label className={`flex items-center justify-between gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                                      <span className={`text-xs ${!isEditing && !varReqs.typeTest ? 'text-gray-400' : 'text-gray-700'}`}>Thí nghiệm điển hình</span>
                                      <input 
                                        type="checkbox" 
                                        disabled={!isEditing}
                                        checked={varReqs.typeTest || false} 
                                        onChange={(e) => handleVariantChange('typeTest', e.target.checked)}
                                        className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                                      />
                                    </label>
                                    <label className={`flex items-center justify-between gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                                      <span className={`text-xs ${!isEditing && !varReqs.catalog ? 'text-gray-400' : 'text-gray-700'}`}>Tài liệu kỹ thuật / Catalogue</span>
                                      <input 
                                        type="checkbox" 
                                        disabled={!isEditing}
                                        checked={varReqs.catalog || false} 
                                        onChange={(e) => handleVariantChange('catalog', e.target.checked)}
                                        className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                                      />
                                    </label>
                                    <label className={`flex items-center justify-between gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                                      <span className={`text-xs ${!isEditing && !varReqs.endUser ? 'text-gray-400' : 'text-gray-700'}`}>Xác nhận của đơn vị SD cuối</span>
                                      <input 
                                        type="checkbox" 
                                        disabled={!isEditing}
                                        checked={varReqs.endUser || false} 
                                        onChange={(e) => handleVariantChange('endUser', e.target.checked)}
                                        className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                                      />
                                    </label>
                                    <label className={`flex items-center justify-between gap-3 ${!isEditing ? 'cursor-default' : 'cursor-pointer'}`}>
                                      <span className={`text-xs ${!isEditing && !varReqs.iso ? 'text-gray-400' : 'text-gray-700'}`}>Chứng chỉ ISO 9001</span>
                                      <input 
                                        type="checkbox" 
                                        disabled={!isEditing}
                                        checked={varReqs.iso || false} 
                                        onChange={(e) => handleVariantChange('iso', e.target.checked)}
                                        className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 disabled:opacity-60" 
                                      />
                                    </label>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </section>
                  );
                })()}

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
                  <div className="px-6 py-4 border-b border-gray-100 bg-gray-50 flex items-center justify-between">
                    <div className="flex items-center gap-2">
                       <Database className="text-blue-500" size={20} />
                       <h2 className="text-lg font-semibold text-gray-800">Bảng thông số kỹ thuật</h2>
                    </div>
                    {isEditing && (
                       <button
                         type="button"
                         onClick={() => setTables([...tables, { id: crypto.randomUUID(), title: `Bảng thông số ${tables.length + 1}`, data: [['STT', 'Hạng mục', 'Đơn vị', 'Yêu cầu', 'Thông số chào'], ['1', 'Nhà sản xuất', '', 'Nêu cụ thể', ''], ['2', 'Nước sản xuất', '', 'Nêu cụ thể', '']], merges: [], tags: {} }])}
                         className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors"
                       >
                         <Plus size={16} /> Thêm bảng
                       </button>
                    )}
                  </div>
                  <div className="p-6 space-y-8">
                    {tables.map((table, tIndex) => (
                      <div key={table.id} className="space-y-4">
                        <div className="flex items-center gap-3">
                           {isEditing ? (
                             <input
                               type="text"
                               value={table.title}
                               onChange={e => {
                                  const newTables = [...tables];
                                  newTables[tIndex].title = e.target.value;
                                  setTables(newTables);
                               }}
                               className="font-semibold text-gray-800 border-b border-dashed border-gray-300 focus:border-blue-500 outline-none w-1/2 py-1"
                               placeholder="Tên bảng..."
                             />
                           ) : (
                             <h3 className="font-semibold text-gray-800">{table.title}</h3>
                           )}
                           
                           {isEditing && tables.length > 1 && (
                             <button
                               type="button"
                               onClick={() => setTables(tables.filter((_, i) => i !== tIndex))}
                               className="text-red-500 hover:text-red-700 p-1"
                               title="Xóa bảng này"
                             >
                               <Trash2 size={16} />
                             </button>
                           )}
                        </div>
                        
                        <ExcelTable 
                          tableData={table.data} 
                          setTableData={newData => {
                             const newTables = [...tables];
                             newTables[tIndex].data = typeof newData === 'function' ? newData(newTables[tIndex].data) : newData;
                             setTables(newTables);
                          }} 
                          merges={table.merges} 
                          setMerges={newMerges => {
                             const newTables = [...tables];
                             newTables[tIndex].merges = typeof newMerges === 'function' ? newMerges(newTables[tIndex].merges) : newMerges;
                             setTables(newTables);
                          }}
                          rowTags={table.tags}
                          setRowTags={newTags => {
                             const newTables = [...tables];
                             newTables[tIndex].tags = typeof newTags === 'function' ? newTags(newTables[tIndex].tags) : newTags;
                             setTables(newTables);
                          }}
                          readOnly={!isEditing}
                        />
                      </div>
                    ))}
                    
                    {tables.length === 0 && (
                       <div className="text-center p-8 text-gray-500 text-sm">Chưa có bảng thông số nào.</div>
                    )}
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
        </div>
      </div>
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
                  {(() => {
                    const firstCreate = auditLogsData.slice().reverse().find(l => l.action === 'create') || auditLogsData.find(l => l.action === 'create');
                    const updates = auditLogsData.filter(l => l.id !== firstCreate?.id).slice(0, 4);
                    const list = firstCreate ? [...updates, firstCreate] : updates;
                    return list.map((log) => (
                       <div key={log.id} className="bg-white rounded-lg border border-gray-200 p-4 shadow-sm">
                         <div className="flex justify-between items-start">
                         <div className="flex gap-2 items-center">
                            <span className={`text-xs font-bold uppercase px-2 py-1 rounded ${log.action === 'create' ? 'bg-emerald-100 text-emerald-700' : log.action === 'update' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                               {log.action === 'create' ? 'Tạo mới' : log.action === 'update' ? 'Cập nhật' : 'Xóa'}
                            </span>
                            <span className="text-sm font-semibold text-gray-800">{log.updatedBy || 'Người dùng không xác định'}</span>
                         </div>
                         <span className="text-sm text-gray-500">{new Date(log.updatedAt).toLocaleString('vi-VN')}</span>
                       </div>
                     </div>
                   ));
                 })()}
                 </div>
              )}
            </div>
          </div>
        </div>
      )}

    </div>
  );
}


