import React, { useState } from 'react';
import { UploadCloud, X, Image as ImageIcon, Download, RotateCw, ZoomIn } from 'lucide-react';

interface ImageUploadProps {
  images: { url: string; name: string }[];
  setImages: React.Dispatch<React.SetStateAction<{ url: string; name: string }[]>>;
  readOnly?: boolean;
}

export function ImageUpload({ images, setImages, readOnly }: ImageUploadProps) {
  const [previewIndex, setPreviewIndex] = useState<number | null>(null);

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files) return;

    Array.from(files).forEach((file) => {
      if (!file.type.startsWith('image/')) return;
      
      const reader = new FileReader();
      reader.onload = (e) => {
        if (e.target?.result) {
          setImages(prev => [...prev, { url: e.target!.result as string, name: file.name }]);
        }
      };
      reader.readAsDataURL(file);
    });
    
    // Reset input
    event.target.value = '';
  };

  const removeImage = (index: number, e: React.MouseEvent) => {
    e.stopPropagation();
    setImages(images.filter((_, i) => i !== index));
    if (previewIndex === index) {
      setPreviewIndex(null);
    }
  };

  const rotateImage = (index: number, e: React.MouseEvent) => {
    e.stopPropagation();
    const imgData = images[index];
    const imgElement = new window.Image();
    imgElement.src = imgData.url;
    imgElement.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = imgElement.height;
      canvas.height = imgElement.width;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.translate(canvas.width / 2, canvas.height / 2);
        ctx.rotate((90 * Math.PI) / 180);
        ctx.drawImage(imgElement, -imgElement.width / 2, -imgElement.height / 2);
        const newDataUrl = canvas.toDataURL('image/png');
        setImages((prev) => {
          const newImages = [...prev];
          newImages[index] = { ...imgData, url: newDataUrl };
          return newImages;
        });
      }
    };
  };

  const downloadImage = (index: number, e: React.MouseEvent) => {
    e.stopPropagation();
    const imgData = images[index];
    const a = document.createElement('a');
    a.href = imgData.url;
    a.download = imgData.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="w-full">
      {!readOnly && (
        <div className="w-full relative border-2 border-dashed border-gray-300 rounded-lg p-6 bg-gray-50 hover:bg-gray-100 transition-colors flex flex-col items-center justify-center min-h-[140px] text-center group cursor-pointer">
          <input
            type="file"
            accept="image/*"
            multiple
            onChange={handleImageUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
          />
          <UploadCloud className="w-8 h-8 text-blue-500 mb-2 group-hover:scale-110 transition-transform" />
          <p className="text-sm font-medium text-gray-700">Kéo thả hoặc nhấn để tải ảnh lên</p>
          <p className="text-xs text-gray-500 mt-1">Hỗ trợ PNG, JPG, GIF</p>
        </div>
      )}

      {images.length > 0 && (
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4 mt-6">
          {images.map((img, index) => (
            <div 
              key={index} 
              className="relative group rounded-lg overflow-hidden border border-gray-200 shadow-sm aspect-square bg-white flex items-center justify-center cursor-pointer"
              onClick={() => setPreviewIndex(index)}
            >
              <img src={img.url} alt={img.name} className="w-full h-full object-cover transition-transform group-hover:scale-105" />
              
              <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center pointer-events-none">
                 <ZoomIn className="text-white w-8 h-8 pointer-events-none" />
              </div>

              {!readOnly && (
                <div className="absolute top-2 right-2 flex flex-col gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                   <button 
                     type="button"
                     onClick={(e) => removeImage(index, e)}
                     className="p-1.5 bg-red-500 hover:bg-red-600 text-white rounded-full shadow-md transition-transform hover:scale-105"
                     title="Xóa ảnh"
                   >
                     <X size={14} />
                   </button>
                </div>
              )}
              
              <div className="absolute bottom-0 left-0 right-0 bg-black/60 p-1.5 text-white text-xs truncate max-w-full opacity-0 group-hover:opacity-100 transition-opacity flex justify-between items-center z-10">
                <span className="truncate flex-1">{img.name}</span>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Fullscreen Preview Modal */}
      {previewIndex !== null && images[previewIndex] && (
        <div 
          className="fixed inset-0 z-[100] bg-black/90 flex items-center justify-center p-4"
          onClick={() => setPreviewIndex(null)}
        >
          <div className="absolute top-4 right-4 flex gap-4">
            <button 
              className="p-2 bg-white/20 hover:bg-white/30 text-white rounded-full transition-colors"
              onClick={(e) => downloadImage(previewIndex, e)}
              title="Tải xuống"
            >
              <Download size={20} />
            </button>
            {!readOnly && (
              <button 
                className="p-2 bg-white/20 hover:bg-white/30 text-white rounded-full transition-colors"
                onClick={(e) => rotateImage(previewIndex, e)}
                title="Xoay 90 độ"
              >
                <RotateCw size={20} />
              </button>
            )}
            <button 
              className="p-2 bg-white/20 hover:bg-white/30 text-white rounded-full transition-colors"
              onClick={(e) => {
                e.stopPropagation();
                setPreviewIndex(null);
              }}
              title="Đóng"
            >
              <X size={20} />
            </button>
          </div>
          
          <img 
            src={images[previewIndex].url} 
            alt={images[previewIndex].name} 
            className="max-w-full max-h-full object-contain select-none"
            onClick={(e) => e.stopPropagation()}
          />
        </div>
      )}
    </div>
  );
}
