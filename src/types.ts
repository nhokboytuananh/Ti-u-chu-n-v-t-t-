import * as xlsx from 'xlsx';

export interface Material {
  id: string;
  code: string;
  name: string;
  richText: string;
  excelData: string[][];
  excelMerges: xlsx.Range[];
  rowTags?: Record<number, string>;
  images: { url: string; name: string }[];
  notes: string;
  updatedBy?: string;
  updatedAt?: string;
}

export interface MaterialAuditLog {
  id: string;
  materialId: string;
  action: string;
  previousData: any;
  newData: any;
  updatedBy: string;
  updatedAt: string;
}

export interface BiddingPackage {
  id: string;
  name: string;
  materialIds: string[];
  hiddenTags?: Record<string, string[]>;
}
