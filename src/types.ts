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
}

export interface BiddingPackage {
  id: string;
  name: string;
  materialIds: string[];
  hiddenTags?: Record<string, string[]>;
}
