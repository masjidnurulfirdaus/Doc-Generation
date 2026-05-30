/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import {
  Plus,
  Trash2,
  Download,
  Printer,
  FileText,
  Edit2,
  Sparkles,
  UploadCloud,
  AlertCircle,
  Eye,
  HelpCircle,
  Check,
  ChevronRight,
  Info,
  Layers,
  Settings,
  FileSpreadsheet,
  RefreshCw,
  X
} from 'lucide-react';
import {
  getAllTemplates,
  saveTemplate,
  deleteTemplate,
  DocTemplateRecord,
  CustomField
} from './db';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
// @ts-ignore
import mammoth from 'mammoth';
// @ts-ignore
import { renderAsync } from 'docx-preview';

// --- Default Templates & Presets ---
const PRESET_INVITATION = `PANITIA REUNI AKBAR SEKOLAH MENENGAH ATAS 1
Jl. Merdeka No. 100, Jakarta Pusat
===================================================

Jakarta, {tanggal}

Hal: Undangan Rapat Pleno Panitia
Sifat: Penting

Kepada Yth,
Sdr/i. {nama}
di Tempat

Dengan hormat,
Sehubungan dengan kelanjutan rencana pelaksanaan Reuni Akbar, kami mengharapkan kehadiran rekan panitia dalam Rapat Koordinasi {acara}, yang akan diselenggarakan pada:

Hari/Tanggal : Sabtu, {tanggal}
Waktu        : 14.00 WIB s.d. Selesai
Tempat       : Aula Serbaguna Lantai 2

Demikian undangan ini kami sampaikan, mengingat pentingnya koordinasi {acara} ini, kehadiran Saudara sangat kami harapkan.

Hormat kami,

( Ketua Panitia )`;

const PRESET_CERTIFICATE = `===================================================
                  PIAGAM PENGHARGAAN                  
        No: {nomor_sertifikat}/PAN-REUNI/2026        
===================================================

Diberikan Kepada Sdr/i:
{nama}

Atas partisipasi aktif dan pengabdian luar biasa sebagai:
PANITIA PELAKSANA (SEKSI {acara})

Dalam menyukseskan Kegiatan Reuni Akbar SMA Negeri 1 
yang diselenggarakan pada tanggal {tanggal}.

Kami mengucapkan terima kasih yang sebesar-besarnya atas kontribusi,
kerja keras, dan dedikasi yang telah dicurahkan.

Jakarta, {tanggal}


    Kepala SMAN 1 Jakarta         |         Ketua Panitia
   
      ( Drs. H. Budiman )         |       ( Rekan Sejawat )`;

const PRESET_INVITATION_SUB1 = "Joko\nAgung\nSiti Rahma\nBambang\nClara Wijaya";
const PRESET_INVITATION_SUB2 = "Andi\nBudi\nJoni\nDewi Lestari";

// --- Word Generation Helpers ---
const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = error => reject(error);
  });
};

const dataURLToArrayBuffer = (dataUrl: string): ArrayBuffer => {
  const base64Part = dataUrl.split(',')[1];
  const binaryString = window.atob(base64Part);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
};

const prepareZipFile = (docxFileData: string, maxLen: number): PizZip => {
  const content = dataURLToArrayBuffer(docxFileData);
  const zip = new PizZip(content);

  // 1. Clean mailMerge SQL alert in settings.xml
  try {
    const settingsXml = zip.file("word/settings.xml")?.asText();
    if (settingsXml) {
       const cleanedSettings = settingsXml
         .replace(/<w:mailMerge[^>]*>[\s\S]*?<\/w:mailMerge>/gi, "")
         .replace(/<mailMerge[^>]*>[\s\S]*?<\/mailMerge>/gi, "");
       zip.file("word/settings.xml", cleanedSettings);
    }
  } catch (err) {
    console.warn("Gagal membersihkan settings.xml:", err);
  }

  // 2. Validate and auto-wrap document.xml body in loops if there are no loop tags
  try {
    let docXml = zip.file("word/document.xml")?.asText();
    if (docXml) {
      // Check if there's any manual loop tag
      const hasLoop = docXml.includes("penerima_list") || docXml.includes("{#") || docXml.includes("{/");
      
      if (!hasLoop && maxLen > 0) {
        const bodyStart = docXml.indexOf("<w:body>");
        if (bodyStart !== -1) {
          const sectPrIdx = docXml.lastIndexOf("<w:sectPr");
          if (sectPrIdx !== -1) {
            const pStart = docXml.slice(0, bodyStart + 8);
            const pEnd = docXml.slice(sectPrIdx);
            const pMiddle = docXml.slice(bodyStart + 8, sectPrIdx);
            
            docXml = pStart + 
              "<w:p><w:r><w:t>{#penerima_list}</w:t></w:r></w:p>" + 
              pMiddle + 
              "<w:p><w:r><w:t>{#hasMore}</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r><w:r><w:t>{/hasMore}</w:t></w:r></w:p>" +
              "<w:p><w:r><w:t>{/penerima_list}</w:t></w:r></w:p>" + 
              pEnd;
          } else {
            const bodyEnd = docXml.lastIndexOf("</w:body>");
            if (bodyEnd !== -1) {
              const pStart = docXml.slice(0, bodyStart + 8);
              const pEnd = docXml.slice(bodyEnd);
              const pMiddle = docXml.slice(bodyStart + 8, bodyEnd);
              
              docXml = pStart + 
                "<w:p><w:r><w:t>{#penerima_list}</w:t></w:r></w:p>" + 
                pMiddle + 
                "<w:p><w:r><w:t>{#hasMore}</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r><w:r><w:t>{/hasMore}</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>{/penerima_list}</w:t></w:r></w:p>" + 
                pEnd;
            }
          }
          zip.file("word/document.xml", docXml);
        }
      }
    }
  } catch (err) {
    console.warn("Gagal memodifikasi document.xml:", err);
  }

  return zip;
};

export default function App() {
  // --- States ---
  const [templates, setTemplates] = useState<DocTemplateRecord[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState<DocTemplateRecord | null>(null);
  const [showHelpModal, setShowHelpModal] = useState<boolean>(false);
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [dragActive, setDragActive] = useState<boolean>(false);
  
  // --- Form States for creation/edit ---
  const [formId, setFormId] = useState<string>('');
  const [formTitle, setFormTitle] = useState<string>('');
  const [formTemplateType, setFormTemplateType] = useState<'docx' | 'web'>('docx');
  const [formDocxFileName, setFormDocxFileName] = useState<string | null>(null);
  const [formDocxFileData, setFormDocxFileData] = useState<string | null>(null);
  const [formWebContent, setFormWebContent] = useState<string>(PRESET_INVITATION);
  const [formRecipientsPerPage, setFormRecipientsPerPage] = useState<1 | 2>(1);
  const [formRecipients1, setFormRecipients1] = useState<string>(PRESET_INVITATION_SUB1);
  const [formRecipients2, setFormRecipients2] = useState<string>(PRESET_INVITATION_SUB2);
  const [formCustomFields, setFormCustomFields] = useState<CustomField[]>([
    { id: '1', name: 'acara', value: 'Pembubaran Panitia' },
    { id: '2', name: 'tanggal', value: '30 Mei 2026' },
  ]);

  // --- Preview Navigation State ---
  const [previewPage, setPreviewPage] = useState<number>(0);
  const [generationError, setGenerationError] = useState<string | null>(null);
  const [generationSuccess, setGenerationSuccess] = useState<boolean>(false);

  // --- DOCX Rendering & Print Refs ---
  const docxPreviewContainerRef = useRef<HTMLDivElement>(null);
  const docxPrintContainerRef = useRef<HTMLDivElement>(null);
  const [isRenderingDocx, setIsRenderingDocx] = useState<boolean>(false);

  // --- Initialization ---
  useEffect(() => {
    loadAllTemplatesData();
  }, []);

  // --- Dynamic DOCX Live Preview Compiler ---
  useEffect(() => {
    let active = true;
    const renderDocx = async () => {
      if (!selectedTemplate || selectedTemplate.templateType !== 'docx' || !selectedTemplate.docxFileData) {
        return;
      }
      
      setIsRenderingDocx(true);
      try {
        const { maxLen } = getRecipientsLists(selectedTemplate);
        if (maxLen === 0) {
          if (docxPreviewContainerRef.current) {
            docxPreviewContainerRef.current.innerHTML = '<div class="p-8 text-center text-slate-500 font-medium">Silakan isi baris penerima terlebih dahulu untuk melihat pratinjau.</div>';
          }
          setIsRenderingDocx(false);
          return;
        }

        const zip = prepareZipFile(selectedTemplate.docxFileData, maxLen);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        const pageVars = getPageVariables(selectedTemplate, previewPage);
        const singleItem = {
          ...pageVars,
          hasMore: false,
          isLast: true
        };

        doc.render({
          penerima_list: [singleItem],
          ...singleItem
        });

        const out = doc.getZip().generate({
          type: 'arraybuffer',
        });

        if (active && docxPreviewContainerRef.current) {
          docxPreviewContainerRef.current.innerHTML = ''; // Clean prior render
          await renderAsync(out, docxPreviewContainerRef.current, undefined, {
            className: "docx-preview",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            breakPages: true
          });
        }
      } catch (err: any) {
        console.error('Gagal merender pratonton DOCX dengan docx-preview:', err);
        if (active && docxPreviewContainerRef.current) {
          docxPreviewContainerRef.current.innerHTML = `
            <div class="p-8 text-red-700 bg-red-50 rounded-xl border border-red-200">
              <p class="font-bold">Gagal menggabungkan variabel ke pratinjau DOCX.</p>
              <p class="text-xs mt-1.5 font-mono">${err?.message || err}</p>
              <p class="text-xs text-slate-500 mt-2">Pastikan seluruh tag kurung kurawal di dokumen Word Anda tertutup dengan benar, misalnya {nama} dan tidak ada tag kosong {} atau tag yang salah penulisan.</p>
            </div>`;
        }
      } finally {
        if (active) {
          setIsRenderingDocx(false);
        }
      }
    };

    // Delay render briefly to let the container mount if templates changed
    const timer = setTimeout(() => {
      renderDocx();
    }, 50);

    return () => {
      active = false;
      clearTimeout(timer);
    };
  }, [selectedTemplate, previewPage, templates]);

  // --- Dynamic DOCX Print Pages Pre-compiler ---
  useEffect(() => {
    let active = true;
    const preparePrintPages = async () => {
      if (!selectedTemplate || selectedTemplate.templateType !== 'docx' || !selectedTemplate.docxFileData) {
        return;
      }
      const { r1, r2, maxLen } = getRecipientsLists(selectedTemplate);
      if (maxLen === 0) {
        return;
      }

      try {
        const zip = prepareZipFile(selectedTemplate.docxFileData!, maxLen);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        // global variables as fallback
        const globalVars: Record<string, string> = {};
        selectedTemplate.customFields.forEach(f => {
          globalVars[f.name] = f.value;
        });

        const listData = [];
        for (let i = 0; i < maxLen; i++) {
          const item: Record<string, any> = { ...globalVars };
          if (selectedTemplate.recipientsPerPage === 1) {
            item['nama'] = r1[i] || '-';
          } else {
            item['nama1'] = r1[i] || '-';
            item['nama2'] = r2[i] || '-';
            item['nama'] = `${r1[i] || '-'} & ${r2[i] || '-'}`;
          }
          item['hasMore'] = (i < maxLen - 1);
          item['isLast'] = (i === maxLen - 1);
          listData.push(item);
        }

        doc.render({
          penerima_list: listData,
          ...globalVars,
          nama: listData[0]?.nama || '-',
          nama1: listData[0]?.nama1 || '-',
          nama2: listData[0]?.nama2 || '-'
        });

        const out = doc.getZip().generate({
          type: 'arraybuffer',
        });

        if (active && docxPrintContainerRef.current) {
          docxPrintContainerRef.current.innerHTML = '';
          await renderAsync(out, docxPrintContainerRef.current, undefined, {
            className: "docx-print",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            breakPages: true
          });
        }
      } catch (err) {
        console.error('Gagal menyiapkan cetakan DOCX dengan docx-preview:', err);
      }
    };

    const timer = setTimeout(() => {
      preparePrintPages();
    }, 100);

    return () => {
      active = false;
      clearTimeout(timer);
    };
  }, [selectedTemplate, templates]);

  const loadAllTemplatesData = async () => {
    try {
      const all = await getAllTemplates();
      setTemplates(all);
      if (all.length > 0 && !selectedTemplate) {
        setSelectedTemplate(all[0]);
      }
    } catch (err) {
      console.error('Gagal memuat database templates:', err);
    }
  };

  // --- Helper Helpers ---
  const getRecipientsLists = (t: DocTemplateRecord | null) => {
    if (!t) return { r1: [], r2: [], maxLen: 0 };
    const r1 = t.recipients1.split('\n').map(x => x.trim()).filter(x => x.length > 0);
    const r2 = t.recipientsPerPage === 2 
      ? t.recipients2.split('\n').map(x => x.trim()).filter(x => x.length > 0)
      : [];
    const maxLen = t.recipientsPerPage === 1 ? r1.length : Math.max(r1.length, r2.length);
    return { r1, r2, maxLen };
  };

  const getPageVariables = (t: DocTemplateRecord, pageIdx: number) => {
    const { r1, r2 } = getRecipientsLists(t);
    const vars: Record<string, string> = {};
    
    // Global fields
    t.customFields.forEach(f => {
      if (f.name) vars[f.name] = f.value;
    });

    if (t.recipientsPerPage === 1) {
      vars['nama'] = r1[pageIdx] || '-';
    } else {
      vars['nama1'] = r1[pageIdx] || '-';
      vars['nama2'] = r2[pageIdx] || '-';
      vars['nama'] = `${r1[pageIdx] || '-'} & ${r2[pageIdx] || '-'}`;
    }
    return vars;
  };

  const replacePlaceholders = (text: string, vars: Record<string, string>) => {
    let output = text;
    Object.entries(vars).forEach(([k, v]) => {
      // replace all occurrences of {k}
      const regex = new RegExp(`{${k}}`, 'g');
      output = output.replace(regex, v);
    });
    return output;
  };

  // --- Form & Action Handlers ---
  const handleAddNewTemplate = () => {
    setFormId('');
    setFormTitle('Undangan Panitia Baru');
    setFormTemplateType('docx');
    setFormDocxFileName(null);
    setFormDocxFileData(null);
    setFormWebContent(PRESET_INVITATION);
    setFormRecipientsPerPage(1);
    setFormRecipients1(PRESET_INVITATION_SUB1);
    setFormRecipients2(PRESET_INVITATION_SUB2);
    setFormCustomFields([
      { id: '1', name: 'acara', value: 'Pembubaran Panitia' },
      { id: '2', name: 'tanggal', value: '30 Mei 2026' }
    ]);
    setIsEditing(true);
  };

  const handleEditTemplate = (t: DocTemplateRecord) => {
    setFormId(t.id);
    setFormTitle(t.title);
    setFormTemplateType(t.templateType);
    setFormDocxFileName(t.docxFileName);
    setFormDocxFileData(t.docxFileData);
    setFormWebContent(t.webContent);
    setFormRecipientsPerPage(t.recipientsPerPage);
    setFormRecipients1(t.recipients1);
    setFormRecipients2(t.recipients2);
    setFormCustomFields([...t.customFields]);
    setIsEditing(true);
  };

  const handleSaveForm = async () => {
    if (!formTitle.trim()) {
      alert('Judul dokumen wajib diisi!');
      return;
    }
    if (formTemplateType === 'docx' && !formDocxFileData) {
      alert('Silakan upload file template (.docx) terlebih dahulu!');
      return;
    }

    const record: DocTemplateRecord = {
      id: formId || `tmpl_${Date.now()}`,
      title: formTitle.trim(),
      templateType: formTemplateType,
      docxFileName: formDocxFileName,
      docxFileData: formDocxFileData,
      webContent: formWebContent,
      recipientsPerPage: formRecipientsPerPage,
      recipients1: formRecipients1,
      recipients2: formRecipients2,
      customFields: formCustomFields.filter(f => f.name.trim().length > 0),
      createdAt: Date.now()
    };

    try {
      await saveTemplate(record);
      await loadAllTemplatesData();
      setSelectedTemplate(record);
      setIsEditing(false);
      setPreviewPage(0);
      setGenerationSuccess(true);
      setTimeout(() => setGenerationSuccess(false), 3000);
    } catch (err) {
      console.error('Error saving template:', err);
      alert('Gagal mendaur ulang template ke database.');
    }
  };

  const handleDeleteTemplate = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm('Apakah Anda yakin ingin menghapus record ini secara permanen?')) {
      return;
    }
    try {
      await deleteTemplate(id);
      const updated = templates.filter(t => t.id !== id);
      setTemplates(updated);
      if (selectedTemplate?.id === id) {
        setSelectedTemplate(updated.length > 0 ? updated[0] : null);
        setPreviewPage(0);
      }
    } catch (err) {
      console.error('Gagal menghapus template:', err);
    }
  };

  // --- Dynamic Var Handlers ---
  const handleAddCustomField = () => {
    const newField: CustomField = {
      id: `field_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      name: '',
      value: ''
    };
    setFormCustomFields([...formCustomFields, newField]);
  };

  const handleUpdateCustomFieldName = (id: string, val: string) => {
    setFormCustomFields(
      formCustomFields.map(f => (f.id === id ? { ...f, name: val.replace(/[^a-zA-Z0-9_]/g, '') } : f))
    );
  };

  const handleUpdateCustomFieldValue = (id: string, val: string) => {
    setFormCustomFields(
      formCustomFields.map(f => (f.id === id ? { ...f, value: val } : f))
    );
  };

  const handleRemoveCustomField = (id: string) => {
    setFormCustomFields(formCustomFields.filter(f => f.id !== id));
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (!file.name.endsWith('.docx')) {
        alert('Format file harus berupa .docx (Microsoft Word)!');
        return;
      }
      try {
        const base64 = await fileToBase64(file);
        setFormDocxFileName(file.name);
        setFormDocxFileData(base64);
      } catch (err) {
        console.error('Gagal memproses file:', err);
        alert('Terjadi kesalahan saat membaca file docx.');
      }
    }
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (!file.name.endsWith('.docx')) {
        alert('Format file harus berupa .docx (Microsoft Word)!');
        return;
      }
      try {
        const base64 = await fileToBase64(file);
        setFormDocxFileName(file.name);
        setFormDocxFileData(base64);
      } catch (err) {
        console.error('Gagal memproses file:', err);
        alert('Gagal membaca file docx.');
      }
    }
  };

  // --- DOCX Mail Merge Generator ---
  const handleGeneratedocx = () => {
    if (!selectedTemplate) return;
    if (selectedTemplate.templateType !== 'docx' || !selectedTemplate.docxFileData) {
      alert('Aplikasi ini tidak memiliki berkas docx yang valid.');
      return;
    }

    setGenerationError(null);

    const { r1, r2, maxLen } = getRecipientsLists(selectedTemplate);
    if (maxLen === 0) {
      alert('Harap isi baris penerima terlebih dahulu sebelum menggenerasi!');
      return;
    }

    try {
      const zip = prepareZipFile(selectedTemplate.docxFileData, maxLen);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // global variables as fallback
      const globalVars: Record<string, string> = {};
      selectedTemplate.customFields.forEach(f => {
        globalVars[f.name] = f.value;
      });

      // Build loop structure for {#penerima_list} ... {/penerima_list}
      const listData = [];
      for (let i = 0; i < maxLen; i++) {
        const item: Record<string, any> = { ...globalVars };
        if (selectedTemplate.recipientsPerPage === 1) {
          item['nama'] = r1[i] || '-';
        } else {
          item['nama1'] = r1[i] || '-';
          item['nama2'] = r2[i] || '-';
          item['nama'] = `${r1[i] || '-'} & ${r2[i] || '-'}`;
        }
        item['hasMore'] = (i < maxLen - 1);
        item['isLast'] = (i === maxLen - 1);
        listData.push(item);
      }

      // Render document
      doc.render({
        penerima_list: listData,
        ...globalVars, // include also in root for non-loop static templates
        nama: listData[0]?.nama || '-',
        nama1: listData[0]?.nama1 || '-',
        nama2: listData[0]?.nama2 || '-'
      });

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      // Trigger automatic browser download
      const url = URL.createObjectURL(out);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Mail_Merge_${selectedTemplate.title.replace(/\s+/g, '_')}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (error: any) {
      console.error(error);
      setGenerationError(
        'Gagal menggabungkan variabel ke DOCX. Pastikan tag template Anda ditulis dengan benar. Contoh: {nama} atau {nama1}, bukan {[nama]}.'
      );
    }
  };

  // --- Dynamic Printing Action ---
  const handlePrintDocument = () => {
    window.print();
  };

  // --- Render lists & states values ---
  const currentDetails = selectedTemplate;
  const { r1: previewR1, r2: previewR2, maxLen: previewMaxLen } = getRecipientsLists(currentDetails);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col antialiased">
      
      {/* HEADER BAR */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40 px-6 py-4 flex items-center justify-between shadow-xs print:hidden">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 text-white p-2 rounded-xl shadow-md shadow-blue-100 flex items-center justify-center">
            <Layers className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">DocuMerge Studio</h1>
            <p className="text-xs text-slate-500 font-medium">Auto-Mailing Creator & Document Generator</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowHelpModal(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-600 text-sm font-semibold transition"
          >
            <HelpCircle className="w-4 h-4" />
            Panduan Word
          </button>
          
          <button
            id="create-template-btn"
            onClick={handleAddNewTemplate}
            className="flex items-center gap-2 bg-slate-900 hover:bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-semibold shadow-xs transition"
          >
            <Plus className="w-4 h-4" />
            Template Baru
          </button>
        </div>
      </header>

      {/* BODY WRAPPER */}
      <div className="flex-1 flex flex-col md:flex-row print:block">
        
        {/* SIDEBAR: SAVED LIST */}
        <aside className="w-full md:w-80 bg-white border-b md:border-b-0 md:border-r border-slate-200 p-5 flex flex-col gap-4 print:hidden shrink-0">
          <div className="flex items-center justify-between">
            <h2 className="font-bold text-slate-800 text-sm uppercase tracking-wider">Daftar Dokumen ({templates.length})</h2>
          </div>

          <div className="flex-1 overflow-y-auto space-y-2.5 max-h-[65vh] md:max-h-none">
            {templates.length === 0 ? (
              <div className="text-center py-8 px-4 border-2 border-dashed border-slate-200 rounded-xl bg-slate-50/50">
                <FileText className="w-8 h-8 text-slate-400 mx-auto mb-2" />
                <p className="text-xs text-slate-500 font-medium font-sans">Belum ada template tersimpan</p>
              </div>
            ) : (
              templates.map(t => {
                const isActive = selectedTemplate?.id === t.id && !isEditing;
                const { maxLen } = getRecipientsLists(t);
                return (
                  <div
                    key={t.id}
                    onClick={() => {
                      setSelectedTemplate(t);
                      setIsEditing(false);
                      setPreviewPage(0);
                    }}
                    className={`group w-full text-left p-3.5 rounded-xl cursor-pointer border transition flex items-start gap-3 relative ${
                      isActive
                        ? 'bg-blue-50/60 border-blue-200 text-blue-900 ring-1 ring-blue-100'
                        : 'border-slate-100 hover:bg-slate-50 hover:border-slate-200 text-slate-700'
                    }`}
                  >
                    <div className={`p-2 rounded-lg shrink-0 ${
                      t.templateType === 'docx' 
                        ? 'bg-emerald-50 text-emerald-600' 
                        : 'bg-amber-50 text-amber-600'
                    }`}>
                      <FileText className="w-4 h-4" />
                    </div>
                    <div className="flex-1 min-w-0 pr-6">
                      <h3 className="font-semibold text-sm text-slate-900 truncate leading-snug">{t.title}</h3>
                      <div className="flex items-center gap-2 mt-1 text-xs text-slate-500 font-medium">
                        <span>{t.templateType.toUpperCase()}</span>
                        <span>•</span>
                        <span>{maxLen} Penerima</span>
                      </div>
                    </div>
                    <button
                      onClick={(e) => handleDeleteTemplate(t.id, e)}
                      className="absolute right-3 top-3.5 p-1 rounded-md text-slate-400 hover:text-red-600 opacity-0 group-hover:opacity-100 focus:opacity-100 transition"
                      title="Hapus template"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                );
              })
            )}
          </div>

          <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 text-xs text-slate-600 leading-relaxed font-sans mt-auto">
            <div className="flex items-center gap-1.5 font-bold text-slate-800 mb-1.5">
              <Sparkles className="w-3.5 h-3.5 text-blue-650 shrink-0" />
              <span>Database Server Aktif</span>
            </div>
            Data template, variabel, dan berkas Word Anda disimpan secara aman dalam **Database Server Lokal** (<code className="bg-slate-200 px-1 py-0.5 rounded text-blue-700 font-bold">templates_db.json</code>).
            <div className="mt-2 text-slate-500">
              Sangat cocok untuk **deploy di server lokal** dan dibuka kembali dari komputer mana saja tanpa takut kehilangan data. Browser IndexedDB juga terus dicadangkan sebagai backup hibrida.
            </div>
          </div>
        </aside>

        {/* CONTAINER UTAMA */}
        <main className="flex-1 flex flex-col min-w-0 bg-slate-50 print:bg-white print:p-0">
          
          {/* JIKA FORM TAMBAH / EDIT DIBUKA */}
          {isEditing ? (
            <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-6 print:hidden max-w-4xl mx-auto w-full">
              
              <div className="flex items-center justify-between border-b border-slate-200 pb-4">
                <div>
                  <span className="text-xs font-bold text-blue-600 tracking-wider uppercase">Konfigurasi Record</span>
                  <h2 className="text-2xl font-extrabold text-slate-900 tracking-tight">
                    {formId ? 'Edit Draft Template' : 'Buat Template Baru'}
                  </h2>
                </div>
                <button
                  onClick={() => setIsEditing(false)}
                  className="px-3.5 py-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 font-semibold text-sm transition"
                >
                  Batal
                </button>
              </div>

              {/* SECTION 1: PROPERTI DOKUMEN */}
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                <div>
                  <label className="block text-sm font-bold text-slate-800 mb-1.5">Judul Template / Nama Kegiatan</label>
                  <input
                    type="text"
                    value={formTitle}
                    onChange={(e) => setFormTitle(e.target.value)}
                    placeholder="Contoh: Undangan Rapat Pleno atau Piagam Anggota..."
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-900 font-medium"
                  />
                </div>

                <div>
                  <label className="block text-sm font-bold text-slate-800 mb-1.5">Model Implementasi</label>
                  <div className="grid grid-cols-2 gap-3">
                    <button
                      type="button"
                      onClick={() => setFormTemplateType('docx')}
                      className={`flex items-center justify-center gap-2 p-3 rounded-xl border text-sm font-bold transition ${
                        formTemplateType === 'docx'
                          ? 'border-emerald-600 bg-emerald-50 text-emerald-800 ring-2 ring-emerald-100'
                          : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      <FileText className="w-4 h-4" />
                      Upload File Template Word (.docx)
                    </button>
                    <button
                      type="button"
                      onClick={() => setFormTemplateType('web')}
                      className={`flex items-center justify-center gap-2 p-3 rounded-xl border text-sm font-bold transition ${
                        formTemplateType === 'web'
                          ? 'border-amber-600 bg-amber-50 text-amber-800 ring-2 ring-amber-100'
                          : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      <Edit2 className="w-4 h-4" />
                      Visual Web Template (Tanpa Upload)
                    </button>
                  </div>
                </div>
              </div>

              {/* SECTION 2: TEMPLATE INPUTS */}
              {formTemplateType === 'docx' ? (
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                  <div className="flex items-center justify-between">
                    <label className="block text-sm font-bold text-slate-800">Berkas Template Word (.docx)</label>
                    <button
                      type="button"
                      onClick={() => setShowHelpModal(true)}
                      className="text-xs text-blue-600 hover:underline flex items-center gap-1 font-semibold"
                    >
                      <Info className="w-3 h-3" /> Bagaimana memasukkan variabel di Word?
                    </button>
                  </div>

                  <div
                    onDragEnter={handleDrag}
                    onDragOver={handleDrag}
                    onDragLeave={handleDrag}
                    onDrop={handleDrop}
                    className={`border-2 border-dashed rounded-2xl p-6 text-center transition flex flex-col items-center justify-center ${
                      dragActive ? 'border-blue-500 bg-blue-50/50' : 'border-slate-200 hover:border-slate-300'
                    }`}
                  >
                    <UploadCloud className="w-10 h-10 text-slate-400 mb-2.5" />
                    {formDocxFileName ? (
                      <div className="space-y-1">
                        <p className="text-sm font-semibold text-slate-800">{formDocxFileName}</p>
                        <p className="text-xs text-emerald-600 font-bold flex items-center justify-center gap-1">
                          <Check className="w-3.5 h-3.5" /> Berhasil Diunggah & Disimpan
                        </p>
                      </div>
                    ) : (
                      <p className="text-sm text-slate-600 font-medium">
                        Drag & Drop berkas .docx Anda kemari, atau{' '}
                        <label className="text-blue-600 hover:underline cursor-pointer font-bold">
                          pilih file dari explorer
                          <input
                            type="file"
                            accept=".docx"
                            onChange={handleFileChange}
                            className="hidden"
                          />
                        </label>
                      </p>
                    )}
                    <p className="text-xs text-slate-400 mt-2">Format yang didukung: Hanya Microsoft Word (.docx)</p>
                  </div>
                </div>
              ) : (
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <label className="block text-sm font-bold text-slate-800">Isi Visual Web Template</label>
                      <p className="text-xs text-slate-500">Teks template di bawah akan langsung di-layout menjadi lembar cetak.</p>
                    </div>
                    <div className="flex gap-2">
                      <button
                        type="button"
                        onClick={() => setFormWebContent(PRESET_INVITATION)}
                        className="px-2.5 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-semibold transition"
                      >
                        Preset Undangan Resmi
                      </button>
                      <button
                        type="button"
                        onClick={() => setFormWebContent(PRESET_CERTIFICATE)}
                        className="px-2.5 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-semibold transition"
                      >
                        Preset Sertifikat
                      </button>
                    </div>
                  </div>

                  <textarea
                    rows={12}
                    value={formWebContent}
                    onChange={(e) => setFormWebContent(e.target.value)}
                    placeholder="Mulai ketik layout surat Anda..."
                    className="w-full p-4 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800 font-mono text-sm leading-relaxed"
                  />
                  
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs text-slate-600 flex items-start gap-2">
                    <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                    <div>
                      Gunakan tag seperti <code className="bg-slate-200 px-1 py-0.5 rounded text-blue-700 font-bold font-mono">{`{nama}`}</code> (atau <code className="bg-slate-200 px-1 py-0.5 rounded text-blue-700 font-mono font-bold">{`{nama1}`}</code> & <code className="bg-slate-200 px-1 py-0.5 rounded text-blue-700 font-mono font-bold">{`{nama2}`}</code> jika memilih 2 penerima per halaman), serta variabel unik global yang Anda buat di bawah.
                    </div>
                  </div>
                </div>
              )}

              {/* SECTION 3: PENERIMA */}
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                <div>
                  <h3 className="block text-sm font-bold text-slate-800 mb-1.5">Jumlah Penerima per Halaman</h3>
                  <div className="flex gap-4">
                    <button
                      type="button"
                      onClick={() => setFormRecipientsPerPage(1)}
                      className={`flex-1 py-2.5 rounded-xl border text-sm font-bold transition flex items-center justify-center gap-2 ${
                        formRecipientsPerPage === 1
                          ? 'border-blue-600 bg-blue-50 text-blue-800 ring-2 ring-blue-100'
                          : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <span>1 Penerima per Halaman</span>
                      <span className="text-xs bg-slate-200/80 px-2 py-0.5 rounded-full text-slate-600">Tag: {`{nama}`}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setFormRecipientsPerPage(2)}
                      className={`flex-1 py-2.5 rounded-xl border text-sm font-bold transition flex items-center justify-center gap-2 ${
                        formRecipientsPerPage === 2
                          ? 'border-blue-600 bg-blue-50 text-blue-800 ring-2 ring-blue-100'
                          : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <span>2 Penerima per Halaman</span>
                      <span className="text-xs bg-slate-200/80 px-2 py-0.5 rounded-full text-slate-600">Tag: {`{nama1}`} & {`{nama2}`}</span>
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* PENERIMA 1 */}
                  <div>
                    <label className="block text-sm font-bold text-slate-800 mb-1.5">
                      {formRecipientsPerPage === 2 ? 'Penerima Kiri (Penerima 1)' : 'Daftar Penerima'} 
                      <span className="ml-1.5 font-normal text-xs text-slate-500">
                        ({formRecipients1.split('\n').filter(x => x.trim().length > 0).length} baris)
                      </span>
                    </label>
                    <textarea
                      rows={6}
                      value={formRecipients1}
                      onChange={(e) => setFormRecipients1(e.target.value)}
                      placeholder="Joko&#10;Agung&#10;Budi"
                      className="w-full p-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800 text-sm leading-relaxed"
                    />
                  </div>

                  {/* PENERIMA 2 */}
                  {formRecipientsPerPage === 2 && (
                    <div>
                      <label className="block text-sm font-bold text-slate-800 mb-1.5">
                        Penerima Kanan (Penerima 2)
                        <span className="ml-1.5 font-normal text-xs text-slate-500">
                          ({formRecipients2.split('\n').filter(x => x.trim().length > 0).length} baris)
                        </span>
                      </label>
                      <textarea
                        rows={6}
                        value={formRecipients2}
                        onChange={(e) => setFormRecipients2(e.target.value)}
                        placeholder="Andi&#10;Budi&#10;Joni"
                        className="w-full p-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-800 text-sm leading-relaxed"
                      />
                    </div>
                  )}
                </div>

                <p className="text-xs text-slate-400 italic">
                  *Penerima dipisahkan dengan baris baru (newline). Lembaran halaman otomatis digenerate mengikuti baris paling banyak.
                </p>
              </div>

              {/* SECTION 4: CUSTOM FIELDS */}
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-bold text-slate-800">Mendefinisikan Variabel Kustom (Global)</h3>
                    <p className="text-xs text-slate-500">Variabel berikut akan menggantikan tag kustom yang berlaku global di seluruh cetakan.</p>
                  </div>
                  <button
                    type="button"
                    onClick={handleAddCustomField}
                    className="flex items-center gap-1 text-xs bg-blue-50 hover:bg-blue-100 text-blue-700 font-bold px-3 py-1.5 rounded-lg transition"
                  >
                    <Plus className="w-3.5 h-3.5" /> Tambah Field
                  </button>
                </div>

                <div className="space-y-3">
                  {formCustomFields.length === 0 ? (
                    <div className="text-center py-4 bg-slate-50 rounded-xl text-xs text-slate-500 font-medium border border-dashed border-slate-200">
                      Belum ada variabel global kustom didefinisikan.
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 gap-2.5">
                      {formCustomFields.map((f) => (
                        <div key={f.id} className="flex gap-2 items-center">
                          <div className="relative flex-1">
                            <span className="absolute left-3 top-2.5 text-xs font-bold text-slate-400 font-mono">{`{`}</span>
                            <input
                              type="text"
                              value={f.name}
                              onChange={(e) => handleUpdateCustomFieldName(f.id, e.target.value)}
                              placeholder="nama_variabel"
                              className="w-full pl-6 pr-6 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 text-sm font-mono text-slate-800"
                            />
                            <span className="absolute right-3 top-2.5 text-xs font-bold text-slate-400 font-mono">{`}`}</span>
                          </div>
                          
                          <ChevronRight className="w-4 h-4 text-slate-300 shrink-0" />
                          
                          <input
                            type="text"
                            value={f.value}
                            onChange={(e) => handleUpdateCustomFieldValue(f.id, e.target.value)}
                            placeholder="Nilai yang menggantikan..."
                            className="flex-3 px-4 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 text-sm text-slate-800 font-medium"
                          />
                          
                          <button
                            type="button"
                            onClick={() => handleRemoveCustomField(f.id)}
                            className="p-2 text-slate-400 hover:text-red-500 rounded-lg hover:bg-slate-100 transition shrink-0"
                            title="Hapus field"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* SAVE ACTION BAR */}
              <div className="flex items-center justify-end gap-3 border-t border-slate-200 pt-5">
                <button
                  type="button"
                  onClick={() => setIsEditing(false)}
                  className="px-5 py-2.5 rounded-xl border border-slate-300 text-slate-700 hover:bg-slate-50 text-sm font-bold transition"
                >
                  Kembali
                </button>
                <button
                  type="button"
                  onClick={handleSaveForm}
                  className="px-6 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold shadow-md shadow-blue-100 transition flex items-center gap-2"
                >
                  <Check className="w-4 h-4" />
                  Simpan & Generate Preview
                </button>
              </div>

            </div>
          ) : (
            /* LAYER RECORD DETAIL DAN GENERATOR PREVIEW */
            <div className="flex-1 flex flex-col md:flex-row h-full">
              
              {!currentDetails ? (
                /* EMPTY STATE */
                <div className="flex-1 flex flex-col items-center justify-center p-8 text-center max-w-xl mx-auto">
                  <div className="w-16 h-16 bg-blue-50 border border-blue-200 text-blue-600 rounded-3xl flex items-center justify-center mb-6 shadow-md shadow-blue-50">
                    <Sparkles className="w-8 h-8" />
                  </div>
                  <h3 className="text-xl font-bold text-slate-900 tracking-tight mb-2">Belum Ada Template Terseleksi</h3>
                  <p className="text-sm text-slate-500 font-medium leading-relaxed mb-6">
                    Mulai dengan membuat template baru dari sidebar untuk mendesain surat visual web atau mengunggah berkas Microsoft Word (.docx).
                  </p>
                  <button
                    onClick={handleAddNewTemplate}
                    className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl text-sm font-bold shadow-md shadow-blue-100 transition"
                  >
                    <Plus className="w-4 h-4" />
                    Buat Template Pertama Anda
                  </button>
                </div>
              ) : (
                /* SELEKSI SUKSES: DASHBOARD DESIGN DETAILES */
                <div className="flex-1 flex flex-col md:flex-row print:flex-col w-full">
                  
                  {/* PANEL DATA/SETTINGS */}
                  <div className="w-full md:w-[400px] bg-white border-b md:border-b-0 md:border-r border-slate-200 p-6 flex flex-col gap-5 shrink-0 print:hidden overflow-y-auto max-h-[85vh] md:max-h-[calc(100vh-80px)]">
                    
                    <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                      <div>
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          currentDetails.templateType === 'docx' 
                            ? 'bg-emerald-50 text-emerald-700' 
                            : 'bg-amber-50 text-amber-700'
                        }`}>
                          {currentDetails.templateType.toUpperCase()} TEMPLATE
                        </span>
                        <h2 className="text-lg font-bold text-slate-900 leading-snug mt-1">{currentDetails.title}</h2>
                      </div>
                      <div className="flex gap-1.5 shrink-0">
                        <button
                          onClick={() => handleEditTemplate(currentDetails)}
                          className="p-2 border border-slate-200 hover:bg-slate-50 rounded-xl text-slate-600 transition"
                          title="Edit data ini"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
                        <button
                          onClick={(e) => handleDeleteTemplate(currentDetails.id, e)}
                          className="p-2 border border-slate-200 hover:border-red-200 hover:bg-red-50 rounded-xl text-slate-500 hover:text-red-600 transition"
                          title="Hapus template"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>

                    {/* STATS */}
                    <div className="grid grid-cols-2 gap-3.5 bg-slate-50 p-4 rounded-xl border border-slate-200">
                      <div>
                        <div className="text-xs text-slate-400 font-semibold uppercase">Total Halaman</div>
                        <div className="text-lg font-bold text-slate-800">{previewMaxLen} Lembar</div>
                      </div>
                      <div>
                        <div className="text-xs text-slate-400 font-semibold uppercase">Penerima / Page</div>
                        <div className="text-lg font-bold text-slate-800">{currentDetails.recipientsPerPage} Orang</div>
                      </div>
                    </div>

                    {/* PRATINJAU VARIABEL GLOBAL */}
                    <div className="space-y-2.5">
                      <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">Variabel Global Terdeteksi</h4>
                      {currentDetails.customFields.length === 0 ? (
                        <p className="text-xs text-slate-400 italic">Tidak ada variabel global kustom.</p>
                      ) : (
                        <div className="space-y-1.5 max-h-40 overflow-y-auto">
                          {currentDetails.customFields.map((f, idx) => (
                            <div key={idx} className="flex justify-between items-center text-xs p-2 bg-slate-50 border border-slate-200 rounded-lg">
                              <span className="font-mono font-bold text-blue-700">{`{${f.name}}`}</span>
                              <span className="text-slate-700 font-medium truncate max-w-[200px]">{f.value}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* PENERIMA LIST SIDE */}
                    <div className="space-y-2.5 flex-1 min-h-[200px] flex flex-col">
                      <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">Navigasi Penerima</h4>
                      <div className="flex-1 overflow-y-auto space-y-1 bg-slate-50/50 p-2 border border-slate-200 rounded-xl max-h-64 md:max-h-none">
                        {Array.from({ length: previewMaxLen }).map((_, idx) => {
                          const isActive = previewPage === idx;
                          return (
                            <div
                              key={idx}
                              onClick={() => setPreviewPage(idx)}
                              className={`p-2.5 rounded-lg cursor-pointer transition text-xs font-medium flex items-center justify-between border ${
                                isActive
                                  ? 'bg-blue-600 text-white border-blue-600'
                                  : 'hover:bg-slate-100/80 text-slate-700 bg-white border-slate-100'
                              }`}
                            >
                              <div className="truncate">
                                <span className={isActive ? 'text-blue-100' : 'text-slate-400'}>Halaman {idx + 1}:</span>
                                {currentDetails.recipientsPerPage === 1 ? (
                                  <span className="ml-1 px-1 py-0.5 rounded font-extrabold truncate inline-block">
                                    {previewR1[idx] || '-'}
                                  </span>
                                ) : (
                                  <span className="ml-1 px-1 py-0.5 rounded font-semibold truncate inline-block">
                                    {previewR1[idx] || '-'} & {previewR2[idx] || '-'}
                                  </span>
                                )}
                              </div>
                              <ChevronRight className={`w-3.5 h-3.5 shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* UTILITY GENERATOR TRIGGERS */}
                    <div className="pt-4 border-t border-slate-100 space-y-2 shrink-0">
                      {currentDetails.templateType === 'docx' && (
                        <button
                          onClick={handleGeneratedocx}
                          className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold p-3 rounded-xl transition text-sm shadow-md shadow-emerald-50"
                        >
                          <Download className="w-4 h-4" />
                          Download Gabungan DOCX
                        </button>
                      )}
                      
                      <button
                        onClick={handlePrintDocument}
                        className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-slate-800 text-white font-bold p-3 rounded-xl transition text-sm shadow-md shadow-slate-200"
                      >
                        <Printer className="w-4 h-4" />
                        Print / Simpan PDF Halaman
                      </button>
                      
                      <p className="text-[10px] text-center text-slate-400 mt-1">
                        *Gunakan CTRL+P untuk menyimpan semua halaman sebagai file PDF di dialog printer bawaan.
                      </p>
                    </div>

                  </div>

                  {/* PANEL LIVE PREVIEW LAYOUT */}
                  <div className="flex-1 bg-slate-100/80 flex flex-col p-6 items-center overflow-y-auto max-h-[85vh] md:max-h-[calc(100vh-80px)] print:bg-white print:p-0 print:max-h-none print:overflow-visible">
                    
                    {generationError && (
                      <div className="w-full max-w-2xl bg-red-50 border border-red-200 rounded-xl p-4 mb-4 flex items-start gap-2.5 text-red-800 text-xs font-medium print:hidden">
                        <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                        <div>{generationError}</div>
                      </div>
                    )}

                    {generationSuccess && (
                      <div className="w-full max-w-2xl bg-emerald-50 border border-emerald-200 rounded-xl p-3.5 mb-4 flex items-center gap-2.5 text-emerald-800 text-xs font-bold print:hidden">
                        <Check className="w-4 h-4 shrink-0" />
                        <span>Sistem diperbarui! Perubahan berhasil disimpan ke database.</span>
                      </div>
                    )}

                    <div className="w-full max-w-2xl flex items-center justify-between mb-4 print:hidden">
                      <div className="text-xs text-slate-500 font-bold">
                        PRATINJAU LEMBAR KE-{previewPage + 1} DARI {previewMaxLen}
                      </div>
                      <div className="flex gap-1.5">
                        <button
                          disabled={previewPage === 0}
                          onClick={() => setPreviewPage(p => Math.max(0, p - 1))}
                          className="px-2.5 py-1 rounded bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-bold disabled:opacity-50 disabled:cursor-not-allowed transition"
                        >
                          Sebelumnya
                        </button>
                        <button
                          disabled={previewPage >= previewMaxLen - 1}
                          onClick={() => setPreviewPage(p => Math.min(previewMaxLen - 1, p + 1))}
                          className="px-2.5 py-1 rounded bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-bold disabled:opacity-50 disabled:cursor-not-allowed transition"
                        >
                          Berikutnya
                        </button>
                      </div>
                    </div>

                    {/* LIVE PAPER DESIGN SHEET */}
                    {currentDetails.templateType === 'web' ? (
                      /* WEB TEMPLATE LAYOUT PREVIEW */
                      <div 
                        id="web-paper-preview" 
                        className="bg-white w-full max-w-[21cm] min-h-[29.7cm] shadow-lg border border-slate-250 p-[2.5cm] flex flex-col font-sans relative flex-1 text-slate-900 leading-relaxed font-sans print:shadow-none print:border-none print:p-0 print:m-0"
                      >
                        {/* Jika recipients per page adalah 2, kita format dengan 2 kolom kiri & kanan */}
                        {currentDetails.recipientsPerPage === 2 ? (
                          <div className="grid grid-cols-2 gap-10 flex-1">
                            {/* KOLOM KIRI (PENERIMA 1) */}
                            <div className="border-r border-slate-100 pr-5 flex flex-col whitespace-pre-wrap text-sm leading-relaxed justify-between">
                              <div>
                                {replacePlaceholders(currentDetails.webContent, getPageVariables(currentDetails, previewPage))}
                              </div>
                              <div className="text-[10px] text-slate-300 select-none text-center pt-4 border-t border-slate-50">
                                Lembar Mail Merge • Penerima Sisi Kiri
                              </div>
                            </div>
                            {/* KOLOM KANAN (PENERIMA 2) */}
                            <div className="flex flex-col whitespace-pre-wrap text-sm leading-relaxed justify-between">
                              <div>
                                {replacePlaceholders(currentDetails.webContent, {
                                  ...getPageVariables(currentDetails, previewPage),
                                  nama: previewR2[previewPage] || '-', // override generic nama to nama2
                                })}
                              </div>
                              <div className="text-[10px] text-slate-300 select-none text-center pt-4 border-t border-slate-50">
                                Lembar Mail Merge • Penerima Sisi Kanan
                              </div>
                            </div>
                          </div>
                        ) : (
                          /* JIKA 1 PENERIMA PER PAGE */
                          <div className="whitespace-pre-wrap text-sm leading-relaxed flex-1 flex flex-col justify-between">
                            <div>
                              {replacePlaceholders(currentDetails.webContent, getPageVariables(currentDetails, previewPage))}
                            </div>
                            <div className="text-[10px] text-slate-300 select-none text-right pt-6 border-t border-slate-50 print:hidden">
                              Halaman {previewPage + 1} dari {previewMaxLen}
                            </div>
                          </div>
                        )}
                      </div>
                    ) : (
                      /* DOCX TEMPLATE LAYOUT PREVIEW */
                      <div 
                        id="docx-paper-preview" 
                        className="w-full flex-1 flex flex-col items-center print:hidden"
                      >
                        {isRenderingDocx && (
                          <div className="flex-1 flex flex-col items-center justify-center p-12 text-slate-400 gap-2 min-h-[400px]">
                            <RefreshCw className="w-8 h-8 animate-spin text-blue-600" />
                            <span className="text-xs font-semibold">Menghubungkan variabel & merender Word...</span>
                          </div>
                        )}
                        <div 
                          ref={docxPreviewContainerRef} 
                          className={isRenderingDocx ? "hidden" : "w-full flex justify-center"} 
                        />
                        {!isRenderingDocx && (
                          <div className="w-full max-w-[21cm] text-[10px] text-slate-400 select-none text-right pt-4 border-t border-slate-150 print:hidden mt-2">
                            Halaman {previewPage + 1} dari {previewMaxLen} (Pratinjau Word Asli)
                          </div>
                        )}
                      </div>
                    )}

                    {/* BLOK PRINT KHUSUS MEDIA YANG DISINGKIRKAN SAAT SCREEN LAYOUT TAMPIL, NAMUN TERLIHAT DI PRINTER */}
                    <div id="print-area" className="hidden print:block whitespace-pre-wrap">
                      {currentDetails.templateType === 'web' ? (
                        Array.from({ length: previewMaxLen }).map((_, pageIdx) => {
                          const pageVars = getPageVariables(currentDetails, pageIdx);
                          
                          if (currentDetails.recipientsPerPage === 2) {
                            return (
                              <div 
                                key={pageIdx} 
                                className="print-page bg-white w-[21cm] min-h-[29.7cm] p-[2cm] grid grid-cols-2 gap-10 leading-relaxed text-slate-900"
                                style={{ pageBreakAfter: 'always', boxSizing: 'border-box' }}
                              >
                                {/ * SISI KIRI * /}
                                <div className="border-r border-slate-100 pr-5 flex flex-col justify-between">
                                  <div className="text-[12px] whitespace-pre-wrap leading-relaxed">
                                    {replacePlaceholders(currentDetails.webContent, pageVars)}
                                  </div>
                                </div>
                                {/ * SISI KANAN * /}
                                <div className="flex flex-col justify-between">
                                  <div className="text-[12px] whitespace-pre-wrap leading-relaxed">
                                    {replacePlaceholders(currentDetails.webContent, {
                                      ...pageVars,
                                      nama: previewR2[pageIdx] || '-',
                                    })}
                                  </div>
                                </div>
                              </div>
                            );
                          } else {
                            return (
                              <div 
                                key={pageIdx} 
                                className="print-page bg-white w-[21cm] min-h-[29.7cm] p-[2cm] leading-relaxed text-slate-900"
                                style={{ pageBreakAfter: 'always', boxSizing: 'border-box' }}
                              >
                                <div className="text-[12px] whitespace-pre-wrap leading-relaxed">
                                  {replacePlaceholders(currentDetails.webContent, pageVars)}
                                </div>
                              </div>
                            );
                          }
                        })
                      ) : (
                        <div ref={docxPrintContainerRef} className="w-full" />
                      )}
                    </div>

                  </div>

                </div>
              )}

            </div>
          )}

        </main>

      </div>

      {/* MODAL PANDUAN WORD TEMPLATE */}
      {showHelpModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 transition print:hidden">
          <div className="bg-white rounded-3xl max-w-xl w-full p-6 shadow-2xl border border-slate-100 max-h-[90vh] overflow-y-auto space-y-4">
            
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <FileText className="w-5 h-5 text-emerald-600" />
                <h3 className="font-extrabold text-slate-900 text-lg">Panduan Mailing MS Word (.docx)</h3>
              </div>
              <button
                onClick={() => setShowHelpModal(false)}
                className="p-1 px-2.5 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-600 text-sm font-bold transition"
              >
                Tutup
              </button>
            </div>

            <div className="space-y-4 text-sm text-slate-700 leading-relaxed font-sans">
              <p>
                Fitur ini membiarkan Anda mendesain format surat di Microsoft Word (.docx), mengunggahnya, lalu melipatgandakan data penerima secara otomatis tanpa terikat konfigurasi rumit di MS Word.
              </p>

              <div className="space-y-3">
                <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider">Langkah-langkah Desain di Word:</h4>
                <ol className="list-decimal list-inside space-y-2.5 pl-1">
                  <li>
                    <span className="font-semibold text-slate-900">Mulai Loop Template</span>: Di bagian paling awal (atas) dokumen Word, ketik tag pembuka: <code className="bg-slate-100 text-red-600 px-1 rounded font-mono font-bold">{`{#penerima_list}`}</code>.
                  </li>
                  <li>
                    <span className="font-semibold text-slate-900">Gunakan Tag Pengganti</span>: Tulis kalimat surat Anda dan letakkan tag statik global atau dinamis seperti <code className="bg-slate-100 text-blue-600 px-1 rounded font-mono font-semibold">{`{nama}`}</code> (untuk 1 orang per page) atau <code className="bg-slate-100 text-blue-600 px-1 rounded font-mono font-semibold">{`{nama1}`}</code> & <code className="bg-slate-100 text-blue-600 px-1 rounded font-mono font-semibold">{`{nama2}`}</code> (untuk 2 orang per page) di posisi pencetakan nama.
                  </li>
                  <li>
                    <span className="font-semibold text-slate-900">Sisipkan Page Break</span>: Masukkan fitur Page Break bawaan di Word (<kbd className="bg-slate-100 border text-slate-600 px-1 py-0.5 rounded text-xs select-none">Ctrl + Enter</kbd>) tepat setelah tulisan surat Anda selesai agar tiap penerima terpisah halamannya.
                  </li>
                  <li>
                    <span className="font-semibold text-slate-900">Tutup Loop Template</span>: Di baris setelah Page Break tersebut, ketik tag penutup: <code className="bg-slate-100 text-red-600 px-1 rounded font-mono font-bold">{`{/penerima_list}`}</code>.
                  </li>
                </ol>
              </div>

              <div className="bg-blue-50 p-3.5 rounded-xl border border-blue-150 space-y-2 text-blue-900 text-xs">
                <div className="font-bold flex items-center gap-1">
                  <Info className="w-4 h-4 shrink-0" />
                  <span>Mengapa Harus Loop?</span>
                </div>
                Jika Anda tidak menyisipkan tag pembuka <code className="font-mono">{`{#penerima_list}`}</code> dan penutup <code className="font-mono">{`{/penerima_list}`}</code> di Word, template hanya akan digenerate sebagai satu lembar tunggal dengan nama penerima pertama saja (tidak ter-mail merge otomatis).
              </div>

              <div className="bg-slate-55 bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs">
                <span className="font-bold text-slate-800">Contoh Tag Global Kustom:</span><br />
                Jika Anda membuat variabel bernama <code className="font-mono text-emerald-800 font-bold">{`acara`}</code> atau <code className="font-mono text-emerald-800 font-bold">{`tanggal`}</code> di formulir edit, ketik <code className="bg-white px-1 border rounded text-emerald-700 font-mono">{`{acara}`}</code> dan <code className="bg-white px-1 border rounded text-emerald-700 font-mono">{`{tanggal}`}</code> di dalam Word untuk menggantikannya.
              </div>
            </div>

            <div className="border-t border-slate-100 pt-3 flex justify-end">
              <button
                onClick={() => setShowHelpModal(false)}
                className="px-5 py-2.5 rounded-xl bg-slate-900 text-white font-bold hover:bg-slate-800 transition"
              >
                Saya Mengerti
              </button>
            </div>

          </div>
        </div>
      )}

      {/* TAMBAHAN PRINT CSS DI SCREEN UNTUK PRESISI CETAKAN */}
      <style>{`
        /* PRATINJAU DOKUMEN WORD (DOCX) */
        .docx-preview-content {
          font-family: 'Inter', system-ui, sans-serif;
          font-size: 13.5px;
          color: #1e293b;
          line-height: 1.6;
        }

        /* Overrides custom untuk docx-preview library visual rendering di mode live preview screen */
        .docx-preview .docx-wrapper {
          background-color: transparent !important;
          padding: 0 !important;
          font-family: inherit !important;
          display: flex !important;
          justify-content: center !important;
        }
        .docx-preview .docx-wrapper > section.docx {
          box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1) !important;
          border: 1px solid #e2e8f0 !important;
          border-radius: 12px !important;
          background-color: white !important;
          box-sizing: border-box !important;
          margin: 10px auto !important;
        }

        @media print {
          /* CSS PRINT ADJUSTMENT */
          body, html {
            background: white !important;
            color: black !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          #root {
            padding: 0 !important;
            margin: 0 !important;
          }
          aside, header, main > div > div:first-child, .print:hidden, button, .print\\:hidden {
            display: none !important;
          }
          #print-area {
            display: block !important;
            visibility: visible !important;
            width: 100% !important;
          }
          .print-page {
            page-break-after: always !important;
            page-break-inside: avoid !important;
            margin: 0 !important;
            padding: 2cm !important;
            min-height: 29.7cm !important;
            width: 21cm !important;
            box-sizing: border-box !important;
            background: white !important;
          }

          /* Print formatting untuk high-fidelity cetakan docx-preview */
          .docx-print .docx-wrapper {
            padding: 0 !important;
            background: transparent !important;
            box-shadow: none !important;
            margin: 0 !important;
            display: block !important;
          }
          .docx-print .docx-wrapper > section.docx {
            margin: 0 !important;
            box-shadow: none !important;
            border: none !important;
            border-radius: 0 !important;
            page-break-before: auto !important;
            page-break-after: always !important;
            page-break-inside: avoid !important;
            width: 100% !important;
            box-sizing: border-box !important;
            background: white !important;
          }
        }
      `}</style>

    </div>
  );
}
