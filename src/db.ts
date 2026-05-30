/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface CustomField {
  id: string;
  name: string;
  value: string;
}

export interface DocTemplateRecord {
  id: string;
  title: string;
  templateType: 'docx' | 'web';
  docxFileName: string | null;
  docxFileData: string | null; // stored as base64 string
  webContent: string; // fallback web template text/HTML
  recipientsPerPage: 1 | 2;
  recipients1: string; // newline separated names
  recipients2: string; // newline separated names (used if recipientsPerPage === 2)
  customFields: CustomField[];
  createdAt: number;
}

const DB_NAME = 'DocuMergeDB';
const DB_VERSION = 1;
const STORE_NAME = 'templates';

export function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => {
      reject(request.error);
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onupgradeneeded = (event) => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
  });
}

async function saveIndexedDB(record: DocTemplateRecord): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    const request = store.put(record);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

async function deleteIndexedDB(id: string): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// REST API + Fallback IndexedDB
export async function getAllTemplates(): Promise<DocTemplateRecord[]> {
  try {
    const response = await fetch('/api/templates');
    if (!response.ok) {
      throw new Error('API response not ok');
    }
    const data = await response.json();
    return data;
  } catch (err) {
    console.warn('Gagal terkoneksi ke database server, beralih ke database lokal browser:', err);
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.getAll();
      request.onsuccess = () => {
        const result = request.result as DocTemplateRecord[];
        result.sort((a, b) => b.createdAt - a.createdAt);
        resolve(result);
      };
      request.onerror = () => {
        reject(request.error);
      };
    });
  }
}

export async function getTemplateById(id: string): Promise<DocTemplateRecord | null> {
  try {
    const response = await fetch('/api/templates');
    if (!response.ok) {
      throw new Error('API response not ok');
    }
    const data: DocTemplateRecord[] = await response.json();
    const found = data.find(t => t.id === id);
    return found || null;
  } catch (err) {
    console.warn('Gagal terkoneksi ke database server untuk getTemplateById:', err);
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(id);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  }
}

export async function saveTemplate(record: DocTemplateRecord): Promise<void> {
  try {
    const response = await fetch('/api/templates', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(record),
    });
    if (!response.ok) {
      throw new Error('Gagal menyimpan di server');
    }
    // Backup locally to indexeddb
    try {
      await saveIndexedDB(record);
    } catch (_) {}
  } catch (err) {
    console.warn('Gagal menyimpan ke database server, mencadangkan ke browser lokal:', err);
    return saveIndexedDB(record);
  }
}

export async function deleteTemplate(id: string): Promise<void> {
  try {
    const response = await fetch(`/api/templates/${id}`, {
      method: 'DELETE',
    });
    if (!response.ok) {
      throw new Error('Gagal menghapus di server');
    }
    try {
      await deleteIndexedDB(id);
    } catch (_) {}
  } catch (err) {
    console.warn('Gagal menghapus dari database server, menghapus dari lokal browser:', err);
    return deleteIndexedDB(id);
  }
}
