import { Capacitor, CapacitorHttp } from '@capacitor/core';

import {
  getEntries,
  getCategories,
  getTabs,
  getAffirmations,
  getDiaryEntries,
  getSettings,
  saveSettings,
  addSyncLog,
  getPendingSyncQueue,
  clearPendingSyncQueue,
} from './storage';

export interface SyncResponse {
  success: boolean;
  message: string;
  syncedRecordsCount?: number;
  timestamp: string;
  error?: string;
}

export const APPS_SCRIPT_TEMPLATE = `/**
 * Google Apps Script for "My Daily Life Manager"
 * Instructions:
 * 1. Open your Google Sheet
 * 2. Click Extensions > Apps Script
 * 3. Replace all code with this script
 * 4. Click Save
 * 5. Deploy > New deployment > Web app
 * 6. Execute as: Me
 * 7. Who has access: Anyone
 * 8. Copy the Web app URL into My Daily Life Manager Settings
 */

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "My Daily Life Manager Google Sheets endpoint is running",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!e || !e.postData || !e.postData.contents) throw new Error("No POST data received");
    var data = JSON.parse(e.postData.contents);
    ensureSheets(ss);

    if (data.entries && data.entries.length > 0) {
      syncSheet(ss, "All Entries", [
        "ID", "Date", "Time", "Module", "Type", "Title", "Category", "Subcategory",
        "Priority", "Status", "Amount", "Payment Mode", "Person / Vendor", "Organization",
        "Due Date", "Ref / File No", "Tags", "Description", "Archived", "Created At", "Updated At"
      ], data.entries.map(function(item) {
        return [
          item.id, item.date, item.time || "", item.moduleId, item.type, item.title,
          item.categoryName || "", item.subcategory || "", item.priority || "", item.status || "",
          item.amount !== undefined ? item.amount : "", item.paymentMode || "",
          item.vendorOrPerson || item.person || "", item.organization || "",
          item.dueDate || "", item.referenceNumber || item.fileNumber || "",
          (item.tags || []).join(", "), item.description || "",
          item.isArchived ? "Yes" : "No", item.createdAt, item.updatedAt
        ];
      }));
      syncModuleSheet(ss, "Office", data.entries.filter(function(x) { return x.moduleId === 'office'; }));
      syncModuleSheet(ss, "Union", data.entries.filter(function(x) { return x.moduleId === 'union'; }));
      syncModuleSheet(ss, "Personal", data.entries.filter(function(x) { return x.moduleId === 'personal'; }));
      syncModuleSheet(ss, "Home", data.entries.filter(function(x) { return x.moduleId === 'home'; }));
      syncModuleSheet(ss, "Expenses", data.entries.filter(function(x) { return x.moduleId === 'expense' || x.type === 'Expense'; }));
      syncModuleSheet(ss, "Credits", data.entries.filter(function(x) { return x.moduleId === 'credit' || x.type === 'Credit / Income'; }));
      syncModuleSheet(ss, "Bills", data.entries.filter(function(x) { return x.moduleId === 'bills' || x.type === 'Bill'; }));
      syncModuleSheet(ss, "Tasks", data.entries.filter(function(x) { return x.moduleId === 'tasks' || x.type === 'Task'; }));
      syncModuleSheet(ss, "Notes", data.entries.filter(function(x) { return x.moduleId === 'notes' || x.type === 'Note'; }));
      syncModuleSheet(ss, "Events", data.entries.filter(function(x) { return x.moduleId === 'events' || x.type === 'Event'; }));
      syncModuleSheet(ss, "Wishlist", data.entries.filter(function(x) { return x.moduleId === 'wishlist'; }));
    }

    if (data.categories && data.categories.length > 0) {
      syncSheet(ss, "Categories", ["ID", "Module", "Name", "Subcategories", "Archived", "Order"],
        data.categories.map(function(c) { return [c.id, c.moduleId, c.name, (c.subcategories || []).join(", "), c.isArchived ? "Yes" : "No", c.order]; }));
    }

    if (data.diary && data.diary.length > 0) {
      syncSheet(ss, "Daily Summary", ["Date", "Mood", "Highlight", "Rating (1-5)", "Daily Summary", "Updated At"],
        data.diary.map(function(d) { return [d.date, d.mood || "", d.highlight || "", d.rating || "", d.summaryText || "", d.updatedAt]; }));
    }

    updateDashboardSummary(ss, data);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Synced " + (data.entries ? data.entries.length : 0) + " records to Google Sheets successfully",
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", message: err.toString(), timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function ensureSheets(ss) {
  ["Dashboard Summary", "All Entries", "Office", "Union", "Personal", "Home", "Expenses", "Credits", "Bills", "Tasks", "Notes", "Events", "Wishlist", "Affirmations", "Daily Summary", "Categories", "Custom Fields", "Sync Log"].forEach(function(name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
}

function syncSheet(ss, sheetName, headers, rows) {
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  if (rows && rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

function syncModuleSheet(ss, sheetName, items) {
  if (!items || items.length === 0) return;
  syncSheet(ss, sheetName, ["ID", "Date", "Time", "Title", "Category", "Subcategory", "Amount", "Person / Vendor", "Status", "Priority", "Due Date", "Description"], items.map(function(item) {
    return [item.id, item.date, item.time || "", item.title, item.categoryName || "", item.subcategory || "", item.amount !== undefined ? item.amount : "", item.vendorOrPerson || item.person || "", item.status || "", item.priority || "", item.dueDate || "", item.description || ""];
  }));
}

function updateDashboardSummary(ss, data) {
  var sheet = ss.getSheetByName("Dashboard Summary") || ss.insertSheet("Dashboard Summary");
  sheet.clear();
  var entries = data.entries || [];
  var totalExpenses = entries.filter(function(e) { return e.moduleId === 'expense' || e.type === 'Expense'; }).reduce(function(a, e) { return a + (e.amount || 0); }, 0);
  var totalCredits = entries.filter(function(e) { return e.moduleId === 'credit' || e.type === 'Credit / Income'; }).reduce(function(a, e) { return a + (e.amount || 0); }, 0);
  [["MY DAILY LIFE MANAGER - DASHBOARD SUMMARY", ""], ["Last Synced", new Date().toLocaleString()], ["", ""], ["Metric", "Value"], ["Total Entries", entries.length], ["Total Income / Credits", totalCredits], ["Total Expenses", totalExpenses], ["Net Difference (Credits - Expenses)", totalCredits - totalExpenses], ["Pending Tasks", entries.filter(function(e) { return (e.moduleId === 'tasks' || e.type === 'Task') && e.status !== 'completed'; }).length], ["Completed Tasks", entries.filter(function(e) { return (e.moduleId === 'tasks' || e.type === 'Task') && e.status === 'completed'; }).length], ["Office Activities", entries.filter(function(e) { return e.moduleId === 'office'; }).length], ["Union Activities", entries.filter(function(e) { return e.moduleId === 'union'; }).length], ["Home Activities", entries.filter(function(e) { return e.moduleId === 'home'; }).length], ["Personal Activities", entries.filter(function(e) { return e.moduleId === 'personal'; }).length]].forEach(function(row) { sheet.appendRow(row); });
  sheet.getRange("A1:B1").merge().setFontWeight("bold").setFontSize(14).setBackground("#0f172a").setFontColor("#ffffff");
  sheet.getRange("A4:B4").setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
  sheet.autoResizeColumns(1, 2);
}
`;

export async function syncToGoogleSheets(isTest: boolean = false): Promise<SyncResponse> {
  const settings = getSettings();
  const { webAppUrl, spreadsheetId } = settings.sheetsConfig;
  if (!webAppUrl) return { success: false, message: 'Google Apps Script Web App URL is not configured. Please paste your deployed Web App URL in Settings.', timestamp: new Date().toISOString(), error: 'Missing Web App URL' };

  const entries = getEntries();
  const categories = getCategories();
  const tabs = getTabs();
  const affirmations = getAffirmations();
  const diary = getDiaryEntries();
  const payload = { action: isTest ? 'test_connection' : 'full_sync', spreadsheetId, timestamp: new Date().toISOString(), entries, categories, tabs, affirmations, diary, pendingQueue: getPendingSyncQueue() };

  saveSettings({ sheetsConfig: { ...settings.sheetsConfig, syncStatus: 'syncing', syncError: undefined } });
  try {
    if (!navigator.onLine) throw new Error('Device is currently offline. Changes will automatically sync when connection returns.');
    let result: any;
    let httpStatus = 200;

    if (Capacitor.isNativePlatform()) {
      const nativeResponse = await CapacitorHttp.post({
        url: webAppUrl,
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        data: JSON.stringify(payload),
        responseType: 'json',
        disableRedirects: false,
        connectTimeout: 20000,
        readTimeout: 60000,
      });
      httpStatus = nativeResponse.status;
      result = nativeResponse.data;
    } else {
      const response = await fetch(webAppUrl, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload), redirect: 'follow' });
      httpStatus = response.status;
      result = await response.json().catch(() => ({ status: response.ok ? 'success' : 'error', message: response.ok ? 'Sync request dispatched successfully' : `HTTP ${response.status}` }));
    }

    if (httpStatus < 200 || httpStatus >= 300) throw new Error(`Google Apps Script returned HTTP ${httpStatus}`);
    if (result && (result.status === 'error' || result.success === false)) throw new Error(result.message || result.error || 'Google Apps Script rejected the sync request');

    const now = new Date().toISOString();
    clearPendingSyncQueue();
    saveSettings({ sheetsConfig: { ...settings.sheetsConfig, isConnected: true, lastSynced: now, syncStatus: 'synced', syncError: undefined } });
    addSyncLog({ operation: isTest ? 'Test Connection' : 'Full Sync', status: 'success', details: `Synced ${entries.length} entries & categories to Google Sheets (${result?.message || 'OK'})` });
    return { success: true, message: isTest ? 'Google Sheets Connection Verified Successfully!' : `Synchronized ${entries.length} records to Google Sheets!`, syncedRecordsCount: entries.length, timestamp: now };
  } catch (err: any) {
    const errorMsg = err?.message || 'Sync failed';
    const now = new Date().toISOString();
    saveSettings({ sheetsConfig: { ...settings.sheetsConfig, syncStatus: 'error', syncError: errorMsg } });
    addSyncLog({ operation: isTest ? 'Test Connection' : 'Full Sync', status: 'failed', details: errorMsg });
    return { success: false, message: errorMsg, timestamp: now, error: errorMsg };
  }
}

export function setupAutoSync(): () => void {
  const handler = () => {
    const settings = getSettings();
    if (settings.sheetsConfig.autoSync && settings.sheetsConfig.webAppUrl) syncToGoogleSheets(false).catch(console.error);
  };
  window.addEventListener('online', handler);
  return () => window.removeEventListener('online', handler);
}
