import React, { useState } from 'react';
import { FileSpreadsheet, Copy, Check, RefreshCw, CheckCircle2, AlertCircle, Code2, ShieldCheck } from 'lucide-react';
import { getSettings, updateSettings } from '../../services/storage';
import { syncToGoogleSheets, APPS_SCRIPT_TEMPLATE } from '../../services/sheetsSync';

export const GoogleSheetsSettings: React.FC = () => {
  const [settings, setSettings] = useState(getSettings());
  const [scriptUrl, setScriptUrl] = useState(settings.sheetsConfig.webAppUrl || settings.sheetsConfig.scriptUrl || '');
  const [spreadsheetId, setSpreadsheetId] = useState(settings.sheetsConfig.spreadsheetId || '');
  const [isTesting, setIsTesting] = useState(false);
  const [isCopied, setIsCopied] = useState(false);
  const [showCode, setShowCode] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleCopyCode = async () => {
    await navigator.clipboard.writeText(APPS_SCRIPT_TEMPLATE);
    setIsCopied(true);
    window.setTimeout(() => setIsCopied(false), 1800);
  };

  const handleSaveConfig = () => {
    const url = scriptUrl.trim();
    const updated = updateSettings({ sheetsConfig: { ...settings.sheetsConfig, webAppUrl: url, scriptUrl: url, spreadsheetId: spreadsheetId.trim(), isConnected: Boolean(url), syncStatus: 'idle', syncError: undefined } });
    setSettings(updated);
    setTestResult({ success: true, message: 'Web App URL saved on this device.' });
  };

  const handleTestConnection = async () => {
    const url = scriptUrl.trim();
    if (!url) {
      setTestResult({ success: false, message: 'Please enter your Google Apps Script Web App URL first.' });
      return;
    }
    const updated = updateSettings({ sheetsConfig: { ...getSettings().sheetsConfig, webAppUrl: url, scriptUrl: url, spreadsheetId: spreadsheetId.trim(), isConnected: false, syncStatus: 'idle', syncError: undefined } });
    setSettings(updated);
    setIsTesting(true);
    setTestResult(null);
    try {
      const result = await syncToGoogleSheets(true);
      setTestResult(result);
      setSettings(getSettings());
    } finally {
      setIsTesting(false);
    }
  };

  return (
    <div id="google_sheets_settings_container" className="space-y-4">
      <div className="bg-white dark:bg-slate-900 p-4 rounded-2xl border border-slate-200/80 dark:border-slate-800 shadow-xs space-y-2">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-2.5"><div className="p-2 rounded-xl bg-emerald-600 text-white"><FileSpreadsheet className="w-5 h-5" /></div><div><h3 className="text-sm font-bold">Google Sheets Real-Time Sync</h3><p className="text-xs text-slate-500">Sync entries through Google Apps Script</p></div></div>
          <span className={`text-[10px] font-bold px-2.5 py-1 rounded-full uppercase ${settings.sheetsConfig.isConnected ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}`}>{settings.sheetsConfig.isConnected ? 'Connected' : 'Not Connected'}</span>
        </div>
      </div>
      <div className="bg-white dark:bg-slate-900 p-4 rounded-2xl border border-slate-200/80 dark:border-slate-800 space-y-3.5">
        <h4 className="text-xs font-bold">Web App URL Configuration</h4>
        <div className="space-y-1"><label className="text-xs font-semibold">Google Apps Script Web App URL *</label><input type="url" value={scriptUrl} onChange={(e) => setScriptUrl(e.target.value)} placeholder="https://script.google.com/macros/s/AKfycb.../exec" autoCapitalize="none" autoCorrect="off" spellCheck={false} className="w-full px-3.5 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-xs font-mono" /><p className="text-[11px] text-slate-400">Use the complete deployed URL ending in /exec.</p></div>
        <div className="flex items-center space-x-2 pt-1"><button type="button" onClick={handleSaveConfig} className="px-4 py-2 bg-slate-800 text-white rounded-xl text-xs font-bold active:opacity-70 touch-manipulation">Save URL</button><button type="button" disabled={!scriptUrl.trim() || isTesting} onClick={handleTestConnection} className="px-4 py-2 bg-emerald-600 text-white rounded-xl text-xs font-bold disabled:opacity-50 flex items-center gap-1.5 touch-manipulation">{isTesting ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <CheckCircle2 className="w-3.5 h-3.5" />}<span>{isTesting ? 'Testing...' : 'Test & Sync Now'}</span></button></div>
        {testResult && <div className={`p-3 rounded-xl text-xs font-semibold flex items-center gap-2 ${testResult.success ? 'bg-emerald-50 border border-emerald-300 text-emerald-800' : 'bg-rose-50 border border-rose-300 text-rose-800'}`}>{testResult.success ? <CheckCircle2 className="w-4 h-4" /> : <AlertCircle className="w-4 h-4" />}<span>{testResult.message}</span></div>}
      </div>
      <div className="bg-white dark:bg-slate-900 p-4 rounded-2xl border border-slate-200/80 dark:border-slate-800 space-y-3"><div className="flex items-center justify-between"><h4 className="text-xs font-bold flex items-center gap-1.5"><ShieldCheck className="w-4 h-4 text-emerald-600" />How to setup in Google Sheets</h4><button type="button" onClick={handleCopyCode} className="px-3 py-1 bg-indigo-50 text-indigo-600 rounded-lg text-xs font-bold flex items-center gap-1 touch-manipulation">{isCopied ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}{isCopied ? 'Code Copied!' : 'Copy Script Code'}</button></div><ol className="list-decimal list-inside space-y-2 text-xs text-slate-600 leading-relaxed"><li>Open a Google Sheet.</li><li>Extensions &gt; Apps Script.</li><li>Replace the code and Save.</li><li>Deploy &gt; New deployment &gt; Web app.</li><li>Execute as Me; Who has access: Anyone.</li><li>Copy the Web App URL ending in /exec and paste it above.</li></ol><button type="button" onClick={() => setShowCode(!showCode)} className="text-xs font-semibold text-indigo-600 flex items-center gap-1 touch-manipulation"><Code2 className="w-3.5 h-3.5" />{showCode ? 'Hide Apps Script Source Code' : 'View Apps Script Source Code'}</button>{showCode && <pre className="mt-2 p-3 bg-slate-950 text-emerald-400 font-mono text-[11px] rounded-xl overflow-x-auto max-h-60">{APPS_SCRIPT_TEMPLATE}</pre>}</div>
    </div>
  );
};
