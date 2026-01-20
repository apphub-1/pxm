
import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { 
  Plus, 
  Trash2,
  Search, 
  Edit2, 
  Save, 
  X, 
  Database,
  RefreshCw,
  CheckCircle,
  Network,
  Lock,
  Key,
  Shield,
  Link,
  AlertTriangle,
  Wifi,
  WifiOff,
  Settings,
  ChevronDown,
  Info,
  Link2,
  ShieldCheck,
  FileCheck,
  FileDown,
  Server,
  User,
  Users,
  LogIn,
  LogOut,
  History as HistoryIcon,
  Clock,
  FileText,
  ChevronRight,
  HelpCircle,
  CheckSquare,
  Square,
  Download,
  Rocket,
  FileCog,
  Layers,
  KeyRound,
  ClipboardList,
  Cloud,
  FileSpreadsheet,
  Moon,
  Sun
} from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import './index.css';
import '@fontsource/plus-jakarta-sans/400.css';
import '@fontsource/plus-jakarta-sans/500.css';
import '@fontsource/plus-jakarta-sans/600.css';
import '@fontsource/plus-jakarta-sans/700.css';
import '@fontsource/plus-jakarta-sans/800.css';
import '@fontsource/jetbrains-mono/400.css';

type DataRow = Record<string, any>;
const API_BASE_URL = '/api';

import { login, logout, getAccessToken, getUser, getRole } from './auth';

const DEFAULT_HEADERS = [
  "ICTO", 
  "Name", 
  "Objektstatus", 
  "Stereotyp", 
  "tAV", 
  "Kritikalität", 
  "Anmerkung zur Variante", 
  "Anbindungsvariante", 
  "PAM-Relevanz",
  "Protokollierung privilegierter Rechte auf Anwendungsebene",
  "Protokollierung priviligierter Rechte für DB auf Server",
  "Protokollierung priviligierter Rechte für Betriebssystem auf Server",
  "Schnittstellendokument", 
  "Workorder Abnahme", 
  "Entzug privilegierte Berechtigungen",
  "AbnahmePAMOnboarding"
];

const SELECT_OPTIONS: Record<string, string[]> = {
  "Objektstatus": ["Plan", "Aktiv", "Stillgelegt"],
  "Stereotyp": ["ExternalSystem_ASPSAAS", "InternalSystem", "CloudService"],
  "Kritikalität": ["1-niedrig", "2-mittel", "3-hoch", "4-kritisch"],
  "Anbindungsvariante": ["Variante 1", "Variante 2", "Variante 3", "Sonderlösung"],
  "PAM-Relevanz": ["Ja", "Nein"],
  "Protokollierung privilegierter Rechte auf Anwendungsebene": [
    "Keine PAM-Anbindung - keine privilegierten Rechte",
    "Keine PAM-Anbindung - Ausnahme gemäß Ausnahmeprozess",
    "keine PAM-Anbindung - Externe Anwendung",
    "Protokollierung privilegierter Rechte über CyberArk"
  ],
  "Protokollierung priviligierter Rechte für DB auf Server": [
    "keine PAM-Anbindung für den eingesetzten Datenbanktyp",
    "Keine PAM-Anbindung - keine privilegierten Rechte",
    "Protokollierung privilegierter Rechte über CyberArk",
    "keine PAM-Anbindung - Externe Anwendung"
  ],
  "Protokollierung priviligierter Rechte für Betriebssystem auf Server": [
    "keine PAM-Anbindung für den eingesetzten Betriebssystemtyp",
    "Keine PAM-Anbindung - keine privilegierten Rechte",
    "Protokollierung privilegierter Rechte über CyberArk",
    "keine PAM-Anbindung - Externe Anwendung"
  ],
  "Schnittstellendokument": ["Vorhanden", "Nicht vorhanden", "In Arbeit"],
  "Entzug privilegierte Berechtigungen": ["erfolgt", "in Arbeit", "offen"]
};

const FIELD_GROUPS = [
  {
    title: "Stammdaten",
    icon: <Info className="w-5 h-5 text-blue-500" />,
    fields: ["ICTO", "Name", "Kurzname", "Objektstatus", "Stereotyp", "Kritikalität", "tAV", "Stellvertreter tAV", "fAV", "Betriebsverantwortlicher", "Objektpflege"]
  },
  {
    title: "Anbindung & Relevanz",
    icon: <Link2 className="w-5 h-5 text-indigo-500" />,
    fields: ["Anbindungsvariante", "PAM-Relevanz", "Anmerkung zur Variante"]
  },
  {
    title: "Protokollierung (Compliance)",
    icon: <ShieldCheck className="w-5 h-5 text-emerald-500" />,
    fields: [
      "Protokollierung privilegierter Rechte auf Anwendungsebene",
      "Protokollierung priviligierter Rechte für DB auf Server",
      "Protokollierung priviligierter Rechte für Betriebssystem auf Server"
    ]
  },
  {
    title: "Dokumentation & Abschluss",
    icon: <FileCheck className="w-5 h-5 text-amber-500" />,
    fields: ["Schnittstellendokument", "Workorder Abnahme", "Entzug privilegierte Berechtigungen"]
  }
];

// --- Onboarding Types & Components ---

type OnboardingData = Record<string, any>;

const ONBOARDING_SECTIONS = [
  { id: 'architecture', title: '1. Architektur', icon: <Database className="w-5 h-5" /> },
  { id: 'technology', title: '2. Anwendungstechnologie', icon: <Settings className="w-5 h-5" /> },
  { id: 'login', title: '3. Anmeldeverfahren', icon: <LogIn className="w-5 h-5" /> },
  { id: 'password', title: '4. Passwortwechsel', icon: <ShieldCheck className="w-5 h-5" /> },
  { id: 'test', title: '5. Testuser und Testumgebung', icon: <FileCheck className="w-5 h-5" /> },
  { id: 'emergency', title: '6. Notfallprozess', icon: <AlertTriangle className="w-5 h-5" /> },
  { id: 'matrix', title: '7. Vereinbarte Anbindungsvariante', icon: <Link2 className="w-5 h-5" /> },
  { id: 'bypass', title: '8. PAM-Bypass Regeln', icon: <Shield className="w-5 h-5" /> },
];

const TECHNICAL_SECTIONS = [
  { id: 'servers', title: '1. Server / Betriebssysteme', icon: <Server className="w-5 h-5" /> },
  { id: 'databases', title: '2. Datenbanken / Server', icon: <Database className="w-5 h-5" /> },
  { id: 'ports', title: '3. Portfreischaltungen', icon: <Network className="w-5 h-5" /> },
  { id: 'safes', title: '4. Safe-Struktur (CyberArk)', icon: <Lock className="w-5 h-5" /> },
  { id: 'safeMembers', title: '5. Mitglieder der Safes', icon: <Users className="w-5 h-5" /> },
  { id: 'sharedAccounts', title: '6. Shared Accounts', icon: <Key className="w-5 h-5" /> },
  { id: 'permissions', title: '7. Berechtigungszuordnungen', icon: <Shield className="w-5 h-5" /> },
  { id: 'mapping', title: '8. Shared Accounts zu Safe', icon: <Link className="w-5 h-5" /> },
];

const SECRETS_SECTIONS = [
  { id: 'inventory', title: '1. Secret Inventory', icon: <KeyRound className="w-5 h-5" /> },
  { id: 'safes', title: '2. Safe- / Pfad-Struktur', icon: <Lock className="w-5 h-5" /> },
  { id: 'members', title: '3. Mitglieder der Safes', icon: <Users className="w-5 h-5" /> },
  { id: 'mapping', title: '4. Secrets zu Safe', icon: <Link className="w-5 h-5" /> },
];

const SECRETS_ONBOARDING_SECTIONS = [
  { id: 'infrastructure', title: '1. Infrastruktur', icon: <Server className="w-5 h-5" /> },
  { id: 'tools', title: '2. Tool-Nutzung', icon: <Settings className="w-5 h-5" /> },
  { id: 'container_standalone', title: '3. Container – Standalone', icon: <Database className="w-5 h-5" /> },
  { id: 'container_k8s', title: '4. Container – Kubernetes / OpenShift', icon: <Cloud className="w-5 h-5" /> },
  { id: 'cloud', title: '5. Cloud Computing', icon: <Cloud className="w-5 h-5" /> },
  { id: 'server', title: '6. Server', icon: <Server className="w-5 h-5" /> },
  { id: 'properties', title: '7. Eigenschaften der Secrets', icon: <Key className="w-5 h-5" /> },
  { id: 'rotation_status', title: '8. Rotation – Ist-Status', icon: <RefreshCw className="w-5 h-5" /> },
  { id: 'target_image', title: '9. Anbindungsvariante (Zielbild)', icon: <Link2 className="w-5 h-5" /> },
];

const UserPicker = ({ label, value, onChange, tooltip, className, readOnly }: any) => {
  const [query, setQuery] = useState(value || '');
  const [results, setResults] = useState<any[]>([]);
  const [isOpen, setIsOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const searchTimeout = React.useRef<any>(null);
  const abortController = React.useRef<AbortController | null>(null);

  useEffect(() => {
    setQuery(value || '');
  }, [value]);

  useEffect(() => {
    const handleClickOutside = (event: any) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [wrapperRef]);

  const handleSearch = (val: string) => {
    setQuery(val);
    onChange(val); 
    
    if (searchTimeout.current) clearTimeout(searchTimeout.current);
    if (abortController.current) abortController.current.abort(); // Alte Anfrage abbrechen

    if (val.length < 3) {
        setResults([]);
        setIsOpen(false);
        setLoading(false);
        return;
    }

    searchTimeout.current = setTimeout(async () => {
      setLoading(true);
      const controller = new AbortController();
      abortController.current = controller;

      try {
          const token = getAccessToken();
          const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
          const res = await fetch(`${API_BASE_URL}/directory/search?q=${encodeURIComponent(val)}`, { headers, signal: controller.signal });
          if (res.ok) {
              const data = await res.json();
              setResults(data);
              setIsOpen(true);
          }
      } catch (e) {
          if ((e as Error).name !== 'AbortError') console.error("Search failed:", e);
      } finally {
          if (abortController.current === controller) setLoading(false);
      }
    }, 300);
  };

  const selectUser = (u: any) => {
      const display = `${u.displayName} (${u.username})`;
      setQuery(display);
      onChange(display);
      setIsOpen(false);
  };

  return (
    <div className={`relative ${className || 'mb-4'}`} ref={wrapperRef}>
      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1 flex items-center gap-2">
        {label}
        {tooltip && <div className="group relative"><HelpCircle className="w-3 h-3 text-slate-400 cursor-help" /><div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 p-2 bg-slate-800 text-white text-xs rounded hidden group-hover:block z-50">{tooltip}</div></div>}
      </label>
      <div className="relative">
        <input 
            type="text" 
            className="w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium pr-10 dark:text-slate-200"
            value={query}
            onChange={e => handleSearch(e.target.value)}
            onFocus={() => query.length >= 3 && setIsOpen(true)}
            placeholder="Name oder Kennung suchen..."
            disabled={readOnly}
            maxLength={500}
        />
        <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
            {loading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <User className="w-4 h-4" />}
        </div>
      </div>
      
      {isOpen && results.length > 0 && (
          <div className="absolute z-50 left-0 right-0 mt-1 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg shadow-xl max-h-60 overflow-y-auto custom-scrollbar">
              {results.map((u: any) => (
                  <div 
                    key={u.username} 
                    className="p-3 hover:bg-indigo-50 dark:hover:bg-indigo-900/30 cursor-pointer border-b border-slate-50 dark:border-slate-700 last:border-0"
                    onClick={() => selectUser(u)}
                  >
                      <div className="font-bold text-sm text-slate-800 dark:text-slate-200">{u.displayName}</div>
                      <div className="text-xs text-slate-500 dark:text-slate-400 flex gap-2">
                          <span>{u.username}</span>
                          {u.email && <span>• {u.email}</span>}
                      </div>
                  </div>
              ))}
          </div>
      )}
    </div>
  );
};

const MultiUserPicker = ({ label, value, onChange, tooltip, className, readOnly }: any) => {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<any[]>([]);
  const [isOpen, setIsOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const searchTimeout = React.useRef<any>(null);
  const abortController = React.useRef<AbortController | null>(null);

  const selectedUsers = useMemo(() => {
      return value ? value.split(';').map((s: string) => s.trim()).filter(Boolean) : [];
  }, [value]);

  useEffect(() => {
    const handleClickOutside = (event: any) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [wrapperRef]);

  const handleSearch = (val: string) => {
    setQuery(val);
    
    if (searchTimeout.current) clearTimeout(searchTimeout.current);
    if (abortController.current) abortController.current.abort(); // Alte Anfrage abbrechen

    if (val.length < 3) {
        setResults([]);
        setIsOpen(false);
        setLoading(false);
        return;
    }
    
    searchTimeout.current = setTimeout(async () => {
      setLoading(true);
      const controller = new AbortController();
      abortController.current = controller;

      try {
          const token = getAccessToken();
          const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
          const res = await fetch(`${API_BASE_URL}/directory/search?q=${encodeURIComponent(val)}`, { headers, signal: controller.signal });
          if (res.ok) {
              const data = await res.json();
              setResults(data);
              setIsOpen(true);
          }
      } catch (e) {
          if ((e as Error).name !== 'AbortError') console.error("Search failed:", e);
      } finally {
          if (abortController.current === controller) setLoading(false);
      }
    }, 300);
  };

  const addUser = (u: any) => {
      const display = `${u.displayName} (${u.username})`;
      if (!selectedUsers.some(u => u.toLowerCase() === display.toLowerCase())) {
          const newUsers = [...selectedUsers, display];
          onChange(newUsers.join('; '));
      }
      setQuery('');
      setIsOpen(false);
  };

  const removeUser = (userToRemove: string) => {
      const newUsers = selectedUsers.filter((u: string) => u !== userToRemove);
      onChange(newUsers.join('; '));
  };

  return (
    <div className={`relative ${className || 'mb-4'}`} ref={wrapperRef}>
      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1 flex items-center gap-2">
        {label}
        {tooltip && <div className="group relative"><HelpCircle className="w-3 h-3 text-slate-400 cursor-help" /><div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 p-2 bg-slate-800 text-white text-xs rounded hidden group-hover:block z-50">{tooltip}</div></div>}
      </label>
      
      <div className="flex flex-wrap gap-2 mb-2">
          {selectedUsers.map((u: string) => (
              <div key={u} className="bg-indigo-50 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-2 py-1 rounded-md text-xs font-bold flex items-center gap-1 border border-indigo-100 dark:border-indigo-800">
                  {u}
                  {!readOnly && <button onClick={() => removeUser(u)} className="hover:text-indigo-900"><X className="w-3 h-3" /></button>}
              </div>
          ))}
      </div>

      <div className="relative">
        <input 
            type="text" 
            className={`w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium pr-10 dark:text-slate-200 ${readOnly ? 'cursor-not-allowed opacity-60' : ''}`}
            value={query}
            onChange={e => handleSearch(e.target.value)}
            onFocus={() => !readOnly && query.length >= 3 && setIsOpen(true)}
            placeholder={readOnly ? "Keine Bearbeitung möglich" : "Benutzer hinzufügen..."}
            disabled={readOnly}
            maxLength={500}
        />
        <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
            {loading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />}
        </div>
      </div>
      
      {isOpen && results.length > 0 && (
          <div className="absolute z-50 left-0 right-0 mt-1 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg shadow-xl max-h-60 overflow-y-auto custom-scrollbar">
              {results.map((u: any) => (
                  <div 
                    key={u.username} 
                    className="p-3 hover:bg-indigo-50 dark:hover:bg-indigo-900/30 cursor-pointer border-b border-slate-50 dark:border-slate-700 last:border-0"
                    onClick={() => addUser(u)}
                  >
                      <div className="font-bold text-sm text-slate-800 dark:text-slate-200">{u.displayName}</div>
                      <div className="text-xs text-slate-500 dark:text-slate-400 flex gap-2">
                          <span>{u.username}</span>
                          {u.email && <span>• {u.email}</span>}
                      </div>
                  </div>
              ))}
          </div>
      )}
    </div>
  );
};

const MultiSelectPicker = ({ value, onChange, options, readOnly }: any) => {
  const [isOpen, setIsOpen] = useState(false);
  const wrapperRef = React.useRef<HTMLDivElement>(null);

  const selectedValues = useMemo(() => {
      return value ? value.split(';').map((s: string) => s.trim()).filter(Boolean) : [];
  }, [value]);

  useEffect(() => {
    const handleClickOutside = (event: any) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [wrapperRef]);

  const toggleValue = (val: string) => {
      let newValues;
      if (selectedValues.includes(val)) {
          newValues = selectedValues.filter((v: string) => v !== val);
      } else {
          newValues = [...selectedValues, val];
      }
      onChange(newValues.join('; '));
  };

  return (
    <div className="relative w-full" ref={wrapperRef}>
      <div 
        className={`w-full p-2 bg-white dark:bg-slate-800 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 min-h-[38px] flex flex-wrap gap-1 items-center cursor-pointer transition-all duration-200 ease-in-out ${readOnly ? 'opacity-60 cursor-not-allowed' : ''} border-slate-200 dark:border-slate-700`}
        onClick={() => !readOnly && setIsOpen(!isOpen)}
      >
        {selectedValues.length === 0 && <span className="text-slate-400 text-sm">Wählen...</span>}
        {selectedValues.map((v: string) => (
            <span key={v} className="bg-indigo-50 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-1.5 py-0.5 rounded text-xs font-bold border border-indigo-100 dark:border-indigo-800">{v}</span>
        ))}
        <div className="ml-auto"><ChevronDown className="w-3 h-3 text-slate-400" /></div>
      </div>
      
      {isOpen && (
          <div className="mt-1 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg shadow-sm max-h-60 overflow-y-auto custom-scrollbar z-50 absolute w-full">
              {options.map((opt: string) => (
                  <div 
                    key={opt} 
                    className={`p-2 hover:bg-indigo-50 dark:hover:bg-indigo-900/30 cursor-pointer border-b border-slate-50 dark:border-slate-700 last:border-0 text-sm flex items-center gap-2 ${selectedValues.includes(opt) ? 'bg-indigo-50/50 dark:bg-indigo-900/20 text-indigo-700 dark:text-indigo-300 font-bold' : 'text-slate-700 dark:text-slate-300'}`}
                    onClick={() => toggleValue(opt)}
                  >
                      {selectedValues.includes(opt) ? <CheckSquare className="w-4 h-4" /> : <Square className="w-4 h-4 text-slate-300" />}
                      {opt}
                  </div>
              ))}
          </div>
      )}
    </div>
  );
};

const PaginatedHistoryList = ({ history, cardClassName = "bg-white dark:bg-slate-800 p-5 rounded-lg border border-slate-100 dark:border-slate-700 shadow-sm" }: { history: any[], cardClassName?: string }) => {
  const [page, setPage] = useState(1);
  const pageSize = 10;
  
  useEffect(() => {
      setPage(1);
  }, [history]);

  const totalPages = Math.ceil(history.length / pageSize);
  const paginatedHistory = history.slice((page - 1) * pageSize, page * pageSize);

  if (history.length === 0) {
      return <p className="text-slate-400 pl-10">Keine Änderungen protokolliert.</p>;
  }

  return (
    <div>
        <div className="space-y-4 relative before:absolute before:left-[19px] before:top-2 before:bottom-2 before:w-0.5 before:bg-slate-100 dark:before:bg-slate-700">
            {paginatedHistory.map((entry: any) => (
                <div key={entry.id} className="relative pl-10">
                    <div className="absolute left-0 top-1.5 w-[40px] h-[40px] bg-white dark:bg-slate-800 border-4 border-slate-50 dark:border-slate-900 rounded-full flex items-center justify-center z-10">
                        <div className={`w-3 h-3 rounded-full ${entry.action === 'ERSTELLT' ? 'bg-emerald-400' : 'bg-indigo-400'}`}></div>
                    </div>
                    <div className={cardClassName}>
                        <div className="flex justify-between items-start mb-2">
                            <div>
                                <span className="block font-bold text-slate-800 dark:text-slate-200">{entry.username}</span>
                                <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">{entry.action}</span>
                            </div>
                            <span className="text-xs font-mono text-slate-400 bg-slate-50 dark:bg-slate-900 px-2 py-1 rounded-md border border-slate-100 dark:border-slate-700">
                                {new Date(entry.timestamp).toLocaleString()}
                            </span>
                        </div>
                        {(() => {
                            try {
                                const details = JSON.parse(entry.details);
                                if (Array.isArray(details)) {
                                    return (
                                        <div className="mt-3 space-y-1 border-t border-slate-100 dark:border-slate-700 pt-3">
                                            {details.map((change: any, i: number) => (
                                                <div key={i} className="text-xs flex flex-wrap items-center gap-x-2 gap-y-1 text-slate-600 dark:text-slate-400">
                                                    <span className="font-bold bg-slate-50 dark:bg-slate-900 px-1.5 py-0.5 rounded-sm border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300">{change.field}</span>
                                                    <span className="line-through opacity-50 decoration-rose-400/50">{change.old || <span className="italic text-[10px]">leer</span>}</span>
                                                    <span className="text-slate-300">➜</span>
                                                    <span className="font-bold text-indigo-600 dark:text-indigo-400">{change.new || <span className="italic text-[10px]">leer</span>}</span>
                                                </div>
                                            ))}
                                        </div>
                                    );
                                }
                            } catch (e) {}
                            return null;
                        })()}
                    </div>
                </div>
            ))}
        </div>
        {totalPages > 1 && (
            <div className="flex justify-center items-center gap-4 mt-8 pl-10">
                <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page === 1} className="px-4 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 font-bold text-xs disabled:opacity-50 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors">
                    Neuere
                </button>
                <span className="text-xs font-bold text-slate-400">Seite {page} von {totalPages}</span>
                <button onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page === totalPages} className="px-4 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 font-bold text-xs disabled:opacity-50 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors">
                    Ältere
                </button>
            </div>
        )}
    </div>
  );
};

const UnifiedAppModal = ({ 
  isOpen, 
  onClose, 
  governanceRow, 
  user,
  initialTab = 'onboarding',
  mode = 'pam',
  readOnly = false
}: { 
  isOpen: boolean; 
  onClose: () => void; 
  governanceRow: DataRow; 
  user: string | null;
  initialTab?: 'onboarding' | 'technical' | 'secrets' | 'secrets-onboarding';
  mode?: 'pam' | 'secrets';
  readOnly?: boolean;
}) => {
  const [activeTab, setActiveTab] = useState<'onboarding' | 'technical' | 'secrets' | 'secrets-onboarding'>(initialTab);

  // --- Onboarding State ---
  const [onboardingData, setOnboardingData] = useState<OnboardingData>({});
  const [onboardingOpenSections, setOnboardingOpenSections] = useState<Record<string, boolean>>(
    ONBOARDING_SECTIONS.reduce((acc, s) => ({ ...acc, [s.id]: true }), {})
  );
  const [onboardingLoading, setOnboardingLoading] = useState(false);
  const [onboardingSaveStatus, setOnboardingSaveStatus] = useState<'idle' | 'saving' | 'success'>('idle');
  const [onboardingViewMode, setOnboardingViewMode] = useState<'form' | 'history'>('form');
  const [onboardingHistory, setOnboardingHistory] = useState<any[]>([]);
  const [onboardingLoaded, setOnboardingLoaded] = useState(false);

  // --- Technical State ---
  const [technicalData, setTechnicalData] = useState<Record<string, any[]>>({
    servers: [], databases: [], ports: [], safes: [], safeMembers: [], sharedAccounts: [], permissions: [], mapping: []
  });
  const [technicalOpenSections, setTechnicalOpenSections] = useState<Record<string, boolean>>(
    TECHNICAL_SECTIONS.reduce((acc, s) => ({ ...acc, [s.id]: true }), {})
  );
  const [technicalLoading, setTechnicalLoading] = useState(false);
  const [technicalSaveStatus, setTechnicalSaveStatus] = useState<'idle' | 'saving' | 'success'>('idle');
  const [technicalViewMode, setTechnicalViewMode] = useState<'form' | 'history'>('form');
  const [technicalHistory, setTechnicalHistory] = useState<any[]>([]);
  const [technicalLoaded, setTechnicalLoaded] = useState(false);

  // --- Secrets State ---
  const [secretsData, setSecretsData] = useState<Record<string, any[]>>({
    inventory: [], safes: [], members: [], mapping: []
  });
  const [secretsOpenSections, setSecretsOpenSections] = useState<Record<string, boolean>>(
    SECRETS_SECTIONS.reduce((acc, s) => ({ ...acc, [s.id]: true }), {})
  );
  const [secretsLoading, setSecretsLoading] = useState(false);
  const [secretsSaveStatus, setSecretsSaveStatus] = useState<'idle' | 'saving' | 'success'>('idle');
  const [secretsViewMode, setSecretsViewMode] = useState<'form' | 'history'>('form');
  const [secretsHistory, setSecretsHistory] = useState<any[]>([]);
  const [secretsLoaded, setSecretsLoaded] = useState(false);

  // --- Secrets Onboarding State ---
  const [secretsOnboardingData, setSecretsOnboardingData] = useState<Record<string, any>>({});
  const [secretsOnboardingOpenSections, setSecretsOnboardingOpenSections] = useState<Record<string, boolean>>(
    SECRETS_ONBOARDING_SECTIONS.reduce((acc, s) => ({ ...acc, [s.id]: true }), {})
  );
  const [secretsOnboardingLoading, setSecretsOnboardingLoading] = useState(false);
  const [secretsOnboardingSaveStatus, setSecretsOnboardingSaveStatus] = useState<'idle' | 'saving' | 'success'>('idle');
  const [secretsOnboardingViewMode, setSecretsOnboardingViewMode] = useState<'form' | 'history'>('form');
  const [secretsOnboardingHistory, setSecretsOnboardingHistory] = useState<any[]>([]);
  const [secretsOnboardingLoaded, setSecretsOnboardingLoaded] = useState(false);
  const saveLock = React.useRef(false);

  useEffect(() => {
    if (isOpen && governanceRow.id) {
        if (activeTab === 'onboarding' && !onboardingLoaded) {
            loadOnboardingData();
        } else if (activeTab === 'technical' && !technicalLoaded) {
            loadTechnicalData();
        } else if (activeTab === 'secrets' && !secretsLoaded) {
            loadSecretsData();
        } else if (activeTab === 'secrets-onboarding' && !secretsOnboardingLoaded) {
            loadSecretsOnboardingData();
        }
    }
  }, [activeTab, isOpen, governanceRow, onboardingLoaded, technicalLoaded, secretsLoaded, secretsOnboardingLoaded]);

  useEffect(() => {
    const { matrixLogin, matrixPwChange } = onboardingData;
    if (matrixLogin && matrixPwChange) {
      let variant = '';
      if (matrixLogin === 'Automatisch' && matrixPwChange === 'Automatisch') {
        variant = 'Anbindungsvariante 3';
      } else if (matrixLogin === 'Manuell' && matrixPwChange === 'Manuell') {
        variant = 'Anbindungsvariante 1';
      } else {
        variant = 'Anbindungsvariante 2';
      }
      
      if (onboardingData.selectedVariant !== variant) {
        setOnboardingData(prev => ({ ...prev, selectedVariant: variant }));
      }
    }
  }, [onboardingData.matrixLogin, onboardingData.matrixPwChange]);

  useEffect(() => {
    if (!technicalData.mapping) return;
    
    let hasChanges = false;
    const newMapping = technicalData.mapping.map(row => {
        const newRow = { ...row };
        let rowChanged = false;

        if (newRow.techAccount) {
            const account = technicalData.sharedAccounts?.find(a => a.techName === newRow.techAccount);
            if (account && newRow.bizAccount !== account.bizName) {
                newRow.bizAccount = account.bizName;
                rowChanged = true;
            }
        }

        if (newRow.techSafe) {
            const safe = technicalData.safes?.find(s => s.techSafeName === newRow.techSafe);
            if (safe && newRow.safeName !== safe.safeName) {
                newRow.safeName = safe.safeName;
                rowChanged = true;
            }
        }
        
        if (rowChanged) hasChanges = true;
        return newRow;
    });

    if (hasChanges) {
        setTechnicalData(prev => ({ ...prev, mapping: newMapping }));
    }
  }, [technicalData.mapping, technicalData.sharedAccounts, technicalData.safes]);

  const loadOnboardingData = async () => {
    setOnboardingLoading(true);
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const res = await fetch(`${API_BASE_URL}/onboarding/${governanceRow.id}`, { headers });
      if (res.ok) {
        const json = await res.json();
        // Pre-fill some data from governance row if empty
        if (!json.data) {
            setOnboardingData({});
        } else {
            setOnboardingData(JSON.parse(json.data));
        }
        setOnboardingLoaded(true);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setOnboardingLoading(false);
    }
  };

  const loadTechnicalData = async () => {
    setTechnicalLoading(true);
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const res = await fetch(`${API_BASE_URL}/technical/${governanceRow.id}`, { headers });
      if (res.ok) {
        const json = await res.json();
        if (json.data) {
            setTechnicalData(JSON.parse(json.data));
        }
        setTechnicalLoaded(true);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setTechnicalLoading(false);
    }
  };

  const loadSecretsData = async () => {
    setSecretsLoading(true);
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const res = await fetch(`${API_BASE_URL}/secrets/${governanceRow.id}`, { headers });
      if (res.ok) {
        const json = await res.json();
        if (json.data) {
            setSecretsData(JSON.parse(json.data));
        }
        setSecretsLoaded(true);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setSecretsLoading(false);
    }
  };

  const loadSecretsOnboardingData = async () => {
    setSecretsOnboardingLoading(true);
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const res = await fetch(`${API_BASE_URL}/secrets-onboarding/${governanceRow.id}`, { headers });
      if (res.ok) {
        const json = await res.json();
        if (json.data) {
            setSecretsOnboardingData(JSON.parse(json.data));
        }
        setSecretsOnboardingLoaded(true);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setSecretsOnboardingLoading(false);
    }
  };

  const handleOnboardingSave = async () => {
    if (saveLock.current) return;
    setOnboardingSaveStatus('saving');
    saveLock.current = true;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;
      
      await fetch(`${API_BASE_URL}/onboarding/${governanceRow.id}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(onboardingData)
      });
      setOnboardingSaveStatus('success');
      setTimeout(() => setOnboardingSaveStatus('idle'), 2000);
    } catch (e) {
      alert("Fehler beim Speichern.");
      setOnboardingSaveStatus('idle');
    } finally {
      saveLock.current = false;
    }
  };

  const handleTechnicalSave = async () => {
    if (saveLock.current) return;
    setTechnicalSaveStatus('saving');
    saveLock.current = true;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;
      
      await fetch(`${API_BASE_URL}/technical/${governanceRow.id}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(technicalData)
      });
      setTechnicalSaveStatus('success');
      setTimeout(() => setTechnicalSaveStatus('idle'), 2000);
    } catch (e) {
      alert("Fehler beim Speichern.");
      setTechnicalSaveStatus('idle');
    } finally {
      saveLock.current = false;
    }
  };

  const handleSecretsSave = async () => {
    if (saveLock.current) return;
    setSecretsSaveStatus('saving');
    saveLock.current = true;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;
      
      await fetch(`${API_BASE_URL}/secrets/${governanceRow.id}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(secretsData)
      });
      setSecretsSaveStatus('success');
      setTimeout(() => setSecretsSaveStatus('idle'), 2000);
    } catch (e) {
      alert("Fehler beim Speichern.");
      setSecretsSaveStatus('idle');
    } finally {
      saveLock.current = false;
    }
  };

  const handleSecretsOnboardingSave = async () => {
    if (saveLock.current) return;
    setSecretsOnboardingSaveStatus('saving');
    saveLock.current = true;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;
      
      await fetch(`${API_BASE_URL}/secrets-onboarding/${governanceRow.id}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(secretsOnboardingData)
      });
      setSecretsOnboardingSaveStatus('success');
      setTimeout(() => setSecretsOnboardingSaveStatus('idle'), 2000);
    } catch (e) {
      alert("Fehler beim Speichern.");
      setSecretsOnboardingSaveStatus('idle');
    } finally {
      saveLock.current = false;
    }
  };

  const loadOnboardingHistory = async () => {
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const response = await fetch(`${API_BASE_URL}/history/${governanceRow.id}`, { headers });
      if (response.ok) {
        const allHistory = await response.json();
        setOnboardingHistory(allHistory.filter((h: any) => h.action.includes('ONBOARDING')));
      }
    } catch (e) {
      console.error("History load failed", e);
    }
  };

  const loadTechnicalHistory = async () => {
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const response = await fetch(`${API_BASE_URL}/history/${governanceRow.id}`, { headers });
      if (response.ok) {
        const allHistory = await response.json();
        setTechnicalHistory(allHistory.filter((h: any) => h.action.includes('TECHNICAL')));
      }
    } catch (e) {
      console.error("History load failed", e);
    }
  };

  const loadSecretsHistory = async () => {
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const response = await fetch(`${API_BASE_URL}/history/${governanceRow.id}`, { headers });
      if (response.ok) {
        const allHistory = await response.json();
        setSecretsHistory(allHistory.filter((h: any) => h.action.includes('SECRETS')));
      }
    } catch (e) {
      console.error("History load failed", e);
    }
  };

  const loadSecretsOnboardingHistory = async () => {
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const response = await fetch(`${API_BASE_URL}/history/${governanceRow.id}`, { headers });
      if (response.ok) {
        const allHistory = await response.json();
        setSecretsOnboardingHistory(allHistory.filter((h: any) => h.action.includes('SECRETS_ONB')));
      }
    } catch (e) {
      console.error("History load failed", e);
    }
  };

  const toggleOnboardingSection = (id: string) => {
    setOnboardingOpenSections(prev => ({ ...prev, [id]: !prev[id] }));
  };

  const toggleTechnicalSection = (id: string) => {
    setTechnicalOpenSections(prev => ({ ...prev, [id]: !prev[id] }));
  };

  const toggleSecretsSection = (id: string) => {
    setSecretsOpenSections(prev => ({ ...prev, [id]: !prev[id] }));
  };

  const toggleSecretsOnboardingSection = (id: string) => {
    setSecretsOnboardingOpenSections(prev => ({ ...prev, [id]: !prev[id] }));
  };

  const updateField = (field: string, value: any) => {
    setOnboardingData(prev => ({ ...prev, [field]: value }));
  };

  const renderInput = (label: string, field: string, type = 'text', tooltip?: string, options?: string[]) => (
    <div className="mb-4">
      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1 flex items-center gap-2">
        {label}
        {tooltip && <div className="group relative"><HelpCircle className="w-3 h-3 text-slate-400 cursor-help" /><div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 p-2 bg-slate-800 text-white text-xs rounded hidden group-hover:block z-50">{tooltip}</div></div>}
      </label>
      {options ? (
        <select 
            className="w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium dark:text-slate-200"
            value={onboardingData[field] || ''}
            onChange={e => updateField(field, e.target.value)}
            disabled={readOnly}
        >
            <option value="">Bitte wählen...</option>
            {options.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : (
        <input 
            type={type} 
            className="w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium dark:text-slate-200"
            value={onboardingData[field] || ''}
            onChange={e => updateField(field, e.target.value)}
            disabled={readOnly}
            maxLength={500}
        />
      )}
    </div>
  );

  const renderCheckbox = (label: string, field: string) => (
      <div className={`mb-4 flex items-center gap-3 ${readOnly ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'}`} onClick={() => !readOnly && updateField(field, !onboardingData[field])}>
          {onboardingData[field] ? <CheckSquare className="w-5 h-5 text-indigo-600" /> : <Square className="w-5 h-5 text-slate-300" />}
          <span className="text-sm font-medium text-slate-700 dark:text-slate-300">{label}</span>
      </div>
  );

  const renderUserPicker = (label: string, field: string, tooltip?: string) => (
      <UserPicker 
        label={label} 
        value={onboardingData[field]} 
        onChange={(val: string) => updateField(field, val)} 
        tooltip={tooltip} 
        readOnly={readOnly}
      />
  );

  const updateSecretsOnboardingField = (field: string, value: any) => {
    setSecretsOnboardingData(prev => ({ ...prev, [field]: value }));
  };

  const updateServerRegistry = (index: number, field: string, value: any) => {
      setSecretsOnboardingData(prev => {
          const list = [...(prev.serverRegistry || [])];
          list[index] = { ...list[index], [field]: value };
          return { ...prev, serverRegistry: list };
      });
  };

  const addServerRegistryRow = () => {
      setSecretsOnboardingData(prev => ({
          ...prev,
          serverRegistry: [...(prev.serverRegistry || []), { segment: 'Produktion' }]
      }));
  };

  const removeServerRegistryRow = (index: number) => {
      setSecretsOnboardingData(prev => {
          const list = [...(prev.serverRegistry || [])];
          list.splice(index, 1);
          return { ...prev, serverRegistry: list };
      });
  };

  const resolveHostname = async (index: number, hostname: string) => {
      if (!hostname) return;
      try {
          const token = getAccessToken();
          const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
          const res = await fetch(`${API_BASE_URL}/dns/resolve?hostname=${encodeURIComponent(hostname)}`, { headers });
          if (res.ok) {
              const { address } = await res.json();
              if (address) {
                  updateServerRegistry(index, 'ip', address);
              }
          }
      } catch (e) {
          console.error(e);
      }
  };

  const renderSecretsOnboardingInput = (label: string, field: string, type = 'text', tooltip?: string, options?: string[]) => (
    <div className="mb-4">
      <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1 flex items-center gap-2">
        {label}
        {tooltip && <div className="group relative"><HelpCircle className="w-3 h-3 text-slate-400 cursor-help" /><div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 p-2 bg-slate-800 text-white text-xs rounded hidden group-hover:block z-50">{tooltip}</div></div>}
      </label>
      {options ? (
        <select 
            className="w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium dark:text-slate-200"
            value={secretsOnboardingData[field] || ''}
            onChange={e => updateSecretsOnboardingField(field, e.target.value)}
            disabled={readOnly}
        >
            <option value="">Bitte wählen...</option>
            {options.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : (
        <input 
            type={type} 
            className="w-full p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm font-medium dark:text-slate-200"
            value={secretsOnboardingData[field] || ''}
            onChange={e => updateSecretsOnboardingField(field, e.target.value)}
            disabled={readOnly}
            maxLength={500}
        />
      )}
    </div>
  );

  const renderSecretsOnboardingMultiSelect = (label: string, field: string, options: string[]) => {
      const currentValues = (secretsOnboardingData[field] as string[]) || [];
      const toggleValue = (val: string) => {
          const newValues = currentValues.includes(val)
              ? currentValues.filter(v => v !== val)
              : [...currentValues, val];
          updateSecretsOnboardingField(field, newValues);
      };

      return (
          <div className="mb-4">
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">{label}</label>
              <div className="space-y-2">
                  {options.map(opt => (
                      <div key={opt} className={`flex items-center gap-2 p-1 rounded ${readOnly ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800'}`} onClick={() => !readOnly && toggleValue(opt)}>
                          {currentValues.includes(opt) ? <CheckSquare className="w-4 h-4 text-indigo-600" /> : <Square className="w-4 h-4 text-slate-300" />}
                          <span className="text-sm text-slate-700 dark:text-slate-300">{opt}</span>
                      </div>
                  ))}
              </div>
          </div>
      );
  };

  const renderMatrix = () => {
      const variants = ['Anbindungsvariante 1', 'Anbindungsvariante 2', 'Anbindungsvariante 3'];
      const criticalities = ['1-niedrig', '2-mittel', '3-hoch', '4-kritisch'];
      const currentCriticality = governanceRow['Kritikalität'];
      
      return (
          <div className="overflow-x-auto">
              <table className="w-full border-collapse text-sm">
                  <thead>
                      <tr>
                          <th className="p-3 border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-left dark:text-slate-300">Kritikalität</th>
                          {variants.map(v => <th key={v} className={`p-3 border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 dark:text-slate-300 ${v.includes('3') ? 'bg-indigo-50 dark:bg-indigo-900/20 text-indigo-700 dark:text-indigo-300' : ''}`}>{v}</th>)}
                      </tr>
                  </thead>
                  <tbody>
                      {criticalities.map(crit => (
                          <tr key={crit}>
                              <td className="p-3 border border-slate-200 dark:border-slate-700 font-bold text-slate-600 dark:text-slate-300">{crit}</td>
                              {variants.map(v => {
                                  const isSelected = onboardingData.selectedVariant === v && currentCriticality === crit;
                                  return (
                                      <td 
                                        key={v} 
                                        className={`p-3 border border-slate-200 dark:border-slate-700 text-center transition-colors ${isSelected ? 'bg-emerald-50 dark:bg-emerald-900/20 ring-2 ring-inset ring-emerald-500' : (currentCriticality === crit ? 'bg-slate-50 dark:bg-slate-800' : '')} ${readOnly ? 'pointer-events-none' : ''}`}
                                      >
                                          {isSelected && <CheckCircle className="w-5 h-5 text-emerald-500 mx-auto" />}
                                      </td>
                                  );
                              })}
                          </tr>
                      ))}
                  </tbody>
              </table>
              <div className="mt-4 p-4 bg-indigo-50 dark:bg-indigo-900/20 rounded-xl border border-indigo-100 dark:border-indigo-800">
                  <p className="text-sm font-bold text-indigo-900 dark:text-indigo-300">Ausgewählte Konfiguration:</p>
                  <p className="text-indigo-700 dark:text-indigo-400">
                    {currentCriticality ? `Kritikalität: ${currentCriticality}` : 'Keine Kritikalität gewählt'} 
                    {' ➜ '} 
                    {onboardingData.selectedVariant || 'Bitte Konfiguration wählen'}
                  </p>
              </div>
          </div>
      );
  };

  const updateTechnicalRow = (section: string, index: number, field: string, value: any) => {
    setTechnicalData(prev => {
      const newList = [...(prev[section] || [])];
      newList[index] = { ...newList[index], [field]: value };
      return { ...prev, [section]: newList };
    });
  };

  const addTechnicalRow = (section: string) => {
    setTechnicalData(prev => ({ ...prev, [section]: [...(prev[section] || []), {}] }));
  };

  const removeTechnicalRow = (section: string, index: number) => {
    setTechnicalData(prev => {
      const newList = [...(prev[section] || [])];
      newList.splice(index, 1);
      return { ...prev, [section]: newList };
    });
  };

  const renderTable = (sectionId: string, columns: { key: string, label: string, type?: string, options?: string[], width?: string, required?: boolean, readOnly?: boolean }[], enableImport = false) => {
    const rows = technicalData[sectionId] || [];
    return (
      <div className="overflow-x-auto">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700">
              {columns.map(col => (
                <th key={col.key} className="p-3 text-left font-bold text-slate-600 dark:text-slate-300 whitespace-nowrap" style={{ minWidth: col.width }}>
                  {col.label} {col.required && <span className="text-rose-500">*</span>}
                </th>
              ))}
              <th className="p-3 w-10"></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-slate-100 dark:border-slate-700 hover:bg-slate-50/50 dark:hover:bg-slate-800/50">
                {columns.map(col => (
                  <td key={col.key} className="p-2" style={{ minWidth: col.width }}>
                    {col.type === 'select' ? (
                      <select 
                        className={`w-full p-2 bg-white dark:bg-slate-800 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200 ${col.required && !row[col.key] ? 'border-rose-300' : 'border-slate-200 dark:border-slate-700'}`}
                        value={row[col.key] || ''}
                        onChange={e => updateTechnicalRow(sectionId, idx, col.key, e.target.value)}
                        disabled={readOnly}
                      >
                        <option value="">-</option>
                        {col.options?.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    ) : col.type === 'multiselect' ? (
                      <MultiSelectPicker 
                        value={row[col.key]} 
                        onChange={(val: string) => updateTechnicalRow(sectionId, idx, col.key, val)}
                        options={col.options || []}
                        readOnly={readOnly}
                      />
                    ) : col.type === 'checkbox' ? (
                      <div className="flex justify-center">
                        <input 
                          type="checkbox" 
                          className="w-5 h-5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                          checked={!!row[col.key]}
                          onChange={e => updateTechnicalRow(sectionId, idx, col.key, e.target.checked)}
                          disabled={readOnly}
                        />
                      </div>
                    ) : (
                      <input 
                        type={col.type || 'text'}
                        className={`w-full p-2 bg-white dark:bg-slate-800 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200 ${col.required && !row[col.key] ? 'border-rose-300' : 'border-slate-200 dark:border-slate-700'} ${col.readOnly ? 'bg-slate-100 dark:bg-slate-900 text-slate-500 dark:text-slate-400' : ''}`}
                        value={row[col.key] || ''}
                        onChange={e => !col.readOnly && updateTechnicalRow(sectionId, idx, col.key, e.target.value)}
                        placeholder={col.label}
                        disabled={readOnly || col.readOnly}
                        maxLength={500}
                      />
                    )}
                  </td>
                ))}
                <td className="p-2 text-center">
                  {!readOnly && (
                  <button onClick={() => removeTechnicalRow(sectionId, idx)} className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-colors">
                    <Trash2 className="w-4 h-4" />
                  </button>
                  )}
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr>
                <td colSpan={columns.length + 1} className="p-4 text-center text-slate-400 italic">Keine Einträge vorhanden.</td>
              </tr>
            )}
          </tbody>
        </table>
        {!readOnly && (
        <button onClick={() => addTechnicalRow(sectionId)} className="mt-3 flex items-center gap-2 text-sm font-bold text-indigo-600 hover:text-indigo-700 px-3 py-2 hover:bg-indigo-50 rounded-lg transition-colors">
          <Plus className="w-4 h-4" /> Zeile hinzufügen
        </button>
        )}
      </div>
    );
  };

  const updateSecretsRow = (section: string, index: number, field: string, value: any) => {
    setSecretsData(prev => {
      const newList = [...(prev[section] || [])];
      const updatedRow = { ...newList[index], [field]: value };

      if (section === 'mapping' && field === 'safeName') {
          const safe = prev.safes?.find((s: any) => s.safeName === value);
          updatedRow.techSafe = safe ? safe.techSafeName : '';
      }

      newList[index] = updatedRow;
      return { ...prev, [section]: newList };
    });
  };

  const addSecretsRow = (section: string) => {
    setSecretsData(prev => ({ ...prev, [section]: [...(prev[section] || []), {}] }));
  };

  const removeSecretsRow = (section: string, index: number) => {
    setSecretsData(prev => {
      const newList = [...(prev[section] || [])];
      newList.splice(index, 1);
      return { ...prev, [section]: newList };
    });
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>, sectionId: string, columns: any[]) => {
      const file = e.target.files?.[0];
      if (!file) return;

      try {
          const buffer = await file.arrayBuffer();
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(buffer);
          
          const worksheet = workbook.worksheets[0];
          if (!worksheet) throw new Error("Kein Arbeitsblatt gefunden");

          const jsonData: any[] = [];
          const headers: string[] = [];

          // Header lesen
          worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
              headers[colNumber] = cell.text ? cell.text.toString().trim() : '';
          });

          // Daten lesen
          worksheet.eachRow((row, rowNumber) => {
              if (rowNumber === 1) return;
              
              const rowData: any = {};
              let hasData = false;
              
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                  const header = headers[colNumber];
                  if (header) {
                      // cell.text ist sicherer für Import als cell.value (vermeidet Formel-Objekte etc.)
                      const val = cell.text ? cell.text.toString().trim() : '';
                      if (val) {
                          rowData[header] = val;
                          hasData = true;
                      }
                  }
              });
              
              if (hasData) jsonData.push(rowData);
          });

          const mappedData = jsonData.map((row: any) => {
              const newRow: any = {};
              columns.forEach(col => {
                  // Match by Label (preferred) or Key
                  let val = row[col.label];
                  if (val === undefined) val = row[col.key];
                  
                  // Case insensitive fallback
                  if (val === undefined) {
                      const keyMatch = Object.keys(row).find(k => k.toLowerCase() === col.label.toLowerCase() || k.toLowerCase() === col.key.toLowerCase());
                      if (keyMatch) val = row[keyMatch];
                  }

                  if (val !== undefined) {
                      newRow[col.key] = String(val);
                  }
              });
              return newRow;
          });

          if (mappedData.length > 0) {
            setSecretsData(prev => ({
                ...prev,
                [sectionId]: [...(prev[sectionId] || []), ...mappedData]
            }));
            alert(`${mappedData.length} Einträge erfolgreich importiert.`);
          }
      } catch (err) {
          console.error(err);
          alert("Fehler beim Importieren der Datei.");
      } finally {
          e.target.value = '';
      }
  };

  const handleDownloadTemplate = async (columns: any[]) => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Template');

      // Spalten definieren und Styling
      worksheet.columns = columns.map(c => ({
          header: c.label,
          key: c.key,
          width: c.width ? parseInt(c.width) / 7 : 25 // Pixel zu Excel-Breite (ungefähr)
      }));

      // Header Styling (Indigo Hintergrund, Weiße Schrift, Fett)
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF4F46E5' } // Indigo-600
      };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.height = 24;

      // Dropdowns (Data Validation) hinzufügen
      columns.forEach((col, index) => {
          if (col.options && col.options.length > 0) {
              const letter = worksheet.getColumn(index + 1).letter;
              // Validierung für Zeilen 2 bis 1000
              for (let i = 2; i <= 1000; i++) {
                  worksheet.getCell(`${letter}${i}`).dataValidation = {
                      type: 'list',
                      allowBlank: true,
                      formulae: [`"${col.options.join(',')}"`]
                  };
              }
          }
      });

      // Datei schreiben und herunterladen
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = "Secret_Inventory_Template.xlsx";
      a.click();
      URL.revokeObjectURL(url);
  };

  const renderSecretsTable = (sectionId: string, columns: { key: string, label: string, type?: string, options?: string[], width?: string, required?: boolean }[], enableImport = false) => {
    const rows = secretsData[sectionId] || [];
    return (
      <div className="overflow-x-auto">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700">
              {columns.map(col => (
                <th key={col.key} className="p-3 text-left font-bold text-slate-600 dark:text-slate-300 whitespace-nowrap">
                  {col.label} {col.required && <span className="text-rose-500">*</span>}
                </th>
              ))}
              <th className="p-3 w-10"></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-slate-100 dark:border-slate-700 hover:bg-slate-50/50 dark:hover:bg-slate-800/50">
                {columns.map(col => (
                  <td key={col.key} className="p-2" style={{ minWidth: col.width }}>
                    {col.type === 'select' ? (
                      <select 
                        className={`w-full p-2 bg-white dark:bg-slate-800 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200 ${col.required && !row[col.key] ? 'border-rose-300' : 'border-slate-200 dark:border-slate-700'}`}
                        value={row[col.key] || ''}
                        onChange={e => updateSecretsRow(sectionId, idx, col.key, e.target.value)}
                        disabled={readOnly}
                      >
                        <option value="">-</option>
                        {col.options?.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    ) : col.type === 'checkbox' ? (
                      <div className="flex justify-center">
                        <input 
                          type="checkbox" 
                          className="w-5 h-5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
                          checked={!!row[col.key]}
                          onChange={e => updateSecretsRow(sectionId, idx, col.key, e.target.checked)}
                          disabled={readOnly}
                        />
                      </div>
                    ) : (
                      <input 
                        type={col.type || 'text'}
                        className={`w-full p-2 bg-white dark:bg-slate-800 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200 ${col.required && !row[col.key] ? 'border-rose-300' : 'border-slate-200 dark:border-slate-700'} ${col.readOnly ? 'bg-slate-100 dark:bg-slate-900 text-slate-500 dark:text-slate-400' : ''}`}
                        value={row[col.key] || ''}
                        onChange={e => !col.readOnly && updateSecretsRow(sectionId, idx, col.key, e.target.value)}
                        placeholder={col.label}
                        disabled={readOnly || col.readOnly}
                        maxLength={500}
                      />
                    )}
                  </td>
                ))}
                <td className="p-2 text-center">
                  {!readOnly && (
                  <button onClick={() => removeSecretsRow(sectionId, idx)} className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-colors">
                    <Trash2 className="w-4 h-4" />
                  </button>
                  )}
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr>
                <td colSpan={columns.length + 1} className="p-4 text-center text-slate-400 italic">Keine Einträge vorhanden.</td>
              </tr>
            )}
          </tbody>
        </table>
        {!readOnly && (
        <div className="flex gap-2 mt-3">
            <button onClick={() => addSecretsRow(sectionId)} className="flex items-center gap-2 text-sm font-bold text-indigo-600 hover:text-indigo-700 px-3 py-2 hover:bg-indigo-50 rounded-lg transition-colors">
            <Plus className="w-4 h-4" /> Zeile hinzufügen
            </button>
            {enableImport && (
                <>
                    <button onClick={() => handleDownloadTemplate(columns)} className="flex items-center gap-2 text-sm font-bold text-slate-600 hover:text-slate-700 px-3 py-2 hover:bg-slate-50 rounded-lg transition-colors" title="Excel-Vorlage herunterladen">
                        <FileDown className="w-4 h-4" /> Vorlage
                    </button>
                    <label className="flex items-center gap-2 text-sm font-bold text-emerald-600 hover:text-emerald-700 px-3 py-2 hover:bg-emerald-50 rounded-lg transition-colors cursor-pointer" title="Excel-Datei importieren">
                        <FileSpreadsheet className="w-4 h-4" /> Import Excel
                        <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={(e) => handleImport(e, sectionId, columns)} />
                    </label>
                </>
            )}
        </div>
        )}
      </div>
    );
  };

  const exportJSON = () => {
      let dataToExport;
      let prefix;
      if (activeTab === 'onboarding') {
          dataToExport = onboardingData;
          prefix = 'Onboarding';
      } else if (activeTab === 'technical') {
          dataToExport = technicalData;
          prefix = 'Technical_Structure';
      } else if (activeTab === 'secrets') {
          dataToExport = secretsData;
          prefix = 'Secrets_Management';
      } else {
          dataToExport = secretsOnboardingData;
          prefix = 'Secrets_Onboarding';
      }

      const blob = new Blob([JSON.stringify(dataToExport, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${prefix}_${governanceRow.ICTO}.json`;
      a.click();
  };

  // Derived options for Technical
  const serverOptions = (technicalData.servers || []).map(s => s.serverName).filter(Boolean);
  const accountOptions = (technicalData.sharedAccounts || []).map(a => a.techName).filter(Boolean);
  const safeOptions = (technicalData.safes || []).map(s => s.techSafeName).filter(Boolean);

  // Derived options for Secrets
  const secretInventoryOptions = (secretsData.inventory || []).map(s => s.name).filter(Boolean);
  const secretSafeOptions = (secretsData.safes || []).map(s => s.safeName).filter(Boolean);
  const secretTechSafeOptions = (secretsData.safes || []).map(s => s.techSafeName).filter(Boolean);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 print:p-0 animate-in zoom-in-95 duration-200 dark:text-slate-200">
      <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm print:hidden" onClick={onClose}></div>
      <div className="relative bg-white dark:bg-slate-900 w-full max-w-[90rem] h-[95vh] rounded-xl shadow-2xl flex flex-col overflow-hidden print:h-auto print:w-full print:max-w-none print:rounded-none print:shadow-none print:absolute print:inset-0 print:z-[100]">
        
        {/* Header */}
        <div className="px-8 py-6 border-b border-slate-100 dark:border-slate-800 flex justify-between items-center bg-white dark:bg-slate-900 sticky top-0 z-10 print:hidden">
            <div>
                <h2 className="text-2xl font-black text-slate-900">
                   {activeTab === 'onboarding' ? 'PAM Onboarding' : activeTab === 'technical' ? 'Technische Struktur Produktion' : activeTab === 'secrets' ? 'Secrets Management Inventar' : 'Onboarding Secrets Management'}
                </h2>
                <p className="text-slate-500 font-mono text-sm mt-1">{governanceRow.Name} ({governanceRow.ICTO})</p>
            </div>
            
            {/* Tab Switcher */}
            <div className="flex bg-slate-100 dark:bg-slate-800 p-1 rounded-lg mx-4">
                {mode === 'pam' && (
                  <>
                    <button 
                      onClick={() => setActiveTab('onboarding')}
                      className={`px-6 py-2.5 rounded-md text-sm font-bold transition-all flex items-center gap-2 ${activeTab === 'onboarding' ? 'bg-white dark:bg-slate-700 text-rose-600 dark:text-rose-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200'}`}
                    >
                      <Rocket className="w-4 h-4" /> Onboarding
                    </button>
                    <button 
                      onClick={() => setActiveTab('technical')}
                      className={`px-6 py-2.5 rounded-md text-sm font-bold transition-all flex items-center gap-2 ${activeTab === 'technical' ? 'bg-white dark:bg-slate-700 text-emerald-600 dark:text-emerald-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200'}`}
                    >
                      <Layers className="w-4 h-4" /> Technische Struktur
                    </button>
                  </>
                )}
                {mode === 'secrets' && (
                  <>
                    <button 
                      onClick={() => setActiveTab('secrets-onboarding')}
                      className={`px-6 py-2.5 rounded-md text-sm font-bold transition-all flex items-center gap-2 ${activeTab === 'secrets-onboarding' ? 'bg-white dark:bg-slate-700 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200'}`}
                    >
                      <ClipboardList className="w-4 h-4" /> Onboarding Secrets
                    </button>
                    <button 
                      onClick={() => setActiveTab('secrets')}
                      className={`px-6 py-2.5 rounded-md text-sm font-bold transition-all flex items-center gap-2 ${activeTab === 'secrets' ? 'bg-white dark:bg-slate-700 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200'}`}
                    >
                      <KeyRound className="w-4 h-4" /> Secrets Management Inventar
                    </button>
                  </>
                )}
            </div>

            <div className="flex items-center gap-3">
                <button 
                    onClick={() => {
                        if (activeTab === 'onboarding') {
                            if (onboardingViewMode === 'form') loadOnboardingHistory();
                            setOnboardingViewMode(v => v === 'form' ? 'history' : 'form');
                        } else if (activeTab === 'technical') {
                            if (technicalViewMode === 'form') loadTechnicalHistory();
                            setTechnicalViewMode(v => v === 'form' ? 'history' : 'form');
                        } else if (activeTab === 'secrets') {
                            if (secretsViewMode === 'form') loadSecretsHistory();
                            setSecretsViewMode(v => v === 'form' ? 'history' : 'form');
                        } else {
                            if (secretsOnboardingViewMode === 'form') loadSecretsOnboardingHistory();
                            setSecretsOnboardingViewMode(v => v === 'form' ? 'history' : 'form');
                        }
                    }}
                    className={`p-2 rounded-md transition-colors ${(activeTab === 'onboarding' ? onboardingViewMode : activeTab === 'technical' ? technicalViewMode : activeTab === 'secrets' ? secretsViewMode : secretsOnboardingViewMode) === 'history' ? 'bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400' : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'}`}
                    title="Änderungshistorie"
                >
                    <HistoryIcon className="w-5 h-5" />
                </button>
                <button onClick={exportJSON} className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-md text-slate-500 dark:text-slate-400" title="JSON Export"><Download className="w-5 h-5" /></button>
                <button onClick={onClose} className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-md text-slate-500 dark:text-slate-400"><X className="w-6 h-6" /></button>
            </div>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-8 bg-slate-50/50 dark:bg-slate-950/50 custom-scrollbar print:overflow-visible print:bg-white">
            {activeTab === 'onboarding' ? (
                onboardingLoading ? <div className="flex justify-center p-10"><RefreshCw className="animate-spin w-8 h-8 text-indigo-500" /></div> : (
                onboardingViewMode === 'history' ? (
                    <div className="max-w-4xl mx-auto">
                        <h3 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800 dark:text-slate-200">
                          <Clock className="w-6 h-6 text-indigo-500" /> Änderungsprotokoll
                        </h3>
                        <PaginatedHistoryList history={onboardingHistory} />
                    </div>
                ) : (
                <div className="space-y-4 max-w-4xl mx-auto">
                    
                    {/* Section 2: Architecture */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('architecture')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Database className="w-5 h-5 text-indigo-500" /> 1. Architektur</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['architecture'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['architecture'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                <div className="mb-4">
                                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1">Architektur des Zielsystems</label>
                                    <textarea className="w-full p-3 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md h-24 text-sm dark:text-slate-200" value={onboardingData.archDesc || ''} onChange={e => updateField('archDesc', e.target.value)} disabled={readOnly} maxLength={500}></textarea>
                                </div>
                                {renderInput("Betriebssysteme", "osList")}
                                {renderInput("Datenbanken", "dbList")}
                                {renderCheckbox("Ist geplant, die Server auszutauschen?", "serverReplace")}
                                {onboardingData.serverReplace && renderInput("Geplantes Austauschdatum", "serverReplaceDate", "date")}
                                {renderCheckbox("Load Balancer / Reverse Proxy vorhanden?", "hasLoadBalancer")}
                            </div>
                        )}
                    </div>

                    {/* Section 3: Technology */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('technology')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Settings className="w-5 h-5 text-slate-500" /> 2. Anwendungstechnologie</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['technology'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['technology'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 grid grid-cols-1 md:grid-cols-2 gap-6">
                                {renderInput("Technologie", "techType", "text", undefined, ["Fat Client", "Webbasiert", "Java", "Mischform", "Cloud-Service"])}
                                {renderCheckbox("Werden Zertifikate benötigt?", "needsCerts")}
                                {renderCheckbox("Aktives 4-Augen-Prinzip?", "fourEyes")}
                                {renderInput("Lizenzart", "licenseType", "text", undefined, ["Accountgebunden", "Identitätsgebunden", "Serverbasiert"])}
                                {renderInput("URL Produktion", "urlProd")}
                                {renderInput("URL Test", "urlTest")}
                            </div>
                        )}
                    </div>

                    {/* Section 4: Login */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('login')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><LogIn className="w-5 h-5 text-emerald-500" /> 3. Anmeldeverfahren</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['login'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['login'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderInput("Anmeldeverfahren", "loginMethod", "text", undefined, ["Single Sign-On", "Anmeldemaske", "Multifaktor-Authentifizierung", "anderes"])}
                                {renderInput("Account-Typ", "accountType", "text", "Wichtig für Passwort-Prozess", ["Lokale Accounts", "Domain Accounts"])}
                                {renderCheckbox("Kann User mehrere Sessions öffnen?", "multiSession")}
                                {renderInput("Username Nomenklatur", "usernameNaming")}
                            </div>
                        )}
                    </div>

                    {/* Section 5: Password (Conditional) */}
                    {onboardingData.accountType !== 'Domain Accounts' && (
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('password')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><ShieldCheck className="w-5 h-5 text-rose-500" /> 4. Passwortwechsel</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['password'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['password'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                <div className="mb-4">
                                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1">Wie werden Passwörter geändert?</label>
                                    <textarea className="w-full p-3 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md h-20 text-sm dark:text-slate-200" value={onboardingData.pwChangeDesc || ''} onChange={e => updateField('pwChangeDesc', e.target.value)} disabled={readOnly} maxLength={500}></textarea>
                                </div>
                                {renderCheckbox("Kann Rotation automatisiert werden?", "autoRotation")}
                                {renderInput("Wer ändert Passwörter?", "whoChangesPw")}
                                {renderInput("Wechselintervall", "pwInterval")}
                            </div>
                        )}
                    </div>
                    )}

                    {/* Section 6: Test */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('test')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><FileCheck className="w-5 h-5 text-amber-500" /> 5. Testuser & Umgebung</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['test'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['test'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderCheckbox("Testuser vorhanden?", "hasTestUsers")}
                                {onboardingData.hasTestUsers && renderInput("Testuser Accounts", "testUsersList")}
                                {renderCheckbox("Rechte über Omada?", "omadaRights")}
                            </div>
                        )}
                    </div>

                    {/* Section 7: Emergency */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('emergency')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><AlertTriangle className="w-5 h-5 text-rose-500" /> 6. Notfallprozess</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['emergency'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['emergency'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderCheckbox("Existieren Notfallaccounts?", "hasEmergencyAccounts")}
                                {onboardingData.hasEmergencyAccounts && renderInput("Notfall Accounts", "emergencyAccountsList")}
                            </div>
                        )}
                    </div>

                    {/* Section 8: Matrix */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('matrix')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Link2 className="w-5 h-5 text-indigo-600" /> 7. Vereinbarte Anbindungsvariante</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['matrix'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['matrix'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                <div className="mb-6 grid grid-cols-1 md:grid-cols-2 gap-4">
                                    {renderInput("Anmeldung", "matrixLogin", "text", undefined, ["Automatisch", "Manuell"])}
                                    {renderInput("Passwortwechsel", "matrixPwChange", "text", undefined, ["Automatisch", "Manuell"])}
                                </div>
                                {renderMatrix()}
                            </div>
                        )}
                    </div>

                    {/* Section 8: Bypass */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleOnboardingSection('bypass')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Shield className="w-5 h-5 text-rose-500" /> 8. PAM-Bypass Regeln</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${onboardingOpenSections['bypass'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(onboardingOpenSections['bypass'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderInput("S-Base Pro Auftrag", "sBaseProOrder")}
                            </div>
                        )}
                    </div>

                </div>
                )
            )) : activeTab === 'technical' ? (
                technicalLoading ? <div className="flex justify-center p-10"><RefreshCw className="animate-spin w-8 h-8 text-emerald-500" /></div> : (
                technicalViewMode === 'history' ? (
                    <div className="max-w-[85rem] mx-auto">
                        <h3 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800 dark:text-slate-200">
                          <Clock className="w-6 h-6 text-indigo-500" /> Änderungsprotokoll
                        </h3>
                        <PaginatedHistoryList history={technicalHistory} />
                    </div>
                ) : (
                <div className="space-y-6 max-w-[85rem] mx-auto">
                    
                    {/* 1. Servers */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('servers')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Server className="w-5 h-5 text-blue-500" /> 1. Server / Betriebssysteme</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['servers'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['servers'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('servers', [
                                  { key: 'serverName', label: 'Servername', required: true },
                                  { key: 'ip', label: 'IP-Adresse' },
                                  { key: 'fqdn', label: 'Adresse / FQDN' },
                                  { key: 'stage', label: 'Stage', type: 'select', options: ['Prod', 'Test', 'Dev'], required: true },
                                  { key: 'dmz', label: 'DMZ', type: 'checkbox', width: '60px' },
                                  { key: 'desc', label: 'Beschreibung / Nutzung' },
                                  { key: 'os', label: 'Betriebssystem' },
                                  { key: 'port', label: 'Zugriff Port/Protokoll' },
                                  { key: 'expiry', label: 'Ablaufdatum', type: 'date' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 2. Databases */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('databases')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Database className="w-5 h-5 text-indigo-500" /> 2. Datenbanken / Server</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['databases'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['databases'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('databases', [
                                  { key: 'serverName', label: 'Servername', required: true },
                                  { key: 'ip', label: 'IP-Adresse' },
                                  { key: 'fqdn', label: 'Adresse / FQDN' },
                                  { key: 'stage', label: 'Stage', type: 'select', options: ['Prod', 'Test', 'Dev'], required: true },
                                  { key: 'dbType', label: 'Datenbanktyp' },
                                  { key: 'instance', label: 'Instanz / DB-Name' },
                                  { key: 'product', label: 'Produktname & Version' },
                                  { key: 'port', label: 'Zugriff Port/Protokoll' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 3. Ports */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('ports')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Network className="w-5 h-5 text-emerald-500" /> 3. Portfreischaltungen</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['ports'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['ports'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('ports', [
                                  { key: 'fromServer', label: 'Von Server', type: 'multiselect', options: ['psm1', 'psm2', 'psm3', 'psm4', 'psm5', 'psm6', 'psmp'], width: '300px' },
                                  { key: 'fromStage', label: 'Von Stage', type: 'select', options: ['Prod', 'Test', 'Dev'] },
                                  { key: 'toServer', label: 'Nach Server', type: 'select', options: serverOptions },
                                  { key: 'toIp', label: 'Nach IP' },
                                  { key: 'toStage', label: 'Nach Stage', type: 'select', options: ['Prod', 'Test', 'Dev'] },
                                  { key: 'port', label: 'Port/Protokoll', required: true },
                                  { key: 'provider', label: 'Dienstleister' },
                                  { key: 'interfaceId', label: 'Schnittstellen-ID' },
                                  { key: 'comment', label: 'Kommentar' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 4. Safes */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('safes')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Lock className="w-5 h-5 text-amber-500" /> 4. Safe-Struktur (CyberArk)</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['safes'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['safes'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('safes', [
                                  { key: 'userGroup', label: 'Nutzergruppe' },
                                  { key: 'safeName', label: 'Safe Name (fachlich)', required: true },
                                  { key: 'safeDesc', label: 'Safe Beschreibung' },
                                  { key: 'techSafeName', label: 'Technischer Safe Name', required: true },
                                  { key: 'adGroup', label: 'AD-Gruppe', required: true },
                                  { key: 'adGroupDesc', label: 'Beschreibung AD-Gruppe' },
                                  { key: 'approver', label: 'Zweitgenehmiger (Omada)' },
                                  { key: 'sod', label: 'SoD-Hinweis' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 5. Safe Members */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('safeMembers')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Users className="w-5 h-5 text-purple-500" /> 5. Mitglieder der Safes</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['safeMembers'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['safeMembers'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                <div className="mb-4 p-4 bg-blue-50 text-blue-800 rounded-lg text-sm flex items-start gap-3">
                                  <Info className="w-5 h-5 shrink-0 mt-0.5" />
                                  <div>
                                    <strong>Hinweis:</strong> Änderungen an der Mitgliederstruktur erfolgen regelmäßig. Die führende Quelle für aktuelle Safe-Mitglieder ist Omada. Diese Tabelle dient der Dokumentation des Soll-Zustands.
                                  </div>
                                </div>
                                {renderTable('safeMembers', [
                                  { key: 'safeName', label: 'Safe Name', type: 'select', options: safeOptions },
                                  { key: 'adGroup', label: 'AD-Gruppe des Safes' },
                                  { key: 'memberName', label: 'Name Mitglied' },
                                  { key: 'identity', label: 'Primäre Identität' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 6. Shared Accounts */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('sharedAccounts')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Key className="w-5 h-5 text-rose-500" /> 6. Shared Accounts</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['sharedAccounts'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['sharedAccounts'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('sharedAccounts', [
                                  { key: 'bizName', label: 'Fachlicher Name' },
                                  { key: 'techName', label: 'Technischer Name', required: true },
                                  { key: 'sam', label: 'sAMAccountName' },
                                  { key: 'login', label: 'Anmeldename (Ziel)' },
                                  { key: 'desc', label: 'Beschreibung' },
                                  { key: 'isAd', label: 'AD Account', type: 'checkbox', width: '80px' },
                                  { key: 'owner', label: 'Accountbesitzer' },
                                  { key: 'ownerId', label: 'Identität Besitzer' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 7. Permissions */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('permissions')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Shield className="w-5 h-5 text-cyan-500" /> 7. Berechtigungszuordnungen</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['permissions'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['permissions'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('permissions', [
                                  { key: 'bizName', label: 'Fachlicher Account', type: 'select', options: accountOptions },
                                  { key: 'techName', label: 'Technischer Account', type: 'select', options: accountOptions },
                                  { key: 'roleId', label: 'Role ID', required: true },
                                  { key: 'roleName', label: 'Role Displayname' },
                                  { key: 'roleDesc', label: 'Role Description' },
                                  { key: 'bizSystem', label: 'Fachliches System' },
                                  { key: 'bizSystemId', label: 'Fachliche System-ID' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 8. Mapping */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleTechnicalSection('mapping')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Link className="w-5 h-5 text-indigo-600" /> 8. Shared Accounts zu Safe</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${technicalOpenSections['mapping'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(technicalOpenSections['mapping'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderTable('mapping', [
                                  { key: 'techAccount', label: 'Technischer Account', type: 'select', options: accountOptions, required: true },
                                  { key: 'bizAccount', label: 'Fachlicher Account', readOnly: true },
                                  { key: 'techSafe', label: 'Technischer Safe', type: 'select', options: safeOptions, required: true },
                                  { key: 'safeName', label: 'Fachlicher Safe', readOnly: true }
                                ])}
                            </div>
                        )}
                    </div>

                </div>
                )
            )) : activeTab === 'secrets' ? (
                secretsLoading ? <div className="flex justify-center p-10"><RefreshCw className="animate-spin w-8 h-8 text-indigo-500" /></div> : (
                secretsViewMode === 'history' ? (
                    <div className="max-w-[85rem] mx-auto">
                        <h3 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800">
                          <Clock className="w-6 h-6 text-indigo-500" /> Änderungsprotokoll
                        </h3>
                        <PaginatedHistoryList history={secretsHistory} />
                    </div>
                ) : (
                <div className="space-y-6 max-w-[85rem] mx-auto">
                    
                    {/* 1. Secret Inventory */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsSection('inventory')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><KeyRound className="w-5 h-5 text-indigo-500" /> 1. Secret Inventory</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOpenSections['inventory'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOpenSections['inventory'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderSecretsTable('inventory', [
                                  { key: 'category', label: 'Kategorie', type: 'select', options: ['Passwort', 'SSH-Key', 'API-Key', 'Zertifikat'], required: true, width: '200px' },
                                  { key: 'name', label: 'Name / ID', required: true, width: '300px' },
                                  { key: 'owner', label: 'Secret Owner', width: '250px' },
                                  { key: 'holder', label: 'Secret Holder', width: '250px' },
                                  { key: 'layer', label: 'Layer', type: 'select', options: ['Anwendung', 'Betriebssystem', 'Datenbank'], width: '200px' },
                                  { key: 'localOrAd', label: 'Lokal oder AD', type: 'select', options: ['Lokal', 'AD'], width: '150px' },
                                  { key: 'stage', label: 'Stage', type: 'select', options: ['Prod', 'Test', 'Dev', 'Int'], required: true, width: '150px' },
                                  { key: 'complexity', label: 'Komplexitätsregeln', width: '300px' },
                                  { key: 'autoRotation', label: 'Auto Rotation', type: 'select', options: ['Ja', 'Nein'], width: '150px' },
                                  { key: 'rotationMech', label: 'Rotationsmechanismus', width: '250px' },
                                  { key: 'frequency', label: 'Häufigkeit', type: 'select', options: ['Täglich', 'Wöchentlich', 'Monatlich', 'Jährlich', 'bei Bedarf'], width: '200px' },
                                  { key: 'timeWindow', label: 'Zeitfenster', width: '250px' }
                                ], true)}
                            </div>
                        )}
                    </div>

                    {/* 2. Safe Structure */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsSection('safes')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Lock className="w-5 h-5 text-amber-500" /> 2. Safe- / Pfad-Struktur</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOpenSections['safes'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOpenSections['safes'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                <div className="mb-4 p-4 bg-blue-50 text-blue-800 rounded-lg text-sm flex items-start gap-3">
                                  <Info className="w-5 h-5 shrink-0 mt-0.5" />
                                  <div>
                                    <strong>Info:</strong> Die Safe- und Pfad-Struktur stellt eine logische Bündelung von Secrets und Shared Accounts dar. Safes werden in CyberArk verwendet, um Berechtigungen strukturiert zu vergeben.
                                  </div>
                                </div>
                                {renderSecretsTable('safes', [
                                  { key: 'userGroup', label: 'Nutzergruppe' },
                                  { key: 'safeName', label: 'Safe Name / Pfad (fachlich)', required: true },
                                  { key: 'safeDesc', label: 'Safe / Pfad Beschreibung' },
                                  { key: 'techSafeName', label: 'Technischer Safe Name / Pfad-Prefix', required: true },
                                  { key: 'adGroup', label: 'AD-Gruppe', required: true },
                                  { key: 'adGroupDesc', label: 'Fachliche Beschreibung AD-Gruppe' },
                                  { key: 'approver', label: 'Zweitgenehmiger in Omada' },
                                  { key: 'sod', label: 'SoD-Hinweis' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 3. Members */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsSection('members')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Users className="w-5 h-5 text-purple-500" /> 3. Mitglieder der Safes</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOpenSections['members'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOpenSections['members'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                <div className="mb-4 p-4 bg-blue-50 text-blue-800 rounded-lg text-sm flex items-start gap-3">
                                  <Info className="w-5 h-5 shrink-0 mt-0.5" />
                                  <div>
                                    <strong>Hinweis:</strong> Änderungen in der Mitgliederstruktur der Safes erfolgen aufgrund laufender Pflege und Aktualisierung nicht im Rahmen dieses Dokuments. Die führende Quelle für aktuelle Safe-Mitglieder ist Omada.
                                  </div>
                                </div>
                                {renderSecretsTable('members', [
                                  { key: 'safeName', label: 'Safe Name', type: 'select', options: secretSafeOptions },
                                  { key: 'adGroup', label: 'AD-Gruppe des Safes' },
                                  { key: 'memberName', label: 'Name Mitglied' },
                                  { key: 'identity', label: 'Primäre Identität' }
                                ])}
                            </div>
                        )}
                    </div>

                    {/* 4. Mapping */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsSection('mapping')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Link className="w-5 h-5 text-emerald-500" /> 4. Secrets zu Safe</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOpenSections['mapping'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOpenSections['mapping'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700">
                                {renderSecretsTable('mapping', [
                                  { key: 'bizSecret', label: 'Fachlicher Secret Name', type: 'select', options: secretInventoryOptions, required: true },
                                  { key: 'safeName', label: 'Safe Name', type: 'select', options: secretSafeOptions, required: true },
                                  { key: 'techSafe', label: 'Technischer Safe', readOnly: true }
                                ])}
                            </div>
                        )}
                    </div>

                </div>
                )
            )) : (
                secretsOnboardingLoading ? <div className="flex justify-center p-10"><RefreshCw className="animate-spin w-8 h-8 text-indigo-500" /></div> : (
                secretsOnboardingViewMode === 'history' ? (
                    <div className="max-w-[85rem] mx-auto">
                        <h3 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800">
                          <Clock className="w-6 h-6 text-indigo-500" /> Änderungsprotokoll
                        </h3>
                        <PaginatedHistoryList history={secretsOnboardingHistory} />
                    </div>
                ) : (
                <div className="space-y-6 max-w-4xl mx-auto">
                    {/* 1. Infrastruktur */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('infrastructure')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Server className="w-5 h-5 text-indigo-500" /> 1. Infrastruktur</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['infrastructure'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['infrastructure'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Von wo aus wird auf die Secrets zugegriffen?", "accessSource")}
                                {renderSecretsOnboardingInput("Betriebsmodell", "operatingModel", "text", undefined, ["On-Prem", "Cloud", "Hybrid"])}
                                
                                <div className="mt-6 border-t border-slate-100 dark:border-slate-700 pt-4">
                                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">Serverregister</label>
                                    <div className="overflow-x-auto">
                                        <table className="w-full text-sm border-collapse">
                                            <thead>
                                                <tr className="bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700">
                                                    <th className="p-3 text-left font-bold text-slate-600 dark:text-slate-300 w-1/3">Hostname</th>
                                                    <th className="p-3 text-left font-bold text-slate-600 dark:text-slate-300 w-1/3">IP-Adresse</th>
                                                    <th className="p-3 text-left font-bold text-slate-600 dark:text-slate-300 w-1/3">Segment</th>
                                                    <th className="p-3 w-10"></th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {(secretsOnboardingData.serverRegistry || []).map((row: any, idx: number) => (
                                                    <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50/50">
                                                        <td className="p-2">
                                                            <input 
                                                                type="text" 
                                                                className="w-full p-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200"
                                                                value={row.hostname || ''}
                                                                onChange={e => updateServerRegistry(idx, 'hostname', e.target.value)}
                                                                onBlur={e => resolveHostname(idx, e.target.value)}
                                                                placeholder="Hostname"
                                                                disabled={readOnly}
                                                            />
                                                        </td>
                                                        <td className="p-2">
                                                            <input 
                                                                type="text" 
                                                                className="w-full p-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200"
                                                                value={row.ip || ''}
                                                                onChange={e => updateServerRegistry(idx, 'ip', e.target.value)}
                                                                placeholder="IP-Adresse"
                                                                disabled={readOnly}
                                                            />
                                                        </td>
                                                        <td className="p-2">
                                                            <select 
                                                                className="w-full p-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 dark:text-slate-200"
                                                                value={row.segment || 'Produktion'}
                                                                onChange={e => updateServerRegistry(idx, 'segment', e.target.value)}
                                                                disabled={readOnly}
                                                            >
                                                                <option value="Produktion">Produktion</option>
                                                                <option value="Test">Test</option>
                                                                <option value="Entwicklung">Entwicklung</option>
                                                            </select>
                                                        </td>
                                                        <td className="p-2 text-center">
                                                            {!readOnly && (
                                                                <button onClick={() => removeServerRegistryRow(idx)} className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-colors">
                                                                    <Trash2 className="w-4 h-4" />
                                                                </button>
                                                            )}
                                                        </td>
                                                    </tr>
                                                ))}
                                                {(secretsOnboardingData.serverRegistry || []).length === 0 && (
                                                    <tr>
                                                        <td colSpan={4} className="p-4 text-center text-slate-400 italic">Keine Server eingetragen.</td>
                                                    </tr>
                                                )}
                                            </tbody>
                                        </table>
                                        {!readOnly && (
                                            <button onClick={addServerRegistryRow} className="mt-3 flex items-center gap-2 text-sm font-bold text-indigo-600 hover:text-indigo-700 px-3 py-2 hover:bg-indigo-50 rounded-lg transition-colors">
                                                <Plus className="w-4 h-4" /> Server hinzufügen
                                            </button>
                                        )}
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>

                    {/* 2. Tool-Nutzung */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('tools')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Settings className="w-5 h-5 text-indigo-500" /> 2. Tool-Nutzung</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['tools'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['tools'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Läuft die Anwendung auf:", "appType", "text", undefined, ["Statischer Server", "Container"])}
                                {renderSecretsOnboardingInput("Betriebssystem", "os")}
                                {renderSecretsOnboardingInput("Wie wird die Anwendung bereitgestellt?", "deployment")}
                                {renderSecretsOnboardingInput("Wie werden Anmeldeinformationen initial bereitgestellt?", "initialAuth")}
                            </div>
                        )}
                    </div>

                    {/* 3. Container - Standalone */}
                    {secretsOnboardingData.appType === 'Container' && (
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('container_standalone')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Database className="w-5 h-5 text-indigo-500" /> 3. Container – Standalone</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['container_standalone'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['container_standalone'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Wurde der Container / die Anwendung selbst erstellt?", "containerSelfBuilt", "text", undefined, ["Ja", "Nein", "N/A"])}
                            </div>
                        )}
                    </div>
                    )}

                    {/* 4. Container - K8s */}
                    {secretsOnboardingData.appType === 'Container' && (
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('container_k8s')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Cloud className="w-5 h-5 text-indigo-500" /> 4. Container – Kubernetes / OpenShift</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['container_k8s'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['container_k8s'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Verwendet die Anwendung Kubernetes Secrets?", "k8sSecrets", "text", undefined, ["Ja", "Nein", "N/A"])}
                                {renderSecretsOnboardingInput("Besteht Kontrolle über den Anwendungscode?", "codeControl", "text", undefined, ["Ja", "Nein", "N/A"])}
                                <div className="p-3 bg-amber-50 text-amber-800 text-sm rounded-lg border border-amber-100 flex gap-2">
                                    <AlertTriangle className="w-5 h-5 shrink-0" />
                                    <span><strong>Hinweis:</strong> Kubernetes Secrets sind standardmäßig nur base64-kodiert und nicht verschlüsselt. Dies stellt ein Sicherheitsrisiko dar.</span>
                                </div>
                            </div>
                        )}
                    </div>
                    )}

                    {/* 5. Cloud Computing */}
                    {(secretsOnboardingData.operatingModel === 'Cloud' || secretsOnboardingData.operatingModel === 'Hybrid') && (
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('cloud')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Cloud className="w-5 h-5 text-indigo-500" /> 5. Cloud Computing</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['cloud'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['cloud'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Cloud in", "cloudIn", "text", undefined, ["FCN", "FCPI", "SAAS"])}
                                {renderSecretsOnboardingInput("Nutzung von Serverless Functions?", "serverless", "text", undefined, ["Ja", "Nein", "N/A"])}
                                {renderSecretsOnboardingInput("Nutzung von VMs in der Cloud (Hyperscaler)?", "cloudVMs", "text", undefined, ["Ja", "Nein", "N/A"])}
                            </div>
                        )}
                    </div>
                    )}

                    {/* 6. Server */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('server')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Server className="w-5 h-5 text-indigo-500" /> 6. Server</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['server'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['server'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {secretsOnboardingData.operatingModel === 'On-Prem' && (
                                    <div className="p-3 bg-blue-50 text-blue-800 text-sm rounded-lg border border-blue-100 mb-4">
                                        <strong>Hinweis:</strong> Bitte beantworten Sie die folgenden Fragen spezifisch für Ihre On-Prem Server-Umgebung.
                                    </div>
                                )}
                                {renderSecretsOnboardingInput("Besteht Kontrolle über Code?", "codeControlServer", "text", undefined, ["Ja", "Nein"])}
                                {renderSecretsOnboardingInput("Besteht Kontrolle über Konfigurationsdateien?", "configControlServer", "text", undefined, ["Ja", "Nein"])}
                                {renderSecretsOnboardingInput("Sind Konfigurationsdateien zentral zugänglich?", "centralConfig", "text", undefined, ["Ja", "Nein"])}
                                {renderSecretsOnboardingInput("Läuft die Anwendung auf einem Java-Webserver?", "javaWebserver", "text", undefined, ["Ja", "Nein", "Teilweise"])}
                                {(secretsOnboardingData.javaWebserver === 'Ja' || secretsOnboardingData.javaWebserver === 'Teilweise') && (
                                    <div className="pl-4 border-l-2 border-indigo-100 mt-2 space-y-4">
                                        {renderSecretsOnboardingInput("Nutzt die Anwendung DataSources?", "dataSources", "text", undefined, ["Ja", "Nein"])}
                                        {renderSecretsOnboardingInput("Zweck der DataSources", "dataSourcesPurpose")}
                                        <div className="text-xs text-slate-500 italic">Relevant für Secrets-Injection in Java-Umgebungen.</div>
                                    </div>
                                )}
                            </div>
                        )}
                    </div>

                    {/* 7. Eigenschaften der Secrets */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('properties')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Key className="w-5 h-5 text-indigo-500" /> 7. Eigenschaften der Secrets</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['properties'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['properties'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingMultiSelect("Welche Arten von Secrets werden genutzt?", "secretTypes", ["Passwörter", "API-Keys", "Zertifikate", "SSH-Keys"])}
                                {renderSecretsOnboardingMultiSelect("Wie werden Secrets aktuell gespeichert?", "secretStorage", ["Config-Files", "ENV-Variablen", "Kubernetes Secrets", "Hard-coded", "Hardware-HSM"])}
                                {((secretsOnboardingData.secretStorage as string[]) || []).includes('Hard-coded') && (
                                    <div className="p-3 bg-rose-50 text-rose-800 text-sm rounded-lg border border-rose-100 flex gap-2">
                                        <AlertTriangle className="w-5 h-5 shrink-0" />
                                        <span><strong>Risiko:</strong> Hard-coded Secrets stellen ein hohes Sicherheitsrisiko dar und müssen priorisiert migriert werden.</span>
                                    </div>
                                )}
                                {renderSecretsOnboardingInput("Werden Secrets auch für manuelle Prozesse genutzt?", "manualProcesses", "text", undefined, ["Ja", "Nein"])}
                            </div>
                        )}
                    </div>

                    {/* 8. Rotation - Ist-Status */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('rotation_status')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><RefreshCw className="w-5 h-5 text-indigo-500" /> 8. Rotation – Ist-Status</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['rotation_status'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['rotation_status'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Gibt es Secrets, deren Rotation in der Verantwortung der Deka liegt?", "rotationResponsibility", "text", undefined, ["Ja", "Nein"])}
                                {renderSecretsOnboardingMultiSelect("Gilt dies für:", "rotationLevel", ["Datenbankebene", "Anwendungsebene"])}
                                {renderSecretsOnboardingInput("Wie erfolgt die bisherige Rotation?", "currentRotation", "text", undefined, ["Manuell", "Teilautomatisiert", "Vollautomatisiert", "Bereits über CyberArk"])}
                            </div>
                        )}
                    </div>

                    {/* 9. Anbindungsvariante (Zielbild) */}
                    <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm print:border-0 print:shadow-none">
                        <button onClick={() => toggleSecretsOnboardingSection('target_image')} className="w-full px-6 py-4 flex items-center justify-between bg-white dark:bg-slate-900 hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors print:hidden">
                            <div className="flex items-center gap-3 font-bold text-slate-700 dark:text-slate-200"><Link2 className="w-5 h-5 text-indigo-500" /> 9. Anbindungsvariante (Zielbild)</div>
                            <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform ${secretsOnboardingOpenSections['target_image'] ? 'rotate-90' : ''}`} />
                        </button>
                        {(secretsOnboardingOpenSections['target_image'] || typeof window !== 'undefined' && window.matchMedia('print').matches) && (
                            <div className="p-6 border-t border-slate-100 dark:border-slate-700 space-y-4">
                                {renderSecretsOnboardingInput("Eingesetztes Secrets-Management-Tool", "targetTool", "text", undefined, ["CyberArk CP", "CyberArk CCP", "HashiCorp Vault"])}
                                {renderSecretsOnboardingInput("Anbindungsvariante", "targetVariant", "text", undefined, ["Vollautomatisiert zur Laufzeit", "Hybrid", "Manuelle Inventarisierung"])}
                                {renderSecretsOnboardingInput("Rotationsmechanismus", "targetRotationMech", "text", undefined, ["Automatisch", "Teilautomatisiert", "Manuell"])}
                                {renderSecretsOnboardingInput("Zeitfenster der Rotation", "targetTimeWindow", "text", undefined, ["Untertägig", "Definiertes Zeitfenster", "Manuell"])}
                                {renderSecretsOnboardingInput("Häufigkeit der Rotation", "targetFrequency", "text", undefined, ["Täglich", "Wöchentlich", "Individueller Zyklus", "1x pro Jahr"])}
                            </div>
                        )}
                    </div>
                </div>
                )
            ))}
        </div>

        {/* Footer */}
        <div className="p-6 border-t border-slate-100 dark:border-slate-800 bg-white dark:bg-slate-900 flex justify-end gap-4 print:hidden">
            <button onClick={onClose} className="px-6 py-3 font-bold text-slate-500 hover:bg-slate-50 rounded-lg transition-colors">Schließen</button>
            {!readOnly && (
            <button 
                onClick={activeTab === 'onboarding' ? handleOnboardingSave : activeTab === 'technical' ? handleTechnicalSave : activeTab === 'secrets' ? handleSecretsSave : handleSecretsOnboardingSave} 
                disabled={(activeTab === 'onboarding' ? onboardingSaveStatus : activeTab === 'technical' ? technicalSaveStatus : activeTab === 'secrets' ? secretsSaveStatus : secretsOnboardingSaveStatus) === 'saving'} 
                className={`px-8 py-3 font-bold rounded-lg shadow-lg transition-all active:scale-95 flex items-center gap-2 ${
                    (activeTab === 'onboarding' ? onboardingSaveStatus : activeTab === 'technical' ? technicalSaveStatus : activeTab === 'secrets' ? secretsSaveStatus : secretsOnboardingSaveStatus) === 'success' 
                    ? 'bg-emerald-500 hover:bg-emerald-600 text-white shadow-emerald-200' 
                    : 'bg-indigo-600 hover:bg-indigo-700 text-white shadow-indigo-200'
                }`}
            >
                {(activeTab === 'onboarding' ? onboardingSaveStatus : activeTab === 'technical' ? technicalSaveStatus : activeTab === 'secrets' ? secretsSaveStatus : secretsOnboardingSaveStatus) === 'saving' ? <RefreshCw className="animate-spin w-5 h-5" /> : 
                 (activeTab === 'onboarding' ? onboardingSaveStatus : activeTab === 'technical' ? technicalSaveStatus : activeTab === 'secrets' ? secretsSaveStatus : secretsOnboardingSaveStatus) === 'success' ? <CheckCircle className="w-5 h-5" /> : 
                 <Save className="w-5 h-5" />} 
                {(activeTab === 'onboarding' ? onboardingSaveStatus : activeTab === 'technical' ? technicalSaveStatus : activeTab === 'secrets' ? secretsSaveStatus : secretsOnboardingSaveStatus) === 'success' ? 'Gespeichert' : 'Speichern'}
            </button>
            )}
        </div>

      </div>
    </div>
  );
};

const App = () => {
  const [data, setData] = useState<DataRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [connectionState, setConnectionState] = useState<'online' | 'api_only' | 'offline'>('offline');
  const [lastError, setLastError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [editingRow, setEditingRow] = useState<DataRow | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'success' | 'error'>('idle');
  const [user, setUser] = useState<string | null>(getUser());
  const [role, setRole] = useState<string | null>(getRole());
  const [loginCreds, setLoginCreds] = useState({ username: '', password: '' });
  const [history, setHistory] = useState<any[]>([]);
  const [viewMode, setViewMode] = useState<'form' | 'history'>('form');
  const [unifiedRow, setUnifiedRow] = useState<DataRow | null>(null);
  const [initialTab, setInitialTab] = useState<'onboarding' | 'technical' | 'secrets' | 'secrets-onboarding'>('onboarding');
  const [modalMode, setModalMode] = useState<'pam' | 'secrets'>('pam');
  const [onboardingVariant, setOnboardingVariant] = useState<string>('');
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [isAcceptanceConfirmOpen, setIsAcceptanceConfirmOpen] = useState(false);
  const [isTableExpanded, setIsTableExpanded] = useState(false);
  const saveLock = React.useRef(false);
  // Initialisierung: Standard ist false (Light Mode), es sei denn 'true' steht im Storage
  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== 'undefined') {
        return localStorage.getItem('pxm_dark_mode') === 'true';
    }
    return false;
  });

  const isReadOnly = role === 'readonly';

  useEffect(() => { 
    checkConnectivity(); 
  }, []);

  useEffect(() => {
    if (isModalOpen && editingRow?.id) {
        const fetchVariant = async () => {
            try {
                const token = getAccessToken();
                const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
                const res = await fetch(`${API_BASE_URL}/onboarding/${editingRow.id}`, { headers });
                if (res.ok) {
                    const json = await res.json();
                    if (json.data) {
                        const d = JSON.parse(json.data);
                        setOnboardingVariant(d.selectedVariant || '');
                    } else {
                        setOnboardingVariant('');
                    }
                }
            } catch (e) {
                console.error(e);
                setOnboardingVariant('');
            }
        };
        fetchVariant();
    } else {
        setOnboardingVariant('');
    }
  }, [isModalOpen, editingRow?.id]);

  useEffect(() => {
    const root = window.document.documentElement;
    if (darkMode) {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
    localStorage.setItem('pxm_dark_mode', String(darkMode));
  }, [darkMode]);

  const checkConnectivity = async () => {
    setLoading(true);
    setLastError(null);
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};

      const healthResp = await fetch(`${API_BASE_URL}/health`);
      if (!healthResp.ok) throw new Error("API unreachable");
      const health = await healthResp.json();
      
      if (health.database === 'connected') {
        setConnectionState('online');
        const dataResp = await fetch(`${API_BASE_URL}/data`, { headers });
        
        if (dataResp.status === 401 || dataResp.status === 403) {
           if (user) handleLogout(); // Token expired
           return;
        }

        const result = await dataResp.json();
        if (Array.isArray(result)) {
          setData(result);
        } else {
          setData([]);
        }
      } else {
        setConnectionState('api_only');
        setLastError(health.message || 'Datenbank-Verbindung fehlgeschlagen.');
      }
    } catch (err: any) {
      setConnectionState('offline');
      setLastError(err.message || 'Server.js nicht erreichbar.');
    } finally {
      setLoading(false);
    }
  };

  const visibleHeaders = useMemo(() => {
    if (isTableExpanded) return DEFAULT_HEADERS;
    const cutoffIndex = DEFAULT_HEADERS.indexOf("Kritikalität");
    return DEFAULT_HEADERS.slice(0, cutoffIndex + 1);
  }, [isTableExpanded]);

  const filteredData = useMemo(() => {
    return data.filter(row => {
      // Global search
      if (search && !JSON.stringify(row).toLowerCase().includes(search.toLowerCase())) {
        return false;
      }
      // Column filters
      for (const h of DEFAULT_HEADERS) {
        const filterVal = filters[h]?.toLowerCase();
        if (filterVal) {
          const cellVal = String(row[h] || '').toLowerCase();
          if (!cellVal.includes(filterVal)) return false;
        }
      }
      return true;
    });
  }, [data, search, filters]);

  const handleLoginSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      const { user, role } = await login(loginCreds.username, loginCreds.password);
      setUser(user);
      setRole(role);
      checkConnectivity();
    } catch (err: any) {
      alert(err.message);
    }
  };

  const handleLogout = () => {
    logout();
    setUser(null);
    setRole(null);
    setData([]);
    setConnectionState('api_only'); // Reset state visually
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (saveLock.current) return;
    if (!editingRow) return;
    saveLock.current = true;
    setSaveStatus('saving');
    try {
      const isUpdate = editingRow.id && !String(editingRow.id).startsWith('temp_');
      const method = isUpdate ? 'PUT' : 'POST';
      const url = isUpdate ? `${API_BASE_URL}/data/${editingRow.id}` : `${API_BASE_URL}/data`;
      
      const payload = { ...editingRow };
      if (!isUpdate) delete payload.id;

      const token = getAccessToken();
      const headers: HeadersInit = { 'Content-Type': 'application/json' };
      if (token) headers['Authorization'] = `Bearer ${token}`;

      const response = await fetch(url, {
        method,
        headers,
        body: JSON.stringify(payload),
      });
      
      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || 'Speichern fehlgeschlagen');
      }
      
      setSaveStatus('success');
      checkConnectivity();
      setTimeout(() => { 
        if (!isUpdate) {
            setIsModalOpen(false);
        }
        setSaveStatus('idle'); 
        saveLock.current = false;
      }, 2000);
    } catch (err: any) {
      setSaveStatus('error');
      setLastError(err.message);
      alert("Fehler beim Speichern: " + err.message);
      saveLock.current = false;
    }
  };

  const handleAcceptance = () => {
    if (!editingRow?.id) return;
    setIsAcceptanceConfirmOpen(true);
  };

  const confirmAcceptance = async () => {
    setIsAcceptanceConfirmOpen(false);
    const timestamp = new Date().toLocaleString('de-DE');
    const acceptanceString = `${user} am ${timestamp}`;
    const updatedRow = { ...editingRow, AbnahmePAMOnboarding: acceptanceString };
    
    setEditingRow(updatedRow);

    try {
        const token = getAccessToken();
        const headers: HeadersInit = { 'Content-Type': 'application/json' };
        if (token) headers['Authorization'] = `Bearer ${token}`;
        
        const url = `${API_BASE_URL}/data/${updatedRow.id}`;
        const response = await fetch(url, {
            method: 'PUT',
            headers,
            body: JSON.stringify(updatedRow),
        });

        if (!response.ok) throw new Error("Speichern fehlgeschlagen");
        checkConnectivity();
    } catch (e: any) {
        alert("Fehler: " + e.message);
    }
  };

  const handleRevokeAcceptance = async () => {
    if (!editingRow?.id) return;
    if (!window.confirm("Soll die Abnahme wirklich zurückgezogen werden?")) return;

    const updatedRow = { ...editingRow, AbnahmePAMOnboarding: '' };
    
    setEditingRow(updatedRow);

    try {
        const token = getAccessToken();
        const headers: HeadersInit = { 'Content-Type': 'application/json' };
        if (token) headers['Authorization'] = `Bearer ${token}`;
        
        const url = `${API_BASE_URL}/data/${updatedRow.id}`;
        const response = await fetch(url, {
            method: 'PUT',
            headers,
            body: JSON.stringify(updatedRow),
        });

        if (!response.ok) throw new Error("Speichern fehlgeschlagen");
        checkConnectivity();
    } catch (e: any) {
        alert("Fehler: " + e.message);
    }
  };

  const handleFetchHistory = async (id: any) => {
    if (!id) return;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const response = await fetch(`${API_BASE_URL}/history/${id}`, { headers });
      if (response.ok) {
        const histData = await response.json();
        setHistory(histData);
        setViewMode('history');
      }
    } catch (err) {
      alert("Fehler beim Laden der Historie");
    }
  };

  const handleFullExcelExport = async () => {
    if (!editingRow?.id) return;
    try {
      const token = getAccessToken();
      const headers: HeadersInit = token ? { 'Authorization': `Bearer ${token}` } : {};
      const id = editingRow.id;

      // Lade alle Daten parallel
      const [onbRes, techRes, secRes, secOnbRes] = await Promise.all([
        fetch(`${API_BASE_URL}/onboarding/${id}`, { headers }),
        fetch(`${API_BASE_URL}/technical/${id}`, { headers }),
        fetch(`${API_BASE_URL}/secrets/${id}`, { headers }),
        fetch(`${API_BASE_URL}/secrets-onboarding/${id}`, { headers })
      ]);

      const onbJson = await onbRes.json();
      const techJson = await techRes.json();
      const secJson = await secRes.json();
      const secOnbJson = await secOnbRes.json();

      const onbData = onbJson.data ? JSON.parse(onbJson.data) : {};
      const techData = techJson.data ? JSON.parse(techJson.data) : {};
      const secData = secJson.data ? JSON.parse(secJson.data) : {};
      const secOnbData = secOnbJson.data ? JSON.parse(secOnbJson.data) : {};

      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'PXM Manager';
      workbook.created = new Date();

      // --- Styles ---
      const headerFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; // Indigo 600
      const headerFont: ExcelJS.Font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      const sectionFont: ExcelJS.Font = { bold: true, size: 14, color: { argb: 'FF1E293B' } }; // Slate 800
      const borderStyle: ExcelJS.Borders = {
          top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // --- 1. Status & Onboarding ---
      const wsOverview = workbook.addWorksheet("Status & Onboarding");
      wsOverview.columns = [
          { header: 'Bereich', key: 'area', width: 25 },
          { header: 'Feld', key: 'field', width: 40 },
          { header: 'Wert', key: 'value', width: 60 }
      ];
      
      const ovHeader = wsOverview.getRow(1);
      ovHeader.font = headerFont;
      ovHeader.fill = headerFill;
      ovHeader.height = 24;
      ovHeader.alignment = { vertical: 'middle' };

      const overviewRows = [
        { area: "Stammdaten", field: "Anwendungsname", value: editingRow.Name },
        { area: "Stammdaten", field: "ICTO", value: editingRow.ICTO },
        { area: "Stammdaten", field: "Kritikalität", value: editingRow.Kritikalität },
        { area: "Stammdaten", field: "tAV", value: editingRow.tAV },
        ...Object.entries(onbData).map(([k, v]) => ({ area: "Onboarding", field: k, value: typeof v === 'boolean' ? (v ? 'Ja' : 'Nein') : v }))
      ];

      overviewRows.forEach(r => {
          const row = wsOverview.addRow(r);
          row.getCell(1).font = { bold: true, color: { argb: 'FF475569' } };
          row.eachCell(cell => { 
              cell.border = borderStyle; 
              cell.alignment = { vertical: 'top', wrapText: true }; 
          });
      });

      // --- Helper for Tables ---
      const addTableToSheet = (ws: ExcelJS.Worksheet, title: string, data: any[], columns: {key: string, label: string}[]) => {
          // Add Title
          const titleRow = ws.addRow([title]);
          titleRow.font = sectionFont;
          titleRow.height = 30;
          titleRow.alignment = { vertical: 'bottom' };
          
          // Add Header
          const headerRow = ws.addRow(columns.map(c => c.label));
          headerRow.font = headerFont;
          headerRow.fill = headerFill;
          headerRow.height = 24;
          headerRow.alignment = { vertical: 'middle' };
          
          // Add Data
          if (data && data.length > 0) {
              data.forEach(item => {
                  const rowValues = columns.map(col => {
                      const val = item[col.key];
                      if (Array.isArray(val)) return val.join('; ');
                      return typeof val === 'boolean' ? (val ? 'Ja' : 'Nein') : val;
                  });
                  const row = ws.addRow(rowValues);
                  row.eachCell(cell => { 
                      cell.border = borderStyle;
                      cell.alignment = { vertical: 'top', wrapText: true };
                  });
              });
          } else {
              const row = ws.addRow(["Keine Einträge vorhanden"]);
              row.getCell(1).font = { italic: true, color: { argb: 'FF94A3B8' } };
              ws.mergeCells(row.number, 1, row.number, columns.length);
              row.getCell(1).border = borderStyle;
          }

          // Add Spacing
          ws.addRow([]);
      };

      // --- 2. Technische Struktur ---
      const wsTech = workbook.addWorksheet("Technische Struktur");
      
      addTableToSheet(wsTech, "1. Server / Betriebssysteme", techData.servers, [
          { key: 'serverName', label: 'Servername' }, { key: 'ip', label: 'IP-Adresse' }, { key: 'fqdn', label: 'Adresse / FQDN' }, 
          { key: 'stage', label: 'Stage' }, { key: 'dmz', label: 'DMZ' }, { key: 'desc', label: 'Beschreibung' }, 
          { key: 'os', label: 'Betriebssystem' }, { key: 'port', label: 'Zugriff Port' }, { key: 'expiry', label: 'Ablaufdatum' }
      ]);
      
      addTableToSheet(wsTech, "2. Datenbanken / Server", techData.databases, [
          { key: 'serverName', label: 'Servername' }, { key: 'ip', label: 'IP-Adresse' }, { key: 'fqdn', label: 'Adresse / FQDN' },
          { key: 'stage', label: 'Stage' }, { key: 'dbType', label: 'Datenbanktyp' }, { key: 'instance', label: 'Instanz' },
          { key: 'product', label: 'Produkt' }, { key: 'port', label: 'Zugriff Port' }
      ]);

      addTableToSheet(wsTech, "3. Portfreischaltungen", techData.ports, [
          { key: 'fromServer', label: 'Von Server' }, { key: 'fromStage', label: 'Von Stage' }, { key: 'toServer', label: 'Nach Server' }, 
          { key: 'toIp', label: 'Nach IP' }, { key: 'toStage', label: 'Nach Stage' }, { key: 'port', label: 'Port/Protokoll' }, 
          { key: 'provider', label: 'Dienstleister' }, { key: 'interfaceId', label: 'Schnittstellen-ID' }, { key: 'comment', label: 'Kommentar' }
      ]);

      addTableToSheet(wsTech, "4. Safe-Struktur (CyberArk)", techData.safes, [
          { key: 'userGroup', label: 'Nutzergruppe' }, { key: 'safeName', label: 'Safe Name (fachlich)' }, { key: 'safeDesc', label: 'Safe Beschreibung' }, 
          { key: 'techSafeName', label: 'Technischer Safe Name' }, { key: 'adGroup', label: 'AD-Gruppe' }, { key: 'adGroupDesc', label: 'Beschreibung AD-Gruppe' }, 
          { key: 'approver', label: 'Zweitgenehmiger' }, { key: 'sod', label: 'SoD-Hinweis' }
      ]);

      addTableToSheet(wsTech, "5. Mitglieder der Safes", techData.safeMembers, [
          { key: 'safeName', label: 'Safe Name' }, { key: 'adGroup', label: 'AD-Gruppe' }, { key: 'memberName', label: 'Name Mitglied' }, { key: 'identity', label: 'Primäre Identität' }
      ]);

      addTableToSheet(wsTech, "6. Shared Accounts", techData.sharedAccounts, [
          { key: 'bizName', label: 'Fachlicher Name' }, { key: 'techName', label: 'Technischer Name' }, { key: 'sam', label: 'sAMAccountName' }, 
          { key: 'login', label: 'Anmeldename' }, { key: 'desc', label: 'Beschreibung' }, { key: 'isAd', label: 'AD Account' }, 
          { key: 'owner', label: 'Accountbesitzer' }, { key: 'ownerId', label: 'Identität Besitzer' }
      ]);

      addTableToSheet(wsTech, "7. Berechtigungszuordnungen", techData.permissions, [
          { key: 'bizName', label: 'Fachlicher Account' }, { key: 'techName', label: 'Technischer Account' }, { key: 'roleId', label: 'Role ID' }, 
          { key: 'roleName', label: 'Role Displayname' }, { key: 'roleDesc', label: 'Role Description' }, { key: 'bizSystem', label: 'Fachliches System' }, 
          { key: 'bizSystemId', label: 'Fachliche System-ID' }
      ]);

      addTableToSheet(wsTech, "8. Shared Accounts zu Safe", techData.mapping, [
          { key: 'techAccount', label: 'Technischer Account' }, { key: 'bizAccount', label: 'Fachlicher Account' }, 
          { key: 'techSafe', label: 'Technischer Safe' }, { key: 'safeName', label: 'Fachlicher Safe' }
      ]);

      // Set generic width for tech sheet
      for(let i=1; i<=10; i++) wsTech.getColumn(i).width = 25;


      // --- 3. Secrets Inventar ---
      const wsSecrets = workbook.addWorksheet("Secrets Inventar");
      
      addTableToSheet(wsSecrets, "1. Secret Inventory", secData.inventory, [
          { key: 'category', label: 'Kategorie' }, { key: 'name', label: 'Name / ID' }, { key: 'owner', label: 'Secret Owner' },
          { key: 'holder', label: 'Secret Holder' }, { key: 'layer', label: 'Layer' }, { key: 'localOrAd', label: 'Lokal/AD' },
          { key: 'stage', label: 'Stage' }, { key: 'complexity', label: 'Komplexität' }, { key: 'autoRotation', label: 'Auto Rotation' },
          { key: 'rotationMech', label: 'Mechanismus' }, { key: 'frequency', label: 'Häufigkeit' }, { key: 'timeWindow', label: 'Zeitfenster' }
      ]);

      addTableToSheet(wsSecrets, "2. Safe- / Pfad-Struktur", secData.safes, [
          { key: 'userGroup', label: 'Nutzergruppe' }, { key: 'safeName', label: 'Safe Name' }, { key: 'safeDesc', label: 'Beschreibung' },
          { key: 'techSafeName', label: 'Tech. Safe Name' }, { key: 'adGroup', label: 'AD-Gruppe' }, { key: 'adGroupDesc', label: 'Beschreibung AD' },
          { key: 'approver', label: 'Zweitgenehmiger' }, { key: 'sod', label: 'SoD' }
      ]);

      addTableToSheet(wsSecrets, "3. Mitglieder der Safes", secData.members, [
          { key: 'safeName', label: 'Safe Name' }, { key: 'adGroup', label: 'AD-Gruppe' }, { key: 'memberName', label: 'Name Mitglied' }, { key: 'identity', label: 'Identität' }
      ]);

      addTableToSheet(wsSecrets, "4. Secrets zu Safe", secData.mapping, [
          { key: 'bizSecret', label: 'Fachl. Secret' }, 
          { key: 'safeName', label: 'Safe Name' }, { key: 'techSafe', label: 'Tech. Safe' }
      ]);

      for(let i=1; i<=12; i++) wsSecrets.getColumn(i).width = 20;


      // --- 4. Secrets Onboarding ---
      const wsSecOnb = workbook.addWorksheet("Secrets Onboarding");
      wsSecOnb.columns = [
          { header: 'Frage', key: 'question', width: 60 },
          { header: 'Antwort', key: 'answer', width: 60 }
      ];
      const secOnbHeader = wsSecOnb.getRow(1);
      secOnbHeader.font = headerFont;
      secOnbHeader.fill = headerFill;
      secOnbHeader.height = 24;
      secOnbHeader.alignment = { vertical: 'middle' };

      Object.entries(secOnbData).forEach(([k, v]) => {
          const row = wsSecOnb.addRow({ 
              question: k, 
              answer: Array.isArray(v) ? v.join(', ') : (typeof v === 'boolean' ? (v ? 'Ja' : 'Nein') : v) 
          });
          row.eachCell(cell => { 
              cell.border = borderStyle; 
              cell.alignment = { vertical: 'top', wrapText: true }; 
          });
      });


      // --- Download ---
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Gesamtexport_${editingRow.ICTO}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);

    } catch (e) {
      console.error(e);
      alert("Fehler beim Erstellen des Excel-Exports.");
    }
  };

  const exportToExcel = () => {
    // Prepare data for Excel using the currently filtered data
    const excelData = filteredData.map(row => {
      const entry: any = {};
      DEFAULT_HEADERS.forEach(h => {
        entry[h] = row[h] || '';
      });
      return entry;
    });

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "PAM Governance");
    
    // Generate filename with current date
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `PAM_Governance_Export_${dateStr}.xlsx`);
  };

  const renderField = (fieldName: string) => {
    const options = SELECT_OPTIONS[fieldName];
    const isTechnical = ["ICTO", "id", "tAV"].includes(fieldName);

    // AD User Picker für Personenfelder verwenden
    if (["tAV", "Stellvertreter tAV", "fAV", "Betriebsverantwortlicher"].includes(fieldName)) {
      return (
        <UserPicker 
          key={fieldName}
          label={fieldName}
          value={editingRow?.[fieldName] || ''}
          onChange={(val: string) => setEditingRow({...editingRow, [fieldName]: val})}
          readOnly={isReadOnly}
          className="space-y-1.5"
        />
      );
    }

    if (fieldName === "Objektpflege") {
        return (
            <MultiUserPicker
                key={fieldName}
                label={fieldName}
                value={editingRow?.[fieldName] || ''}
                onChange={(val: string) => setEditingRow({...editingRow, [fieldName]: val})}
                readOnly={isReadOnly}
                className="space-y-1.5"
            />
        );
    }
    
    // Special handling for Anbindungsvariante
    if (fieldName === "Anbindungsvariante") {
        return (
            <div key={fieldName} className="space-y-1.5">
                <label className="block text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-[0.1em] ml-1">
                  {fieldName}
                </label>
                <div className="relative">
                    <input 
                        type="text" 
                        className="w-full p-3.5 bg-slate-100 border border-slate-200 rounded-xl text-slate-500 font-semibold text-sm outline-none cursor-not-allowed"
                        value={onboardingVariant || 'Nicht definiert'}
                        disabled
                        readOnly
                    />
                    <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
                        <Lock className="w-4 h-4" />
                    </div>
                </div>
                <p className="text-[10px] text-slate-400 ml-1 flex items-center gap-1">
                    <Info className="w-3 h-3" />
                    Wird aus PAM Onboarding ermittelt
                </p>
            </div>
        );
    }

    return (
      <div key={fieldName} className="space-y-1.5">
        <label className="block text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-[0.1em] ml-1">
          {fieldName}
        </label>
        {options ? (
          <div className="relative group">
            <select 
              className={`w-full p-3.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl focus:bg-white dark:focus:bg-slate-900 focus:ring-4 focus:ring-indigo-50 focus:border-indigo-500 transition-all outline-none font-semibold text-sm appearance-none cursor-pointer dark:text-slate-200 ${isTechnical ? 'font-mono' : ''}`}
              value={editingRow?.[fieldName] || ''}
              onChange={(e) => setEditingRow({...editingRow, [fieldName]: e.target.value})}
              disabled={isReadOnly}
            >
              <option value="">Wählen...</option>
              {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
            </select>
            <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
          </div>
        ) : (
          <textarea 
            rows={1}
            className={`w-full p-3.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl focus:bg-white dark:focus:bg-slate-900 focus:ring-4 focus:ring-indigo-50 focus:border-indigo-500 transition-all outline-none font-semibold text-sm resize-none dark:text-slate-200 ${isTechnical ? 'font-mono' : ''}`}
            value={editingRow?.[fieldName] || ''}
            onChange={(e) => setEditingRow({...editingRow, [fieldName]: e.target.value})}
            disabled={isReadOnly}
            onInput={(e: any) => {
              e.target.style.height = 'auto';
              e.target.style.height = e.target.scrollHeight + 'px';
            }}
            maxLength={500}
          />
        )}
      </div>
    );
  };

  if (!user) {
    return (
        <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 w-full max-w-md rounded-xl shadow-2xl p-10 animate-in zoom-in-95 duration-300 border border-slate-100 dark:border-slate-800">
            <div className="flex justify-center mb-8">
                <div className="bg-indigo-600 p-5 rounded-lg shadow-xl shadow-indigo-200">
                    <Database className="w-10 h-10 text-white" />
                </div>
            </div>
            <h2 className="text-3xl font-black mb-2 text-slate-900 text-center tracking-tight">PXM Manager</h2>
            <p className="text-slate-400 text-center mb-10 font-medium">Bitte melden Sie sich an, um fortzufahren.</p>
            <form onSubmit={handleLoginSubmit} className="space-y-5">
              <div>
                <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-2 ml-1">Benutzername</label>
                <input 
                  autoFocus
                  className="w-full p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:bg-white dark:focus:bg-slate-900 focus:ring-4 focus:ring-indigo-50 focus:border-indigo-500 outline-none font-bold text-slate-700 dark:text-slate-200 transition-all"
                  value={loginCreds.username}
                  onChange={e => setLoginCreds({...loginCreds, username: e.target.value})}
                  placeholder="Benutzername"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-2 ml-1">Passwort</label>
                <input 
                  type="password"
                  className="w-full p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:bg-white dark:focus:bg-slate-900 focus:ring-4 focus:ring-indigo-50 focus:border-indigo-500 outline-none font-bold text-slate-700 dark:text-slate-200 transition-all"
                  value={loginCreds.password}
                  onChange={e => setLoginCreds({...loginCreds, password: e.target.value})}
                  placeholder="••••••••"
                />
              </div>
              <button type="submit" className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg font-bold shadow-xl shadow-indigo-200 transition-all mt-6 active:scale-[0.98] flex justify-center items-center gap-2">
                <LogIn className="w-5 h-5" /> Anmelden
              </button>
            </form>
          </div>
        </div>
    );
  }

  return (
    <div className="h-screen bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 selection:bg-indigo-100 selection:text-indigo-900 flex flex-col overflow-hidden">
      <header className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-md border-b border-slate-100 dark:border-slate-800 px-8 py-5 flex justify-between items-center shadow-sm shrink-0 z-40">
        <div className="flex items-center gap-5">
          <div className={`p-2.5 rounded-lg text-white transition-all duration-700 shadow-lg ${
            connectionState === 'online' ? 'bg-indigo-600 shadow-indigo-200' : 
            connectionState === 'api_only' ? 'bg-amber-500 shadow-amber-200' : 'bg-rose-500 shadow-rose-200'
          }`}>
            <Database className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-2xl font-black tracking-[-0.03em] leading-none mb-1 text-slate-900 dark:text-white">PXM Manager</h1>
            <div className="flex items-center gap-2">
               <span className="text-[10px] font-extrabold uppercase tracking-[0.15em] flex items-center gap-1.5">
                 {connectionState === 'online' ? (
                   <span className="text-emerald-500 flex items-center gap-1"><Wifi className="w-3.5 h-3.5" /> SQL ONLINE</span>
                 ) : connectionState === 'api_only' ? (
                   <span className="text-amber-500 flex items-center gap-1"><AlertTriangle className="w-3.5 h-3.5" /> DB KONFIG</span>
                 ) : (
                   <span className="text-rose-500 flex items-center gap-1"><WifiOff className="w-3.5 h-3.5" /> OFFLINE</span>
                 )}
               </span>
            </div>
          </div>
        </div>
        
        <div className="flex items-center gap-4">
          <button onClick={() => setDarkMode(!darkMode)} className="p-3 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg transition-all active:scale-95 group" title={darkMode ? "Light Mode" : "Dark Mode"}>
              {darkMode ? <Sun className="w-5 h-5 text-amber-400" /> : <Moon className="w-5 h-5 text-slate-400 group-hover:text-slate-600" />}
          </button>

          <button onClick={handleLogout} className="px-5 py-2 rounded-lg font-bold flex items-center gap-3 bg-slate-50 dark:bg-slate-900 text-slate-500 dark:text-slate-400 border border-slate-100 dark:border-slate-800 hover:bg-rose-50 dark:hover:bg-rose-900/30 hover:text-rose-600 dark:hover:text-rose-400 hover:border-rose-100 dark:hover:border-rose-800 transition-all">
            <div className="flex flex-col items-end">
                <span className="text-[10px] font-black uppercase tracking-wider text-indigo-500">{role === 'admin' ? 'Admin' : role === 'maintenance' ? 'Pflege' : 'Read Only'}</span>
                <span className="text-sm leading-none">{user}</span>
            </div>
            <div className="h-8 w-px bg-slate-200 mx-1"></div>
            <LogOut className="w-4 h-4" />
          </button>

          <button onClick={checkConnectivity} className="p-3 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg transition-all active:scale-95 group" title="Aktualisieren">
            <RefreshCw className={`w-5 h-5 text-slate-400 group-hover:text-slate-600 ${loading ? 'animate-spin' : ''}`} />
          </button>
          
          <button 
            onClick={exportToExcel}
            disabled={data.length === 0}
            className="px-5 py-3 rounded-lg font-extrabold flex items-center gap-2.5 border-2 border-slate-100 dark:border-slate-800 hover:border-slate-200 dark:hover:border-slate-700 bg-white dark:bg-slate-900 text-slate-600 dark:text-slate-300 transition-all active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed shadow-sm"
          >
            <FileDown className="w-5 h-5" /> Export
          </button>

          {!isReadOnly && role === 'admin' && (
          <button 
            disabled={connectionState !== 'online'}
            onClick={() => { setEditingRow({}); setIsModalOpen(true); setViewMode('form'); }}
            className={`px-6 py-3 rounded-lg font-extrabold flex items-center gap-2.5 shadow-xl transition-all active:scale-[0.98] ${
              connectionState === 'online' 
                ? 'bg-slate-900 hover:bg-black text-white' 
                : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed shadow-none'
            }`}
          >
            <Plus className="w-5 h-5 stroke-[3px]" /> Neu
          </button>
          )}
        </div>
      </header>

      <main className="p-10 flex-1 overflow-hidden flex flex-col">
        {connectionState === 'api_only' && (
          <div className="mb-10 bg-amber-50 border border-amber-100 p-6 rounded-xl flex gap-5 items-start max-w-4xl mx-auto shadow-sm">
            <div className="bg-amber-500 p-2.5 rounded-lg text-white shrink-0 shadow-lg shadow-amber-100"><Settings className="w-5 h-5" /></div>
            <div>
              <h3 className="font-extrabold text-amber-900 mb-1 tracking-tight">SQL Server Verbindung unvollständig</h3>
              <p className="text-sm text-amber-800/80 mb-2 leading-relaxed font-medium">{lastError}</p>
            </div>
          </div>
        )}

        {connectionState === 'offline' ? (
          <div className="bg-white p-16 rounded-xl shadow-2xl border border-slate-100 text-center max-w-2xl mx-auto mt-20">
            <div className="bg-rose-50 w-24 h-24 rounded-xl flex items-center justify-center mx-auto mb-8 text-rose-500 shadow-inner">
              <WifiOff className="w-12 h-12" />
            </div>
            <h2 className="text-3xl font-black mb-4 tracking-tight">Backend nicht erreichbar</h2>
            <p className="text-slate-500 font-medium mb-10">Der Node.js Server wurde unter Port 3001 nicht gefunden.</p>
            <button onClick={checkConnectivity} className="bg-indigo-600 text-white px-10 py-4 rounded-lg font-black hover:bg-indigo-700 transition-all shadow-xl shadow-indigo-100 active:scale-95">Verbindung prüfen</button>
          </div>
        ) : (
          <div className="flex flex-col h-full gap-8">
            <div className="relative max-w-lg group shrink-0">
              <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 w-5 h-5 group-focus-within:text-indigo-500 transition-colors stroke-[2.5px]" />
              <input 
                type="text" 
                placeholder="Katalog durchsuchen..." 
                className="w-full pl-14 pr-6 py-4 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:ring-[6px] focus:ring-indigo-50 focus:border-indigo-500 outline-none transition-all shadow-sm font-semibold placeholder:text-slate-300 dark:placeholder:text-slate-500 dark:text-slate-200"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>

            <div className="bg-white dark:bg-slate-900 rounded-xl shadow-2xl shadow-slate-200/50 dark:shadow-none border border-slate-100 dark:border-slate-800 overflow-hidden flex-1 flex flex-col">
              <div className="overflow-auto custom-scrollbar flex-1">
                <table className={`w-full border-collapse ${isTableExpanded ? 'min-w-[2200px]' : 'min-w-full'}`}>
                  <thead>
                    <tr className="bg-slate-50/40 dark:bg-slate-800/40 border-b border-slate-100 dark:border-slate-800">
                      <th className="sticky left-0 top-0 bg-white dark:bg-slate-900 z-30 px-4 py-4 text-center w-[80px] min-w-[80px] text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] shadow-[0_1px_0_#f1f5f9] dark:shadow-[0_1px_0_#1e293b] align-top">
                        <div className="mb-3">Status</div>
                      </th>
                      <th className="sticky left-[80px] top-0 bg-white dark:bg-slate-900 z-30 px-4 py-4 text-center w-[140px] min-w-[140px] text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] shadow-[0_1px_0_#f1f5f9] dark:shadow-[0_1px_0_#1e293b] align-top">
                        <div className="mb-3">PAM Onboarding</div>
                      </th>
                      {visibleHeaders.map(h => {
                        let stickyClass = "sticky top-0 bg-white dark:bg-slate-900 z-20";
                        if (h === "ICTO") stickyClass = "sticky left-[218px] top-0 bg-white dark:bg-slate-900 z-30 w-[120px] min-w-[120px] border-l border-slate-50 dark:border-slate-800";
                        if (h === "Name") stickyClass = "sticky left-[337px] top-0 bg-white dark:bg-slate-900 z-30 w-[250px] min-w-[250px] border-r border-slate-50 dark:border-slate-800 shadow-[6px_0_12px_-6px_rgba(0,0,0,0.04),0_1px_0_#f1f5f9] dark:shadow-[6px_0_12px_-6px_rgba(0,0,0,0.04),0_1px_0_#1e293b]";
                        
                        return (
                        <th key={h} className={`${stickyClass} text-left px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] whitespace-nowrap shadow-[0_1px_0_#f1f5f9] dark:shadow-[0_1px_0_#1e293b] align-top`}>
                          <div className="mb-3 block">{h}</div>
                          <input 
                            type="text" 
                            className="w-full p-2 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-md text-xs font-medium text-slate-600 dark:text-slate-300 focus:ring-2 focus:ring-indigo-500 outline-none normal-case tracking-normal placeholder:text-slate-300 dark:placeholder:text-slate-500"
                            placeholder={`Filter ${h}...`}
                            value={filters[h] || ''}
                            onChange={e => setFilters(prev => ({...prev, [h]: e.target.value}))}
                          />
                        </th>
                      )})}
                      <th className="sticky top-0 bg-white dark:bg-slate-900 z-20 px-2 py-4 text-center shadow-[0_1px_0_#f1f5f9] dark:shadow-[0_1px_0_#1e293b] align-top w-12">
                        <button 
                            onClick={() => setIsTableExpanded(!isTableExpanded)}
                            className="p-2 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-indigo-600 transition-colors mt-1"
                            title={isTableExpanded ? "Spalten einklappen" : "Spalten ausklappen"}
                        >
                            <ChevronRight className={`w-5 h-5 transition-transform ${isTableExpanded ? 'rotate-180' : ''}`} />
                        </button>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={visibleHeaders.length + 3} className="py-24 text-center">
                          <div className="flex flex-col items-center gap-3">
                            {loading ? (
                              <RefreshCw className="w-10 h-10 text-indigo-200 animate-spin" />
                            ) : (
                              <Database className="w-12 h-12 text-slate-100" />
                            )}
                            <p className="text-slate-400 font-extrabold tracking-tight">
                              {loading ? 'DATEN WERDEN GELADEN...' : (data.length === 0 ? 'KEINE DATEN VORHANDEN' : 'KEINE TREFFER FÜR DIESEN FILTER')}
                            </p>
                          </div>
                        </td>
                      </tr>
                    ) : filteredData.map((row, i) => (
                      <tr key={row.id || i} className="hover:bg-indigo-50/30 dark:hover:bg-indigo-900/20 transition-colors group">
                        <td className="sticky left-0 bg-white dark:bg-slate-900 group-hover:bg-indigo-50/50 dark:group-hover:bg-indigo-900/50 z-10 px-4 py-4 text-center w-[80px] min-w-[80px] whitespace-nowrap">
                          <button 
                            onClick={() => { setEditingRow(row); setIsModalOpen(true); setViewMode('form'); }} 
                            className="p-2.5 text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-md transition-all shadow-sm hover:shadow-md active:scale-90"
                            title="Schnittstelleninformationen"
                          >
                            <FileCog className="w-5 h-5" />
                          </button>
                        </td>
                        <td className="sticky left-[80px] bg-white dark:bg-slate-900 group-hover:bg-indigo-50/50 dark:group-hover:bg-indigo-900/50 z-10 px-4 py-4 text-center w-[140px] min-w-[140px] whitespace-nowrap">
                          <button 
                            onClick={() => { setUnifiedRow(row); setInitialTab('onboarding'); setModalMode('pam'); }} 
                            className="p-2.5 text-rose-600 bg-rose-50 hover:bg-rose-100 rounded-md transition-all shadow-sm hover:shadow-md active:scale-90"
                            title="Status Onboarding"
                          >
                            <Rocket className="w-5 h-5" />
                          </button>
                          <button 
                            onClick={() => { setUnifiedRow(row); setInitialTab('secrets-onboarding'); setModalMode('secrets'); }} 
                            className="p-2.5 ml-2 text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-md transition-all shadow-sm hover:shadow-md active:scale-90"
                            title="Secrets Management"
                          >
                            <KeyRound className="w-5 h-5" />
                          </button>
                        </td>
                        {visibleHeaders.map(h => {
                          const isTech = ["ICTO", "tAV"].includes(h);
                          let stickyClass = "";
                          if (h === "ICTO") stickyClass = "sticky left-[218px] bg-white dark:bg-slate-900 group-hover:bg-indigo-50/50 dark:group-hover:bg-indigo-900/50 z-10 w-[120px] min-w-[120px] border-l border-slate-50 dark:border-slate-800";
                          if (h === "Name") stickyClass = "sticky left-[337px] bg-white dark:bg-slate-900 group-hover:bg-indigo-50/50 dark:group-hover:bg-indigo-900/50 z-10 w-[250px] min-w-[250px] border-r border-slate-50 dark:border-slate-800 shadow-[6px_0_12px_-6px_rgba(0,0,0,0.04)]";

                          return (
                            <td key={h} className={`px-8 py-5 text-sm font-bold text-slate-600 dark:text-slate-300 whitespace-nowrap max-w-md overflow-hidden text-ellipsis ${isTech ? 'font-mono text-indigo-500/80 dark:text-indigo-400' : ''} ${stickyClass}`}>
                              {row[h] || <span className="text-slate-200">/</span>}
                            </td>
                          );
                        })}
                        <td className="px-2 py-5 whitespace-nowrap"></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>

      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-10">
          <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-xl transition-opacity" onClick={() => setIsModalOpen(false)}></div>
          <div className="relative bg-white dark:bg-slate-900 w-full max-w-5xl rounded-xl shadow-[0_32px_64px_-16px_rgba(0,0,0,0.2)] flex flex-col max-h-[94vh] border border-slate-200 dark:border-slate-800 overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="px-10 py-8 border-b border-slate-50 dark:border-slate-800 flex justify-between items-center bg-white dark:bg-slate-900 sticky top-0 z-10">
              <div>
                <h2 className="text-3xl font-black tracking-tighter text-slate-900 dark:text-white">
                  {editingRow?.id ? 'Status Onboarding' : 'Neue Erfassung'}
                </h2>
                <div className="flex items-center gap-2 mt-1.5">
                   <div className="w-2 h-2 rounded-full bg-indigo-500"></div>
                   <p className="text-slate-400 text-[10px] font-black uppercase tracking-[0.2em]">PAM Governance Framework</p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                {editingRow?.id && (
                  <>
                  <button 
                    onClick={handleFullExcelExport}
                    className="p-3 hover:bg-emerald-50 rounded-lg transition-all group text-emerald-600"
                    title="Gesamtexport Excel"
                  >
                    <FileSpreadsheet className="w-6 h-6" />
                  </button>
                  <button 
                    onClick={() => viewMode === 'history' ? setViewMode('form') : handleFetchHistory(editingRow.id)}
                    className={`p-3 rounded-lg transition-all group ${viewMode === 'history' ? 'bg-indigo-50 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400' : 'hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-300 hover:text-slate-600'}`}
                    title="Änderungshistorie"
                  >
                    <HistoryIcon className="w-6 h-6" />
                  </button>
                  </>
                )}
              <button onClick={() => setIsModalOpen(false)} className="p-3 hover:bg-slate-50 dark:hover:bg-slate-800 rounded-lg transition-all group">
                <X className="w-7 h-7 text-slate-300 group-hover:text-slate-600" />
              </button>
              </div>
            </div>
            
            {viewMode === 'history' ? (
              <div className="flex-1 overflow-y-auto p-10 custom-scrollbar">
                <h3 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800 dark:text-slate-200">
                  <Clock className="w-6 h-6 text-indigo-500" /> Änderungsprotokoll
                </h3>
                <PaginatedHistoryList history={history} cardClassName="bg-slate-50 dark:bg-slate-800 p-5 rounded-lg border border-slate-100 dark:border-slate-700 hover:border-indigo-100 transition-colors" />
              </div>
            ) : (
              <form onSubmit={handleSave} className="flex-1 overflow-y-auto p-10 space-y-16 custom-scrollbar">
                {FIELD_GROUPS.map((group, idx) => (
                  <div key={idx} className="space-y-8">
                    <div className="flex items-center gap-4 border-b border-slate-100 dark:border-slate-800 pb-4">
                      <div className="p-2.5 bg-slate-50 dark:bg-slate-800 rounded-lg shadow-inner">{group.icon}</div>
                      <h3 className="text-xs font-black uppercase tracking-[0.25em] text-slate-500">{group.title}</h3>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-x-10 gap-y-8 px-2">
                      {group.fields.map(fieldName => renderField(fieldName))}
                    </div>
                  </div>
                ))}

                {/* Abnahme Bereich */}
                <div className="space-y-8 border-t border-slate-100 dark:border-slate-800 pt-8">
                    <div className="flex items-center gap-4">
                        <div className="p-2.5 bg-slate-50 dark:bg-slate-800 rounded-lg shadow-inner"><CheckCircle className="w-5 h-5 text-emerald-500" /></div>
                        <h3 className="text-xs font-black uppercase tracking-[0.25em] text-slate-500">Abnahme PAM Onboarding</h3>
                    </div>
                    
                    {editingRow?.AbnahmePAMOnboarding ? (
                        <div className="bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-100 dark:border-emerald-800 rounded-xl p-6 flex items-center justify-between gap-4">
                            <div className="flex items-center gap-4">
                                <div className="bg-emerald-100 p-3 rounded-full">
                                    <CheckCircle className="w-8 h-8 text-emerald-600" />
                                </div>
                                <div>
                                    <h4 className="font-bold text-emerald-900 text-lg">Abnahme erteilt</h4>
                                    <p className="text-emerald-700 font-medium">
                                        Durch {editingRow.AbnahmePAMOnboarding}
                                    </p>
                                </div>
                            </div>
                            {role === 'admin' && (
                                <button type="button" onClick={handleRevokeAcceptance} className="px-4 py-2 bg-white border border-emerald-200 text-emerald-700 hover:bg-emerald-100 hover:text-emerald-800 rounded-lg text-sm font-bold transition-colors shadow-sm">
                                    Zurückziehen
                                </button>
                            )}
                        </div>
                    ) : (
                        (role === 'maintenance') ? (
                            <div className="bg-slate-50 dark:bg-slate-800 border border-slate-100 dark:border-slate-700 rounded-xl p-6">
                                <p className="text-slate-500 dark:text-slate-400 mb-4 text-sm">Bitte bestätigen Sie die korrekte Umsetzung der CyberArk Anbindung nach erfolgreichem Test.</p>
                                <button type="button" onClick={handleAcceptance} className="w-full py-4 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg shadow-lg shadow-emerald-100 transition-all active:scale-[0.98] flex justify-center items-center gap-2"><CheckCircle className="w-5 h-5" /> Abnahme erteilen</button>
                            </div>
                        ) : <div className="p-4 text-center text-slate-400 italic bg-slate-50 dark:bg-slate-800 rounded-xl border border-slate-100 dark:border-slate-700">Abnahme noch ausstehend.</div>
                    )}
                </div>

                <div className="pb-10"></div>
              </form>
            )}

            <div className="p-10 bg-white/80 dark:bg-slate-900/80 backdrop-blur-md border-t border-slate-50 dark:border-slate-800 flex justify-end gap-5 sticky bottom-0 z-10">
              <button type="button" onClick={() => setIsModalOpen(false)} className="px-8 py-4 font-extrabold text-slate-400 hover:text-slate-600 transition-colors">Schließen</button>
              {viewMode === 'form' && (
              !isReadOnly && (
              <button 
                type="submit" 
                onClick={handleSave}
                disabled={saveStatus === 'saving'}
                className={`px-8 py-3 font-bold rounded-lg shadow-lg transition-all active:scale-95 flex items-center gap-2 ${
                    saveStatus === 'success' 
                    ? 'bg-emerald-500 hover:bg-emerald-600 text-white shadow-emerald-200' 
                    : 'bg-indigo-600 hover:bg-indigo-700 text-white shadow-indigo-200'
                }`}
              >
                {saveStatus === 'saving' ? <RefreshCw className="animate-spin w-5 h-5" /> : 
                 saveStatus === 'success' ? <CheckCircle className="w-5 h-5" /> : 
                 <Save className="w-5 h-5" />} 
                {saveStatus === 'success' ? 'Gespeichert' : 'Speichern'}
              </button>
              )
              )}
            </div>
          </div>
        </div>
      )}

      {unifiedRow && (
        <UnifiedAppModal 
            isOpen={!!unifiedRow} 
            onClose={() => setUnifiedRow(null)} 
            governanceRow={unifiedRow} 
            user={user} 
            initialTab={initialTab}
            mode={modalMode}
            readOnly={isReadOnly}
        />
      )}

      {isAcceptanceConfirmOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 sm:p-10">
          <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm transition-opacity" onClick={() => setIsAcceptanceConfirmOpen(false)}></div>
          <div className="relative bg-white w-full max-w-lg rounded-xl shadow-2xl p-8 animate-in zoom-in-95 duration-200 border border-slate-100">
            <div className="flex flex-col items-center text-center gap-4">
                <div className="p-4 bg-emerald-50 rounded-full">
                    <CheckCircle className="w-10 h-10 text-emerald-500" />
                </div>
                <h3 className="text-xl font-black text-slate-900">Abnahme bestätigen</h3>
                <p className="text-slate-500 font-medium leading-relaxed">
                    Ich bestätige die technische und fachliche Korrektheit der Cyberark Anbindung. Diese wurde erfolgreich getestet und übergeben.
                </p>
                <div className="grid grid-cols-2 gap-4 w-full mt-4">
                    <button 
                        onClick={() => setIsAcceptanceConfirmOpen(false)}
                        className="py-3 px-4 bg-slate-50 hover:bg-slate-100 text-slate-600 font-bold rounded-lg transition-colors"
                    >
                        Abbrechen
                    </button>
                    <button 
                        onClick={confirmAcceptance}
                        className="py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg shadow-lg shadow-emerald-100 transition-all active:scale-[0.98]"
                    >
                        Bestätigen
                    </button>
                </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = createRoot(rootElement);
  root.render(<App />);
}
