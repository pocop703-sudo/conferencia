import React, { useState, useMemo, useEffect, useRef } from 'react';
import type { ManifestBatch, SheetData, HistoryRecord } from './types';
import {
  normalizeCode,
  parseManifestText,
  readManifestFile,
  DEMO_MANIFEST_TEXT,
} from './utils/manifestParser';
import { calculateJulianOrderNumber } from './utils/excelProcessor';
import {
  extractSpreadsheetId,
  fetchSpreadsheetMeta,
  readGoogleSheet,
  batchUpdateGoogleSheet,
  type SpreadsheetMeta,
} from './utils/googleSheets';
import {
  initAuth,
  googleSignIn,
  switchGoogleAccount,
  logoutGoogle,
  getAccessToken,
  firebaseProjectId,
} from './utils/googleAuth';
import { getHistory, addHistory, clearHistory } from './utils/historyStorage';
import type { User } from 'firebase/auth';
import { VercelAuthHelpModal } from './components/VercelAuthHelpModal';
import {
  FileSpreadsheet,
  CheckCircle2,
  AlertTriangle,
  ShieldAlert,
  History,
  Route,
  Upload,
  RotateCcw,
  Sparkles,
  Trash2,
  ArrowRight,
  Info,
  Globe,
  ExternalLink,
  LogOut,
  Loader2,
  Settings2,
  Check,
  LogIn,
  Mail,
  UserCheck,
  Users,
} from 'lucide-react';

interface StatCardProps {
  label: string;
  value: number | string;
  tone?: 'default' | 'primary' | 'success' | 'warning';
}

function StatCard({ label, value, tone = 'default' }: StatCardProps) {
  const toneClass =
    tone === 'success'
      ? 'text-success'
      : tone === 'warning'
      ? 'text-warning'
      : tone === 'primary'
      ? 'text-primary'
      : 'text-foreground';

  return (
    <div className="rounded-lg border border-border bg-card p-4 shadow-[var(--shadow-card)]">
      <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{label}</p>
      <p className={`mt-1 text-2xl font-bold tabular-nums ${toneClass}`}>{value}</p>
    </div>
  );
}

const STORAGE_SAVED_URL_KEY = 'gs_saved_url';
const STORAGE_SAVED_TAB_KEY = 'gs_saved_tab';

export default function App() {
  const [activeTab, setActiveTab] = useState<'processar' | 'historico' | 'planilha'>('processar');
  const [selectedDate, setSelectedDate] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [manifestFile, setManifestFile] = useState<File | null>(null);
  const [sheetData, setSheetData] = useState<SheetData | null>(null);
  const [manifestBatches, setManifestBatches] = useState<ManifestBatch[]>([]);
  const [step, setStep] = useState<number>(1);
  const [statusMessage, setStatusMessage] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [historyList, setHistoryList] = useState<HistoryRecord[]>([]);
  const [filledStoreKeys, setFilledStoreKeys] = useState<string[]>([]);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // Google Sheets & Google Auth state
  const [googleUser, setGoogleUser] = useState<User | null>(null);
  const [googleAccessToken, setGoogleAccessToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);
  const [googleSheetsUrl, setGoogleSheetsUrl] = useState<string>(() => {
    return localStorage.getItem(STORAGE_SAVED_URL_KEY) || '';
  });
  const [spreadsheetMeta, setSpreadsheetMeta] = useState<SpreadsheetMeta | null>(null);
  const [selectedTabName, setSelectedTabName] = useState<string>(() => {
    return localStorage.getItem(STORAGE_SAVED_TAB_KEY) || '';
  });
  const [isEditingSheetUrl, setIsEditingSheetUrl] = useState<boolean>(() => {
    return !localStorage.getItem(STORAGE_SAVED_URL_KEY);
  });
  const [saveAsDefault, setSaveAsDefault] = useState<boolean>(true);
  const [showConfirmModal, setShowConfirmModal] = useState<boolean>(false);
  const [isSavingGoogleSheets, setIsSavingGoogleSheets] = useState<boolean>(false);
  const [unauthorizedDomain, setUnauthorizedDomain] = useState<string | null>(null);
  const [showVercelHelpModal, setShowVercelHelpModal] = useState<boolean>(false);

  const manifestInputRef = useRef<HTMLInputElement>(null);

  const orderNumber = calculateJulianOrderNumber(selectedDate);

  // Listen to Firebase Google Auth state & persistent token
  useEffect(() => {
    const unsubscribe = initAuth(
      (user, token) => {
        setGoogleUser(user);
        setGoogleAccessToken(token);
      },
      () => {
        setGoogleUser(null);
        setGoogleAccessToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  // Auto-connect to saved Google Sheet when token becomes available
  useEffect(() => {
    const savedUrl = localStorage.getItem(STORAGE_SAVED_URL_KEY);
    if (savedUrl && googleAccessToken && !sheetData) {
      connectToGoogleSheet(savedUrl, localStorage.getItem(STORAGE_SAVED_TAB_KEY) || '', googleAccessToken);
    }
  }, [googleAccessToken]);

  // Flatten all entries from added batches
  const allEntries = useMemo(() => {
    return manifestBatches.length ? manifestBatches.flatMap((b) => b.entries) : null;
  }, [manifestBatches]);

  // Lookup map of stores in the active Google Sheets spreadsheet
  const sheetStoreMap = useMemo(() => {
    const map = new Map<string, { display: string; route: string }>();
    sheetData?.routes.forEach((route) => {
      route.stores.forEach((store) => {
        map.set(store.key, { display: store.display, route: route.name });
      });
    });
    return map;
  }, [sheetData]);

  // Stores in manifest that exist in spreadsheet
  const matchedStores = useMemo(() => {
    return (allEntries ?? []).filter((entry) => sheetStoreMap.has(normalizeCode(entry.code)));
  }, [allEntries, sheetStoreMap]);

  // Stores in manifest that are missing from spreadsheet
  const unmatchedStores = useMemo(() => {
    return (allEntries ?? []).filter((entry) => !sheetStoreMap.has(normalizeCode(entry.code)));
  }, [allEntries, sheetStoreMap]);

  // Stores in spreadsheet that were not present in manifest
  const unmanifestedStores = useMemo(() => {
    return (sheetData?.routes ?? []).flatMap((route) =>
      route.stores.filter((store) => !(allEntries ?? []).some((e) => normalizeCode(e.code) === store.key))
    );
  }, [sheetData, allEntries]);

  // Totals
  const totalVolumes = useMemo(() => {
    return (allEntries ?? []).reduce((acc, cur) => acc + cur.volumes, 0);
  }, [allEntries]);

  // Handle Google Login
  async function handleGoogleLogin() {
    try {
      setIsLoggingIn(true);
      setErrorMessage('');
      setUnauthorizedDomain(null);
      const res = await googleSignIn();
      if (res) {
        setGoogleUser(res.user);
        setGoogleAccessToken(res.accessToken);
        setStatusMessage(`Conectado com sucesso: ${res.user.email}`);

        // If URL is present, connect immediately
        const targetUrl = googleSheetsUrl.trim() || localStorage.getItem(STORAGE_SAVED_URL_KEY) || '';
        if (targetUrl) {
          connectToGoogleSheet(targetUrl, selectedTabName, res.accessToken);
        }
      }
    } catch (err: any) {
      if (err?.code === 'auth/unauthorized-domain' || err?.message?.includes('unauthorized-domain')) {
        const domain = err.domain || (typeof window !== 'undefined' ? window.location.hostname : 'seu-dominio');
        setUnauthorizedDomain(domain);
        setShowVercelHelpModal(true);
      }
      setErrorMessage(`Falha na autenticação do Google: ${err.message}`);
    } finally {
      setIsLoggingIn(false);
    }
  }

  // Handle Switch Google Account (forces account selector so user can choose or type any account)
  async function handleSwitchAccount() {
    try {
      setIsLoggingIn(true);
      setErrorMessage('');
      setUnauthorizedDomain(null);
      const res = await switchGoogleAccount();
      if (res) {
        setGoogleUser(res.user);
        setGoogleAccessToken(res.accessToken);
        setStatusMessage(`Conta alterada com sucesso! Conectado como ${res.user.email}`);

        const targetUrl = googleSheetsUrl.trim() || localStorage.getItem(STORAGE_SAVED_URL_KEY) || '';
        if (targetUrl) {
          connectToGoogleSheet(targetUrl, selectedTabName, res.accessToken);
        }
      }
    } catch (err: any) {
      if (err?.code === 'auth/unauthorized-domain' || err?.message?.includes('unauthorized-domain')) {
        const domain = err.domain || (typeof window !== 'undefined' ? window.location.hostname : 'seu-dominio');
        setUnauthorizedDomain(domain);
        setShowVercelHelpModal(true);
      }
      setErrorMessage(`Falha ao trocar de conta: ${err.message}`);
    } finally {
      setIsLoggingIn(false);
    }
  }

  // Handle Google Logout
  async function handleGoogleLogout() {
    await logoutGoogle();
    setGoogleUser(null);
    setGoogleAccessToken(null);
    setSpreadsheetMeta(null);
    setSheetData(null);
    setStatusMessage('Desconectado do Google.');
  }

  // Reset conference for new batch
  function handleReset() {
    setManifestBatches([]);
    setFilledStoreKeys([]);
    setManifestFile(null);
    setStep(1);
    setStatusMessage('');
    setErrorMessage('');
    if (manifestInputRef.current) manifestInputRef.current.value = '';
  }

  // Internal connect logic to Google Sheet
  async function connectToGoogleSheet(urlOrId: string, tabName: string, token: string) {
    const spreadsheetId = extractSpreadsheetId(urlOrId);
    if (!spreadsheetId) {
      setErrorMessage('Por favor, informe um link válido do Google Sheets ou o ID da planilha.');
      return;
    }

    try {
      setIsLoading(true);
      setStatusMessage('Buscando informações da planilha no Google Sheets...');
      const meta = await fetchSpreadsheetMeta(spreadsheetId, token);
      setSpreadsheetMeta(meta);

      const targetTab = tabName || meta.sheets[0]?.title || 'Página1';
      setSelectedTabName(targetTab);

      setStatusMessage(`Lendo rotas e lojas da aba '${targetTab}'...`);
      const loaded = await readGoogleSheet(spreadsheetId, targetTab, token);
      loaded.spreadsheetTitle = meta.title;

      if (!loaded.routes.length) {
        throw new Error(
          `Nenhuma coluna de conferência (LOJA e VOL) encontrada na aba '${targetTab}'. Verifique se a planilha possui colunas LOJA e VOL.`
        );
      }

      setSheetData(loaded);
      setStatusMessage(`Conectado com sucesso: '${meta.title}' (${loaded.routes.length} rotas mapeadas).`);
      setIsEditingSheetUrl(false);

      if (saveAsDefault) {
        localStorage.setItem(STORAGE_SAVED_URL_KEY, urlOrId);
        localStorage.setItem(STORAGE_SAVED_TAB_KEY, targetTab);
      }
    } catch (err: any) {
      if (
        err.message?.includes('AUTH_401') ||
        err.message?.includes('não autenticada') ||
        err.message?.includes('Não autenticado') ||
        err.message?.includes('401')
      ) {
        setGoogleAccessToken(null);
      }
      setErrorMessage(err.message || 'Falha ao conectar com a planilha do Google Sheets');
      setSheetData(null);
    } finally {
      setIsLoading(false);
    }
  }

  // User click on connect
  async function handleConnectGoogleSheet() {
    setErrorMessage('');
    let token = googleAccessToken;
    if (!token) {
      token = await getAccessToken();
    }
    if (!token) {
      try {
        const authRes = await googleSignIn();
        if (authRes) {
          token = authRes.accessToken;
          setGoogleUser(authRes.user);
          setGoogleAccessToken(authRes.accessToken);
        }
      } catch (err: any) {
        if (err?.code === 'auth/unauthorized-domain' || err?.message?.includes('unauthorized-domain')) {
          const domain = err.domain || (typeof window !== 'undefined' ? window.location.hostname : 'seu-dominio');
          setUnauthorizedDomain(domain);
          setShowVercelHelpModal(true);
        }
        setErrorMessage(`Faça login com sua conta Google para continuar: ${err.message}`);
        return;
      }
    }

    if (!token) return;
    await connectToGoogleSheet(googleSheetsUrl, selectedTabName, token);
  }

  // Process manifest (PDF or image)
  async function processManifestFile(fileToProcess: File | null, useDemo = false) {
    setErrorMessage('');
    setIsLoading(true);
    try {
      setStatusMessage('Lendo manifesto...');
      const rawText = useDemo ? DEMO_MANIFEST_TEXT : await readManifestFile(fileToProcess!, setStatusMessage);
      const parsedEntries = parseManifestText(rawText);

      if (!parsedEntries.length) {
        throw new Error('Nenhum destino com volumes foi identificado no manifesto.');
      }

      // Check duplicates against current batches
      const existingKeys = new Set(manifestBatches.flatMap((b) => b.entries.map((e) => normalizeCode(e.code))));
      const newEntries = parsedEntries.filter((e) => !existingKeys.has(normalizeCode(e.code)));
      const duplicatesCount = parsedEntries.length - newEntries.length;

      if (!newEntries.length) {
        throw new Error('Todas as lojas deste manifesto já estão na conferência atual.');
      }

      const batchName = useDemo ? 'Manifesto demonstração' : (fileToProcess?.name ?? 'Manifesto');
      setManifestBatches((prev) => [
        ...prev,
        {
          id: crypto.randomUUID(),
          name: batchName,
          entries: newEntries,
        },
      ]);

      setManifestFile(null);
      if (manifestInputRef.current) manifestInputRef.current.value = '';

      setStatusMessage(
        `Manifesto processado — ${newEntries.length} lojas adicionadas${
          duplicatesCount ? ` (${duplicatesCount} já estavam na conferência)` : ''
        }`
      );
      setStep(2);
    } catch (err: any) {
      setErrorMessage(err.message || 'Falha ao processar o manifesto');
    } finally {
      setIsLoading(false);
    }
  }

  // Drag and Drop handlers
  function handleDragOver(e: React.DragEvent) {
    e.preventDefault();
    setIsDragging(true);
  }

  function handleDragLeave(e: React.DragEvent) {
    e.preventDefault();
    setIsDragging(false);
  }

  function handleDrop(e: React.DragEvent) {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      const droppedFile = files[0];
      setManifestFile(droppedFile);
      processManifestFile(droppedFile, false);
    }
  }

  // Edit store entry inline in Step 2
  function handleEditEntry(batchId: string, index: number, field: 'code' | 'volumes', rawValue: string) {
    setManifestBatches((prev) =>
      prev.map((batch) => {
        if (batch.id !== batchId) return batch;
        const copy = [...batch.entries];
        const item = { ...copy[index] };
        if (field === 'code') {
          item.code = rawValue.replace(/\D/g, '').slice(0, 5);
        } else {
          item.volumes = Math.max(0, parseInt(rawValue || '0', 10) || 0);
        }
        copy[index] = item;
        return { ...batch, entries: copy };
      })
    );
  }

  // Remove single store entry
  function handleRemoveEntry(batchId: string, index: number) {
    setManifestBatches((prev) =>
      prev.map((batch) =>
        batch.id === batchId ? { ...batch, entries: batch.entries.filter((_, i) => i !== index) } : batch
      )
    );
  }

  // Remove whole batch
  function handleRemoveBatch(batchId: string) {
    setManifestBatches((prev) => {
      const remaining = prev.filter((b) => b.id !== batchId);
      if (!remaining.length) setStep(1);
      return remaining;
    });
  }

  // Save to Google Sheets (after modal confirmation)
  async function handleExecuteGoogleSheetsUpdate() {
    if (!sheetData || !sheetData.spreadsheetId || !allEntries) return;

    let token = googleAccessToken;
    if (!token) {
      token = await getAccessToken();
    }
    if (!token) {
      try {
        const authRes = await googleSignIn();
        if (authRes) {
          token = authRes.accessToken;
          setGoogleUser(authRes.user);
          setGoogleAccessToken(authRes.accessToken);
        }
      } catch (err: any) {
        setErrorMessage('Sessão expirada. Por favor, conecte-se com sua conta Google novamente.');
        setShowConfirmModal(false);
        return;
      }
    }

    if (!token) {
      setErrorMessage('Token de acesso do Google Sheets não encontrado. Faça login com o Google para autorizar.');
      setShowConfirmModal(false);
      return;
    }

    try {
      setIsSavingGoogleSheets(true);
      setErrorMessage('');

      const volumeMap = new Map<string, number>(
        matchedStores.map((store) => [normalizeCode(store.code), store.volumes])
      );

      const result = await batchUpdateGoogleSheet(
        sheetData.spreadsheetId,
        sheetData.sheetName,
        sheetData,
        volumeMap,
        selectedDate,
        orderNumber,
        token
      );

      setFilledStoreKeys([...volumeMap.keys()]);

      addHistory({
        id: crypto.randomUUID(),
        date: selectedDate,
        manifest: manifestBatches.map((b) => b.name).join(' + '),
        stores: allEntries.length,
        volumes: totalVolumes,
        filled: volumeMap.size,
        missing: unmatchedStores.length,
        status: unmatchedStores.length ? 'Com pendências' : 'Concluído',
        targetType: 'google_sheets',
        spreadsheetTitle: sheetData.spreadsheetTitle || sheetData.fileName,
        entries: allEntries.map((e) => ({
          code: e.code,
          volumes: e.volumes,
          occurrences: e.occurrences,
        })),
      });

      setShowConfirmModal(false);
      setStatusMessage(
        `Planilha online atualizada com sucesso! (${result.updatedCells} células gravadas)`
      );
      setStep(3);
    } catch (err: any) {
      setErrorMessage(`Erro ao gravar no Google Sheets: ${err.message}`);
    } finally {
      setIsSavingGoogleSheets(false);
    }
  }

  useEffect(() => {
    if (activeTab === 'historico') {
      setHistoryList(getHistory());
    }
  }, [activeTab]);

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col font-sans">
      {/* Header */}
      <header className="text-primary-foreground shadow-md" style={{ background: 'var(--gradient-header)' }}>
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-4 py-5 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-white/20 backdrop-blur-sm shadow-inner">
              <Globe className="h-6 w-6 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-bold uppercase tracking-wide sm:text-2xl text-white">
                  Conferência de Volumes
                </h1>
                <span className="hidden sm:inline-flex items-center gap-1 rounded-full bg-white/20 px-2 py-0.5 text-[11px] font-semibold text-white">
                  Google Sheets
                </span>
              </div>
              <p className="text-xs sm:text-sm text-white/85">
                Preenchimento automático na sua planilha online via Google Sheets
              </p>
            </div>
          </div>

          {/* Account status in Header & Navigation Tabs */}
          <div className="flex flex-wrap items-center gap-2 self-start sm:self-auto">
            {googleUser ? (
              <div className="flex items-center gap-2 rounded-full bg-white/15 px-3 py-1 text-xs text-white backdrop-blur-xs border border-white/20">
                {googleUser.photoURL ? (
                  <img
                    src={googleUser.photoURL}
                    alt={googleUser.displayName || 'Google'}
                    className="h-5 w-5 rounded-full object-cover border border-white/40"
                  />
                ) : (
                  <div className="flex h-5 w-5 items-center justify-center rounded-full bg-white text-blue-700 font-bold text-[10px]">
                    {(googleUser.displayName || googleUser.email || 'G')[0].toUpperCase()}
                  </div>
                )}
                <span className="font-medium max-w-[150px] truncate" title={googleUser.email || ''}>
                  {googleUser.email}
                </span>
                <button
                  type="button"
                  onClick={handleSwitchAccount}
                  className="ml-1 text-white/90 hover:text-white hover:bg-white/15 px-2 py-0.5 rounded text-[11px] font-medium transition-colors flex items-center gap-1 cursor-pointer"
                  title="Trocar para outro e-mail do Google (ex: pocop703@gmail.com)"
                >
                  <Users className="h-3 w-3" />
                  <span className="hidden sm:inline">Trocar conta</span>
                </button>
                <button
                  type="button"
                  onClick={handleGoogleLogout}
                  className="text-white/70 hover:text-white p-0.5 rounded transition-colors cursor-pointer"
                  title="Sair da conta Google"
                >
                  <LogOut className="h-3.5 w-3.5" />
                </button>
              </div>
            ) : (
              <button
                type="button"
                disabled={isLoggingIn}
                onClick={handleGoogleLogin}
                className="inline-flex items-center gap-1.5 rounded-full bg-white px-3.5 py-1 text-xs font-semibold text-blue-700 shadow-sm hover:bg-white/90 transition-all cursor-pointer disabled:opacity-50"
              >
                {isLoggingIn ? (
                  <Loader2 className="h-3 w-3 animate-spin" />
                ) : (
                  <LogIn className="h-3.5 w-3.5" />
                )}
                Login Google (Qualquer conta)
              </button>
            )}

            {/* Navigation Tabs */}
            <nav className="flex gap-1 rounded-lg bg-black/25 p-1 backdrop-blur-sm">
              <button
                onClick={() => setActiveTab('processar')}
                className={`flex items-center gap-1.5 rounded-md px-3 py-1 text-xs font-medium transition-all ${
                  activeTab === 'processar'
                    ? 'bg-card text-foreground shadow-sm'
                    : 'text-white/85 hover:bg-white/10 hover:text-white'
                }`}
              >
                <Upload className="h-3.5 w-3.5" />
                Processar
              </button>
              <button
                onClick={() => {
                  setActiveTab('historico');
                  setHistoryList(getHistory());
                }}
                className={`flex items-center gap-1.5 rounded-md px-3 py-1 text-xs font-medium transition-all ${
                  activeTab === 'historico'
                    ? 'bg-card text-foreground shadow-sm'
                    : 'text-white/85 hover:bg-white/10 hover:text-white'
                }`}
              >
                <History className="h-3.5 w-3.5" />
                Histórico
              </button>
              <button
                onClick={() => setActiveTab('planilha')}
                className={`flex items-center gap-1.5 rounded-md px-3 py-1 text-xs font-medium transition-all ${
                  activeTab === 'planilha'
                    ? 'bg-card text-foreground shadow-sm'
                    : 'text-white/85 hover:bg-white/10 hover:text-white'
                }`}
              >
                <Route className="h-3.5 w-3.5" />
                Rotas e lojas
              </button>
            </nav>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="mx-auto max-w-6xl w-full flex-1 px-4 py-6">
        {/* Error Alert */}
        {errorMessage && (
          <div
            className={`mb-6 rounded-xl border p-4 shadow-sm transition-all ${
              errorMessage.includes('unauthorized-domain') || unauthorizedDomain
                ? 'border-amber-500/40 bg-amber-500/10 text-amber-950 dark:text-amber-100'
                : 'border-destructive/30 bg-destructive/10 text-destructive'
            }`}
          >
            <div className="flex items-start gap-3">
              {errorMessage.includes('unauthorized-domain') || unauthorizedDomain ? (
                <ShieldAlert className="h-5 w-5 shrink-0 mt-0.5 text-amber-600 dark:text-amber-400" />
              ) : (
                <AlertTriangle className="h-5 w-5 shrink-0 mt-0.5" />
              )}
              <div className="flex-1 space-y-2">
                <div>
                  <p className="font-semibold text-sm">
                    {errorMessage.includes('unauthorized-domain') || unauthorizedDomain
                      ? 'Domínio não autorizado no Firebase (Deploy no Vercel)'
                      : 'Atenção'}
                  </p>
                  <p className="mt-0.5 text-xs opacity-95 leading-relaxed">{errorMessage}</p>
                </div>

                {(errorMessage.includes('unauthorized-domain') || unauthorizedDomain) && (
                  <div className="flex flex-wrap items-center gap-2 pt-1">
                    <button
                      type="button"
                      onClick={() => setShowVercelHelpModal(true)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-600 hover:bg-amber-700 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
                    >
                      <span>Como liberar este domínio no Firebase</span>
                      <ArrowRight className="h-3.5 w-3.5" />
                    </button>
                  </div>
                )}

                {(errorMessage.includes('AUTH_401') || errorMessage.includes('não autenticada') || errorMessage.includes('Não autenticado') || errorMessage.includes('expirou')) && (
                  <div className="flex flex-wrap items-center gap-2 pt-1">
                    <button
                      type="button"
                      disabled={isLoggingIn}
                      onClick={handleGoogleLogin}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
                    >
                      {isLoggingIn ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <LogIn className="h-3.5 w-3.5" />}
                      <span>Fazer Login / Escolher Conta Google</span>
                    </button>
                    <button
                      type="button"
                      disabled={isLoggingIn}
                      onClick={handleSwitchAccount}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-border bg-card hover:bg-muted text-xs font-medium transition-colors cursor-pointer text-foreground"
                    >
                      <Users className="h-3.5 w-3.5 text-primary" />
                      <span>Entrar com outra conta (ex: pocop703@gmail.com)</span>
                    </button>
                  </div>
                )}

                {(errorMessage.includes('AUTH_403') || errorMessage.includes('permissão') || errorMessage.includes('Permissão negada')) && (
                  <div className="flex flex-wrap items-center gap-2 pt-1">
                    <button
                      type="button"
                      disabled={isLoggingIn}
                      onClick={handleSwitchAccount}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
                    >
                      <Users className="h-3.5 w-3.5" />
                      <span>Trocar de Conta Google (ex: pocop703@gmail.com)</span>
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Tab 1: PROCESSAR */}
        {activeTab === 'processar' && (
          <div className="space-y-6">
            {/* Section: Planilha de Destino no Google Sheets */}
            <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
              <div className="pb-3 border-b border-border/60">
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-base font-semibold flex items-center gap-2">
                      <Globe className="h-4 w-4 text-blue-600" />
                      <span>Planilha Online no Google Sheets</span>
                    </h2>
                    <p className="text-xs text-muted-foreground mt-0.5">
                      Aceita qualquer conta Google (@gmail.com ou institucional). O acesso é concedido via credencial do e-mail logado.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowVercelHelpModal(true)}
                    className="inline-flex items-center gap-1 text-[11px] text-muted-foreground hover:text-foreground px-2 py-1 rounded hover:bg-muted transition-colors cursor-pointer"
                    title="Instruções para deploy no Vercel ou domínios personalizados"
                  >
                    <ShieldAlert className="h-3.5 w-3.5 text-amber-500" />
                    <span>Ajuda Vercel</span>
                  </button>
                </div>
              </div>

              {/* Login State & Account Information */}
              <div className="mt-4 space-y-4">
                {googleUser ? (
                  /* Logged-In User Card */
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-3.5 rounded-xl bg-blue-50/50 dark:bg-blue-950/20 border border-blue-200/80 dark:border-blue-900/50">
                    <div className="flex items-center gap-3">
                      {googleUser.photoURL ? (
                        <img
                          src={googleUser.photoURL}
                          alt={googleUser.displayName || 'Google'}
                          className="h-10 w-10 rounded-full object-cover border border-blue-300 dark:border-blue-700 shadow-xs"
                        />
                      ) : (
                        <div className="flex h-10 w-10 items-center justify-center rounded-full bg-blue-600 text-white font-bold text-sm shadow-xs">
                          {(googleUser.displayName || googleUser.email || 'G')[0].toUpperCase()}
                        </div>
                      )}
                      <div>
                        <div className="flex flex-wrap items-center gap-2">
                          <p className="text-sm font-bold text-foreground">
                            {googleUser.displayName || 'Usuário Google'}
                          </p>
                          <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/15 px-2 py-0.5 text-[10px] font-semibold text-emerald-700 dark:text-emerald-300">
                            <CheckCircle2 className="h-3 w-3" />
                            Acesso Google Sheets Ativo
                          </span>
                        </div>
                        <p className="text-xs text-muted-foreground flex flex-wrap items-center gap-1.5 mt-0.5">
                          <Mail className="h-3 w-3 text-blue-600 shrink-0" />
                          <span className="font-mono font-medium text-foreground">{googleUser.email}</span>
                          <span className="text-[11px] text-muted-foreground">
                            (Qualquer conta Google é aceita)
                          </span>
                        </p>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-2 self-start sm:self-auto">
                      <button
                        type="button"
                        disabled={isLoggingIn}
                        onClick={handleSwitchAccount}
                        className="inline-flex items-center gap-1.5 text-xs text-primary bg-primary/10 hover:bg-primary/20 px-3 py-1.5 rounded-md font-semibold transition-colors cursor-pointer"
                        title="Escolha outro e-mail ou faça login com outra conta (ex: pocop703@gmail.com)"
                      >
                        {isLoggingIn ? (
                          <Loader2 className="h-3.5 w-3.5 animate-spin" />
                        ) : (
                          <Users className="h-3.5 w-3.5" />
                        )}
                        Trocar conta Google
                      </button>
                      <button
                        type="button"
                        onClick={handleGoogleLogout}
                        className="inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-destructive px-2.5 py-1.5 rounded-md hover:bg-card border border-border/60 transition-colors cursor-pointer"
                        title="Sair da conta Google"
                      >
                        <LogOut className="h-3.5 w-3.5" />
                        Sair
                      </button>
                    </div>
                  </div>
                ) : (
                  /* Not Logged In Prompt */
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 p-4 rounded-xl border border-blue-500/30 bg-blue-500/5">
                    <div className="flex items-start gap-3">
                      <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-blue-100 dark:bg-blue-950 text-blue-700 dark:text-blue-300 shrink-0">
                        <UserCheck className="h-5 w-5" />
                      </div>
                      <div>
                        <div className="flex flex-wrap items-center gap-2">
                          <p className="text-sm font-bold text-foreground">
                            Faça login com sua conta Google para continuar
                          </p>
                          <span className="text-[10px] uppercase font-bold tracking-wider px-2 py-0.5 rounded bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300">
                            Qualquer e-mail aceito
                          </span>
                        </div>
                        <p className="text-xs text-muted-foreground mt-0.5 max-w-xl leading-relaxed">
                          O aplicativo aceita <strong>qualquer conta Google</strong> (@gmail.com ou institucional).
                          Ao clicar no botão, a janela do Google exibirá a lista de contas salvas para você selecionar (ex: <span className="font-mono text-primary font-semibold">pocop703@gmail.com</span>) ou digitar qualquer outra conta.
                        </p>
                      </div>
                    </div>

                    <button
                      type="button"
                      disabled={isLoggingIn}
                      onClick={handleGoogleLogin}
                      className="inline-flex items-center justify-center gap-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 text-xs font-bold shadow transition-colors disabled:opacity-50 cursor-pointer shrink-0"
                    >
                      {isLoggingIn ? (
                        <Loader2 className="h-4 w-4 animate-spin text-white" />
                      ) : (
                        <svg className="h-4 w-4" viewBox="0 0 24 24">
                          <path
                            fill="#FFFFFF"
                            d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
                          />
                          <path
                            fill="#FFFFFF"
                            d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
                          />
                          <path
                            fill="#FFFFFF"
                            d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z"
                          />
                          <path
                            fill="#FFFFFF"
                            d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z"
                          />
                        </svg>
                      )}
                      Fazer login com Google
                    </button>
                  </div>
                )}

                {/* Connected Status Card with Quick Change button */}
                {sheetData && !isEditingSheetUrl ? (
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 p-3.5 rounded-lg border border-emerald-500/30 bg-emerald-500/5">
                    <div className="flex items-center gap-2.5">
                      <div className="flex h-8 w-8 items-center justify-center rounded-full bg-emerald-100 text-emerald-700 shrink-0">
                        <Check className="h-4 w-4" />
                      </div>
                      <div>
                        <p className="text-xs font-bold text-foreground flex items-center gap-1.5">
                          <span>{sheetData.spreadsheetTitle}</span>
                          <span className="text-[11px] font-normal px-2 py-0.5 rounded bg-muted text-muted-foreground">
                            Aba: {sheetData.sheetName}
                          </span>
                        </p>
                        <p className="text-[11px] text-muted-foreground mt-0.5">
                          {sheetData.routes.length} rotas e {sheetData.routes.reduce((a, b) => a + b.stores.length, 0)} lojas prontas para receber volumes
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2">
                      {sheetData.spreadsheetUrl && (
                        <a
                          href={sheetData.spreadsheetUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-flex items-center gap-1 text-xs text-primary hover:underline px-2.5 py-1 rounded bg-card border border-border"
                        >
                          Abrir planilha
                          <ExternalLink className="h-3 w-3" />
                        </a>
                      )}
                      <button
                        type="button"
                        onClick={() => setIsEditingSheetUrl(true)}
                        className="inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground px-2.5 py-1 rounded border border-input bg-card cursor-pointer"
                      >
                        <Settings2 className="h-3 w-3" />
                        Alterar planilha
                      </button>
                    </div>
                  </div>
                ) : (
                  /* Spreadsheet URL Configuration */
                  <div className="p-4 rounded-lg border border-border bg-card space-y-3">
                    <div className="grid gap-3 sm:grid-cols-12 items-end">
                      <div className="sm:col-span-8 flex flex-col">
                        <label className="text-xs uppercase tracking-wide text-muted-foreground font-semibold mb-1">
                          LINK OU ID DA PLANILHA NO GOOGLE SHEETS
                        </label>
                        <input
                          type="text"
                          value={googleSheetsUrl}
                          onChange={(e) => setGoogleSheetsUrl(e.target.value)}
                          placeholder="https://docs.google.com/spreadsheets/d/.../edit"
                          className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-xs shadow-sm focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring font-mono"
                        />
                      </div>

                      {spreadsheetMeta && spreadsheetMeta.sheets.length > 1 && (
                        <div className="sm:col-span-4 flex flex-col">
                          <label className="text-xs uppercase tracking-wide text-muted-foreground font-semibold mb-1">
                            ABA / PÁGINA
                          </label>
                          <select
                            value={selectedTabName}
                            onChange={(e) => setSelectedTabName(e.target.value)}
                            className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-xs shadow-sm focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                          >
                            {spreadsheetMeta.sheets.map((s) => (
                              <option key={s.sheetId} value={s.title}>
                                {s.title}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}

                      <div className={spreadsheetMeta && spreadsheetMeta.sheets.length > 1 ? 'sm:col-span-12' : 'sm:col-span-4'}>
                        <button
                          type="button"
                          disabled={!googleSheetsUrl.trim() || isLoading}
                          onClick={handleConnectGoogleSheet}
                          className="w-full h-9 inline-flex items-center justify-center gap-2 rounded-md bg-blue-600 px-3 py-1 text-xs font-semibold text-white shadow transition-colors hover:bg-blue-700 disabled:opacity-50 cursor-pointer"
                        >
                          {isLoading ? (
                            <Loader2 className="h-3.5 w-3.5 animate-spin" />
                          ) : (
                            <Globe className="h-3.5 w-3.5" />
                          )}
                          Conectar e Ler Planilha
                        </button>
                      </div>
                    </div>

                    <div className="flex items-center justify-between text-xs text-muted-foreground pt-1">
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={saveAsDefault}
                          onChange={(e) => setSaveAsDefault(e.target.checked)}
                          className="rounded border-input text-primary focus:ring-primary h-3.5 w-3.5"
                        />
                        <span>Salvar esta planilha como padrão neste navegador (não precisará colar novamente)</span>
                      </label>
                      {sheetData && (
                        <button
                          type="button"
                          onClick={() => setIsEditingSheetUrl(false)}
                          className="text-xs text-primary hover:underline cursor-pointer"
                        >
                          Cancelar
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </section>

            {/* Section: Upload & Processamento do Manifesto */}
            <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)] transition-all">
              <div className="grid gap-4 sm:grid-cols-4 pb-4 border-b border-border/60">
                {/* Data */}
                <label className="text-sm font-medium flex flex-col">
                  <span className="text-xs uppercase tracking-wide text-muted-foreground font-semibold mb-1">
                    DATA DA CONFERÊNCIA
                  </span>
                  <input
                    type="date"
                    value={selectedDate}
                    onChange={(e) => {
                      setSelectedDate(e.target.value);
                      handleReset();
                    }}
                    className="flex h-9 w-full rounded-md border border-input bg-transparent px-3 py-1 text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                  />
                </label>

                {/* Número do Pedido */}
                <div className="text-sm font-medium flex flex-col">
                  <span className="text-xs uppercase tracking-wide text-muted-foreground font-semibold mb-1">
                    NÚMERO DO PEDIDO (JULIANO)
                  </span>
                  <div className="flex h-9 items-center rounded-md border border-border bg-muted/60 px-3 font-semibold tabular-nums text-primary shadow-inner">
                    {orderNumber || '—'}
                  </div>
                </div>

                {/* Manifesto Upload Direct Input */}
                <label className="text-sm font-medium flex flex-col sm:col-span-2">
                  <span className="text-xs uppercase tracking-wide text-muted-foreground font-semibold mb-1 flex items-center justify-between">
                    <span>SELECIONAR ARQUIVO DE MANIFESTO</span>
                    {manifestBatches.length > 0 && (
                      <span className="text-[11px] text-success font-medium">
                        {manifestBatches.length} carregado(s)
                      </span>
                    )}
                  </span>
                  <input
                    ref={manifestInputRef}
                    type="file"
                    accept=".pdf,image/*"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) {
                        setManifestFile(f);
                        processManifestFile(f, false);
                      }
                    }}
                    className="flex h-9 w-full rounded-md border border-input bg-transparent px-3 py-1 text-xs shadow-sm transition-colors file:border-0 file:bg-transparent file:text-xs file:font-semibold file:text-primary placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                  />
                </label>
              </div>

              {/* Drag and Drop Zone */}
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                className={`mt-4 flex flex-col items-center justify-center rounded-xl border-2 border-dashed p-6 text-center transition-all ${
                  isDragging
                    ? 'border-primary bg-primary/10 scale-[1.01]'
                    : 'border-border bg-muted/20 hover:border-primary/50'
                }`}
              >
                <div className="flex h-12 w-12 items-center justify-center rounded-full bg-primary/10 text-primary mb-2">
                  <Upload className="h-6 w-6" />
                </div>
                <p className="text-sm font-semibold">
                  Arraste e solte o manifesto aqui (PDF ou imagem)
                </p>
                <p className="mt-1 text-xs text-muted-foreground">
                  Suporta manifestos em PDF texto, PDFs digitalizados ou imagens (OCR)
                </p>

                <div className="mt-3 flex items-center gap-2">
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => manifestInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 rounded-md bg-secondary px-3 py-1 text-xs font-medium text-secondary-foreground shadow-xs hover:bg-secondary/80 cursor-pointer"
                  >
                    Procurar arquivo...
                  </button>
                  <span className="text-xs text-muted-foreground">ou</span>
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => processManifestFile(null, true)}
                    className="inline-flex items-center gap-1.5 rounded-md border border-input bg-card px-3 py-1 text-xs font-medium text-muted-foreground hover:bg-accent hover:text-foreground cursor-pointer"
                  >
                    <Sparkles className="h-3.5 w-3.5 text-amber-500" />
                    Carregar manifesto demo
                  </button>
                </div>
              </div>

              {/* Status feedback banner */}
              {statusMessage && (
                <div className="mt-4 flex items-center gap-2 rounded-lg bg-primary/10 p-3 text-xs font-medium text-primary">
                  {isLoading && <Loader2 className="h-4 w-4 animate-spin shrink-0" />}
                  <span>{statusMessage}</span>
                </div>
              )}
            </section>

            {/* Step 2: Conferência das lojas */}
            {step >= 2 && allEntries && (
              <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 pb-3 border-b border-border">
                  <div>
                    <h2 className="text-lg font-semibold flex items-center gap-2">
                      <CheckCircle2 className="h-5 w-5 text-primary" />
                      Lojas encontradas no manifesto ({allEntries.length})
                    </h2>
                    <p className="text-xs text-muted-foreground mt-0.5">
                      Total de volumes: <strong className="text-primary">{totalVolumes}</strong> · Verifique os dados antes de gravar no Google Sheets.
                    </p>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => manifestInputRef.current?.click()}
                      className="inline-flex items-center gap-1.5 text-xs font-medium text-primary hover:underline"
                    >
                      + Adicionar outro manifesto
                    </button>
                  </div>
                </div>

                {/* Batches badges */}
                <div className="mt-3 flex flex-wrap gap-2">
                  {manifestBatches.map((batch) => (
                    <span
                      key={batch.id}
                      className="inline-flex items-center gap-1.5 rounded-md bg-muted px-2.5 py-1 text-xs font-medium text-foreground border border-border"
                    >
                      <span>{batch.name}</span>
                      <span className="text-muted-foreground">({batch.entries.length} lojas)</span>
                      {manifestBatches.length > 1 && (
                        <button
                          type="button"
                          onClick={() => handleRemoveBatch(batch.id)}
                          className="hover:text-destructive text-muted-foreground"
                          title="Remover este lote"
                        >
                          ✕
                        </button>
                      )}
                    </span>
                  ))}
                </div>

                {/* Stores table */}
                <div className="mt-4 max-h-96 overflow-y-auto rounded-lg border border-border">
                  <table className="w-full text-xs">
                    <thead className="bg-muted text-muted-foreground sticky top-0 uppercase tracking-wider">
                      <tr>
                        <th className="px-3 py-2 text-left">Código / Loja</th>
                        <th className="px-3 py-2 text-left">Rota mapeada</th>
                        <th className="px-3 py-2 text-right">Volumes</th>
                        <th className="px-3 py-2 text-center">Status</th>
                        <th className="px-3 py-2 text-right">Ação</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-border">
                      {manifestBatches.map((batch) =>
                        batch.entries.map((entry, idx) => {
                          const normalized = normalizeCode(entry.code);
                          const storeInfo = sheetStoreMap.get(normalized);
                          const isMatched = Boolean(storeInfo);

                          return (
                            <tr key={`${batch.id}-${idx}`} className="hover:bg-muted/40">
                              <td className="px-3 py-2">
                                <input
                                  type="text"
                                  value={entry.code}
                                  onChange={(e) => handleEditEntry(batch.id, idx, 'code', e.target.value)}
                                  className="w-20 rounded border border-input bg-card px-1.5 py-0.5 text-xs font-bold text-foreground"
                                />
                                {storeInfo && (
                                  <span className="ml-2 text-muted-foreground">
                                    Loja {storeInfo.display}
                                  </span>
                                )}
                              </td>
                              <td className="px-3 py-2 text-muted-foreground">
                                {storeInfo ? (
                                  <span className="font-medium text-foreground">{storeInfo.route}</span>
                                ) : (
                                  <span className="text-warning-foreground italic">Não encontrada</span>
                                )}
                              </td>
                              <td className="px-3 py-2 text-right">
                                <input
                                  type="number"
                                  min={0}
                                  value={entry.volumes}
                                  onChange={(e) => handleEditEntry(batch.id, idx, 'volumes', e.target.value)}
                                  className="w-16 rounded border border-input bg-card px-1.5 py-0.5 text-right text-xs font-bold text-primary"
                                />
                              </td>
                              <td className="px-3 py-2 text-center">
                                {isMatched ? (
                                  <span className="inline-flex rounded-full bg-success/15 px-2 py-0.5 text-[10px] font-semibold text-success">
                                    Pronta
                                  </span>
                                ) : (
                                  <span className="inline-flex rounded-full bg-warning/20 px-2 py-0.5 text-[10px] font-semibold text-warning-foreground">
                                    Fora da planilha
                                  </span>
                                )}
                              </td>
                              <td className="px-3 py-2 text-right">
                                <button
                                  type="button"
                                  onClick={() => handleRemoveEntry(batch.id, idx)}
                                  className="text-[11px] text-muted-foreground hover:text-destructive"
                                >
                                  Remover
                                </button>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Confirm step actions */}
                <div className="mt-5 flex flex-wrap items-center gap-3 pt-3 border-t border-border/60">
                  <button
                    onClick={() => setShowConfirmModal(true)}
                    disabled={!matchedStores.length || isLoading || !sheetData}
                    className="inline-flex items-center justify-center gap-2 rounded-md bg-blue-600 px-5 py-2.5 text-sm font-semibold text-white shadow transition-colors hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring disabled:opacity-50 disabled:pointer-events-none cursor-pointer"
                  >
                    <Globe className="h-4 w-4" />
                    Preencher no Google Sheets
                    <ArrowRight className="h-4 w-4 ml-1" />
                  </button>

                  <button
                    onClick={() => setStep(1)}
                    className="inline-flex items-center justify-center rounded-md border border-input bg-background px-4 py-2.5 text-sm font-medium shadow-sm transition-colors hover:bg-accent hover:text-accent-foreground cursor-pointer"
                  >
                    Voltar
                  </button>

                  {!sheetData && (
                    <span className="text-sm text-amber-600 dark:text-amber-400 flex items-center gap-1.5">
                      <AlertTriangle className="h-4 w-4 text-warning" />
                      Conecte sua planilha do Google Sheets acima para preencher.
                    </span>
                  )}
                </div>
              </section>
            )}

            {/* Step 3: Resultado */}
            {step === 3 && sheetData && allEntries && (
              <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)] animate-fadeIn">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                  <div>
                    <h2 className="text-lg font-semibold flex items-center gap-2">
                      <CheckCircle2 className="h-5 w-5 text-success" />
                      <span>Resultado da Conferência no Google Sheets</span>
                    </h2>
                    <p className="mt-0.5 text-sm text-muted-foreground">
                      Os volumes foram gravados com sucesso na sua planilha online!
                    </p>
                  </div>

                  {sheetData.spreadsheetUrl && (
                    <a
                      href={sheetData.spreadsheetUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-2 px-4 py-2 rounded-md bg-blue-600 text-white text-sm font-semibold shadow hover:bg-blue-700 transition-colors"
                    >
                      <ExternalLink className="h-4 w-4" />
                      Abrir no Google Sheets
                    </a>
                  )}
                </div>

                {/* 4 Stat Cards */}
                <div className="mt-4 grid gap-3 grid-cols-2 sm:grid-cols-4">
                  <StatCard label="Lojas preenchidas" value={filledStoreKeys.length} tone="success" />
                  <StatCard label="Volumes processados" value={totalVolumes} tone="primary" />
                  <StatCard label="Não encontradas" value={unmatchedStores.length} tone="warning" />
                  <StatCard label="Sem volume no manifesto" value={unmanifestedStores.length} tone="warning" />
                </div>

                {/* Action buttons */}
                <div className="mt-4 flex flex-wrap gap-3">
                  {sheetData.spreadsheetUrl && (
                    <a
                      href={sheetData.spreadsheetUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center justify-center gap-2 rounded-md bg-blue-600 px-5 py-2.5 text-sm font-semibold text-white shadow transition-colors hover:bg-blue-700 cursor-pointer"
                    >
                      <Globe className="h-4 w-4" />
                      Acessar planilha no Google Sheets
                      <ExternalLink className="h-3.5 w-3.5 ml-0.5" />
                    </a>
                  )}

                  <button
                    onClick={() => setStep(2)}
                    className="inline-flex items-center justify-center rounded-md border border-input bg-background px-4 py-2.5 text-sm font-medium shadow-sm transition-colors hover:bg-accent hover:text-accent-foreground cursor-pointer"
                  >
                    Voltar e corrigir
                  </button>

                  <button
                    onClick={handleReset}
                    className="inline-flex items-center justify-center gap-1.5 rounded-md border border-input bg-card px-4 py-2.5 text-sm font-medium text-muted-foreground shadow-sm hover:text-foreground cursor-pointer ml-auto"
                  >
                    <RotateCcw className="h-4 w-4" />
                    Iniciar nova conferência
                  </button>
                </div>

                {/* Lojas não encontradas warning list */}
                {unmatchedStores.length > 0 && (
                  <div className="mt-6 rounded-lg border border-warning/30 bg-warning/5 p-4">
                    <h3 className="text-sm font-semibold uppercase tracking-wide text-warning-foreground flex items-center gap-1.5">
                      <AlertTriangle className="h-4 w-4 text-warning" />
                      Lojas não encontradas na planilha ({unmatchedStores.length})
                    </h3>
                    <ul className="mt-2.5 flex flex-wrap gap-2 text-sm">
                      {unmatchedStores.map((store) => (
                        <li
                          key={store.code}
                          className="rounded-md bg-warning/20 px-2.5 py-1 font-medium text-warning-foreground text-xs"
                        >
                          Loja {store.code} · {store.volumes} volumes
                        </li>
                      ))}
                    </ul>
                  </div>
                )}

                {/* Lojas da planilha sem volume list */}
                {unmanifestedStores.length > 0 && (
                  <div className="mt-4 rounded-lg border border-border bg-muted/20 p-4">
                    <h3 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground flex items-center gap-1.5">
                      <Info className="h-4 w-4" />
                      Lojas da planilha sem volume no manifesto ({unmanifestedStores.length})
                    </h3>
                    <ul className="mt-2.5 flex flex-wrap gap-2 text-xs">
                      {unmanifestedStores.map((store) => (
                        <li
                          key={`${store.route}-${store.display}`}
                          className="rounded-md bg-muted px-2.5 py-1 text-muted-foreground border border-border/50"
                        >
                          Loja {store.display} · <span className="opacity-75">{store.route}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}

                {/* Preview of filled workbook */}
                <div className="mt-6 pt-5 border-t border-border">
                  <h3 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground mb-3 flex items-center gap-2">
                    <FileSpreadsheet className="h-4 w-4 text-primary" />
                    Prévia da conferência por rota
                  </h3>
                  <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
                    {sheetData.routes.map((route) => {
                      const highlightSet = new Set(filledStoreKeys);
                      const routeTotal = route.stores.reduce((a, b) => a + (b.currentVol ?? 0), 0);
                      return (
                        <div key={route.name} className="rounded-lg border border-border bg-card overflow-hidden shadow-sm">
                          <p className="border-b border-border bg-muted/60 px-3 py-2 text-xs font-bold uppercase tracking-wider text-primary">
                            {route.name}
                          </p>
                          <table className="w-full text-xs">
                            <tbody className="divide-y divide-border/60">
                              {route.stores.map((store) => {
                                const isHighlighted = highlightSet.has(store.key);
                                return (
                                  <tr
                                    key={store.display}
                                    className={isHighlighted ? 'bg-success/10 font-medium' : ''}
                                  >
                                    <td className="px-3 py-1.5 tabular-nums">Loja {store.display}</td>
                                    <td className="px-3 py-1.5 text-right tabular-nums">
                                      {store.currentVol !== null ? (
                                        <span className={isHighlighted ? 'text-success font-bold' : ''}>
                                          {store.currentVol}
                                        </span>
                                      ) : (
                                        <span className="text-muted-foreground">—</span>
                                      )}
                                    </td>
                                  </tr>
                                );
                              })}
                              <tr className="border-t-2 border-border font-bold bg-muted/30">
                                <td className="px-3 py-2 text-foreground">TOTAL</td>
                                <td className="px-3 py-2 text-right tabular-nums text-primary">
                                  {routeTotal}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </section>
            )}
          </div>
        )}

        {/* Tab 2: HISTÓRICO */}
        {activeTab === 'historico' && (
          <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="flex items-center justify-between pb-3 border-b border-border">
              <div>
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  <History className="h-5 w-5 text-primary" />
                  Histórico de processamentos
                </h2>
                <p className="text-xs text-muted-foreground mt-0.5">
                  Conferências enviadas para o Google Sheets.
                </p>
              </div>
              {historyList.length > 0 && (
                <button
                  onClick={() => {
                    clearHistory();
                    setHistoryList([]);
                  }}
                  className="inline-flex items-center gap-1.5 text-xs text-destructive hover:underline font-medium cursor-pointer"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  Limpar histórico
                </button>
              )}
            </div>

            {historyList.length === 0 ? (
              <div className="py-12 text-center text-muted-foreground text-sm">
                <FileSpreadsheet className="h-10 w-10 mx-auto text-muted-foreground/40 mb-2" />
                <p>Nenhum processamento salvo ainda.</p>
                <p className="text-xs text-muted-foreground/75 mt-1">
                  Processe e confirme uma conferência na aba Processar para registrar o histórico.
                </p>
              </div>
            ) : (
              <div className="mt-4 overflow-x-auto rounded-lg border border-border">
                <table className="w-full text-sm">
                  <thead className="bg-muted text-left text-xs uppercase text-muted-foreground">
                    <tr>
                      <th className="px-3 py-2.5">Data</th>
                      <th className="px-3 py-2.5">Destino</th>
                      <th className="px-3 py-2.5">Manifesto</th>
                      <th className="px-3 py-2.5 text-center">Lojas</th>
                      <th className="px-3 py-2.5 text-center">Volumes</th>
                      <th className="px-3 py-2.5 text-center">Preenchidas</th>
                      <th className="px-3 py-2.5">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {historyList.map((record) => (
                      <tr key={record.id} className="hover:bg-muted/30 transition-colors">
                        <td className="px-3 py-2.5 font-medium tabular-nums">
                          {record.date.split('-').reverse().join('/')}
                        </td>
                        <td className="px-3 py-2.5">
                          <span className="inline-flex items-center gap-1 text-xs">
                            <Globe className="h-3.5 w-3.5 text-blue-600" />
                            <span className="font-medium text-blue-600">Google Sheets</span>
                          </span>
                        </td>
                        <td className="px-3 py-2.5 text-muted-foreground max-w-xs truncate">
                          {record.manifest}
                        </td>
                        <td className="px-3 py-2.5 tabular-nums text-center">{record.stores}</td>
                        <td className="px-3 py-2.5 tabular-nums text-center font-semibold text-primary">
                          {record.volumes}
                        </td>
                        <td className="px-3 py-2.5 tabular-nums text-center text-success font-semibold">
                          {record.filled}
                        </td>
                        <td className="px-3 py-2.5">
                          <span
                            className={`inline-flex rounded-full px-2.5 py-0.5 text-xs font-semibold ${
                              record.status === 'Concluído'
                                ? 'bg-success/15 text-success'
                                : 'bg-warning/20 text-warning-foreground'
                            }`}
                          >
                            {record.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        )}

        {/* Tab 3: ROTAS E LOJAS */}
        {activeTab === 'planilha' && (
          <section className="rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 pb-3 border-b border-border">
              <div>
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  <Route className="h-5 w-5 text-primary" />
                  Rotas e lojas da planilha online
                </h2>
                <p className="text-xs text-muted-foreground mt-0.5">
                  Visualização detalhada das lojas e colunas lidas no Google Sheets.
                </p>
              </div>

              {sheetData && (
                <div className="inline-flex items-center gap-2 text-xs font-medium bg-muted px-3 py-1 rounded-md border border-border">
                  <Globe className="h-3.5 w-3.5 text-blue-600" />
                  <span>{sheetData.spreadsheetTitle}</span>
                </div>
              )}
            </div>

            {sheetData ? (
              <div className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
                {sheetData.routes.map((route) => {
                  const totalVols = route.stores.reduce((acc, s) => acc + (s.currentVol ?? 0), 0);
                  return (
                    <div key={route.name} className="rounded-lg border border-border p-4 bg-muted/10 shadow-sm">
                      <p className="font-bold text-primary text-sm uppercase tracking-wide">{route.name}</p>
                      <p className="text-xs text-muted-foreground mt-0.5">
                        {route.stores.length} lojas · {totalVols} volumes
                      </p>
                      <ul className="mt-3 flex flex-wrap gap-1.5 text-xs">
                        {route.stores.map((store) => (
                          <li
                            key={store.display}
                            className="rounded bg-muted px-2 py-1 tabular-nums border border-border/60 text-foreground flex items-center gap-1"
                          >
                            <span className="font-medium">Loja {store.display}</span>
                            {store.currentVol !== null && (
                              <span className="font-semibold text-primary">· {store.currentVol}</span>
                            )}
                          </li>
                        ))}
                      </ul>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="py-12 text-center text-muted-foreground text-sm">
                <FileSpreadsheet className="h-10 w-10 mx-auto text-muted-foreground/40 mb-2" />
                <p>Nenhuma planilha do Google Sheets conectada no momento.</p>
                <p className="text-xs text-muted-foreground/75 mt-1">
                  Conecte sua planilha na aba <strong>Processar</strong> com sua conta Google para visualizar as rotas.
                </p>
                <button
                  onClick={() => setActiveTab('processar')}
                  className="mt-4 inline-flex items-center gap-1.5 rounded-md bg-blue-600 px-4 py-2 text-xs font-semibold text-white shadow hover:bg-blue-700 cursor-pointer"
                >
                  <Globe className="h-3.5 w-3.5" />
                  Ir para Conectar Google Sheets
                </button>
              </div>
            )}
          </section>
        )}
      </main>

      {/* Confirmation Modal before writing into Google Sheets */}
      {showConfirmModal && sheetData && allEntries && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div className="w-full max-w-md rounded-xl border border-border bg-card p-6 shadow-2xl space-y-4">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-full bg-blue-100 text-blue-700">
                <Globe className="h-5 w-5" />
              </div>
              <div>
                <h3 className="text-base font-bold text-foreground">Confirmar Gravação no Google Sheets</h3>
                <p className="text-xs text-muted-foreground">
                  {sheetData.spreadsheetTitle} • Aba {sheetData.sheetName}
                </p>
              </div>
            </div>

            <div className="rounded-lg bg-muted/40 p-3 space-y-2 text-xs">
              <div className="flex justify-between">
                <span className="text-muted-foreground">E-mail do autor:</span>
                <span className="font-mono font-medium text-foreground">{googleUser?.email || 'Google Account'}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Lojas a atualizar:</span>
                <span className="font-bold text-foreground">{matchedStores.length} lojas</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Total de volumes:</span>
                <span className="font-bold text-primary">{totalVolumes}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Data da conferência:</span>
                <span className="font-medium text-foreground">{selectedDate.split('-').reverse().join('/')}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Nº Pedido (Juliano):</span>
                <span className="font-medium text-foreground">{orderNumber}</span>
              </div>
              {unmatchedStores.length > 0 && (
                <div className="pt-2 border-t border-border flex justify-between text-warning-foreground font-medium">
                  <span>Lojas fora da planilha:</span>
                  <span>{unmatchedStores.length} (serão ignoradas)</span>
                </div>
              )}
            </div>

            <p className="text-xs text-muted-foreground leading-relaxed">
              Os volumes serão gravados diretamente nas células correspondentes de cada loja no Google Sheets.
            </p>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowConfirmModal(false)}
                disabled={isSavingGoogleSheets}
                className="px-4 py-2 text-xs font-medium rounded-md border border-input bg-background hover:bg-accent text-foreground transition-colors cursor-pointer"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={handleExecuteGoogleSheetsUpdate}
                disabled={isSavingGoogleSheets}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-bold rounded-md bg-blue-600 hover:bg-blue-700 text-white shadow transition-colors disabled:opacity-50 cursor-pointer"
              >
                {isSavingGoogleSheets ? (
                  <>
                    <Loader2 className="h-3.5 w-3.5 animate-spin" />
                    <span>Gravando na planilha...</span>
                  </>
                ) : (
                  <>
                    <Check className="h-3.5 w-3.5" />
                    <span>Confirmar e Gravar</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Vercel Domain Authorization Help Modal */}
      <VercelAuthHelpModal
        isOpen={showVercelHelpModal || Boolean(unauthorizedDomain)}
        domain={unauthorizedDomain}
        onClose={() => {
          setShowVercelHelpModal(false);
          setUnauthorizedDomain(null);
        }}
        onRetryLogin={handleGoogleLogin}
        isRetrying={isLoggingIn}
      />

      {/* Footer */}
      <footer className="border-t border-border bg-card/50 py-4 text-center text-xs text-muted-foreground">
        Conferência de Volumes · Preenchimento Automático do Google Sheets
      </footer>
    </div>
  );
}
