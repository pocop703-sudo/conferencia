import React, { useState } from 'react';
import {
  AlertTriangle,
  ExternalLink,
  Copy,
  Check,
  X,
  RefreshCw,
  ShieldAlert,
  FileCode,
  KeyRound,
  CheckCircle2,
} from 'lucide-react';
import { firebaseProjectId } from '../utils/googleAuth';

interface VercelAuthHelpModalProps {
  isOpen: boolean;
  onClose: () => void;
  domain?: string | null;
  onRetryLogin?: () => void;
  isRetrying?: boolean;
}

export function VercelAuthHelpModal({
  isOpen,
  onClose,
  domain,
  onRetryLogin,
  isRetrying = false,
}: VercelAuthHelpModalProps) {
  const [copiedDomain, setCopiedDomain] = useState(false);
  const [copiedEnv, setCopiedEnv] = useState(false);
  const [activeTab, setActiveTab] = useState<'quick' | 'env'>('quick');

  if (!isOpen) return null;

  const currentDomain =
    domain || (typeof window !== 'undefined' ? window.location.hostname : 'seu-projeto.vercel.app');

  const firebaseSettingsUrl = `https://console.firebase.google.com/project/${firebaseProjectId}/authentication/settings`;

  const copyToClipboard = (text: string, type: 'domain' | 'env') => {
    navigator.clipboard.writeText(text);
    if (type === 'domain') {
      setCopiedDomain(true);
      setTimeout(() => setCopiedDomain(false), 2500);
    } else {
      setCopiedEnv(true);
      setTimeout(() => setCopiedEnv(false), 2500);
    }
  };

  const sampleEnvText = `VITE_FIREBASE_API_KEY="AIzaSy..."
VITE_FIREBASE_AUTH_DOMAIN="${firebaseProjectId}.firebaseapp.com"
VITE_FIREBASE_PROJECT_ID="${firebaseProjectId}"
VITE_FIREBASE_STORAGE_BUCKET="${firebaseProjectId}.firebasestorage.app"
VITE_FIREBASE_MESSAGING_SENDER_ID="67251400832"
VITE_FIREBASE_APP_ID="1:67251400832:web:..."`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="relative w-full max-w-2xl rounded-2xl border border-border bg-card p-6 shadow-2xl space-y-5 my-8 max-h-[90vh] overflow-y-auto">
        {/* Close Button */}
        <button
          type="button"
          onClick={onClose}
          className="absolute top-4 right-4 p-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted transition-colors cursor-pointer"
          title="Fechar"
        >
          <X className="h-5 w-5" />
        </button>

        {/* Modal Header */}
        <div className="flex items-start gap-3.5 pr-8">
          <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-amber-500/15 text-amber-600 dark:text-amber-400 shrink-0 border border-amber-500/20">
            <ShieldAlert className="h-6 w-6" />
          </div>
          <div>
            <span className="inline-flex items-center gap-1 text-[11px] font-semibold uppercase tracking-wider text-amber-600 dark:text-amber-400 bg-amber-500/10 px-2 py-0.5 rounded-full mb-1">
              auth/unauthorized-domain no Vercel
            </span>
            <h3 className="text-lg font-bold text-foreground">
              Como autorizar seu domínio no Firebase
            </h3>
            <p className="text-xs text-muted-foreground mt-0.5 leading-relaxed">
              O Google bloqueia o login em novos domínios por segurança até que o endereço seja
              autorizado nas configurações do Firebase Authentication.
            </p>
          </div>
        </div>

        {/* Detected Domain Box */}
        <div className="rounded-xl border border-primary/20 bg-primary/5 p-4 space-y-2">
          <p className="text-xs font-medium text-foreground">
            Domínio que precisa ser autorizado:
          </p>
          <div className="flex flex-col sm:flex-row sm:items-center gap-2">
            <div className="flex-1 flex items-center justify-between px-3 py-2 rounded-lg bg-background border border-border font-mono text-xs font-semibold text-primary">
              <span className="truncate">{currentDomain}</span>
            </div>
            <button
              type="button"
              onClick={() => copyToClipboard(currentDomain, 'domain')}
              className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-card border border-border hover:bg-muted text-xs font-medium text-foreground transition-colors cursor-pointer shrink-0 shadow-xs"
            >
              {copiedDomain ? (
                <>
                  <Check className="h-3.5 w-3.5 text-success" />
                  <span className="text-success font-semibold">Copiado!</span>
                </>
              ) : (
                <>
                  <Copy className="h-3.5 w-3.5 text-muted-foreground" />
                  <span>Copiar Domínio</span>
                </>
              )}
            </button>
          </div>
          <p className="text-[11px] text-muted-foreground">
            Dica: você também pode autorizar <strong className="font-mono text-foreground">vercel.app</strong> para cobrir todos os deploys de preview do Vercel.
          </p>
        </div>

        {/* Navigation Tabs */}
        <div className="flex border-b border-border text-xs font-medium">
          <button
            type="button"
            onClick={() => setActiveTab('quick')}
            className={`pb-2.5 px-3 border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
              activeTab === 'quick'
                ? 'border-primary text-primary font-semibold'
                : 'border-transparent text-muted-foreground hover:text-foreground'
            }`}
          >
            <CheckCircle2 className="h-3.5 w-3.5" />
            Opção 1: Adicionar Domínio (30 segundos)
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('env')}
            className={`pb-2.5 px-3 border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
              activeTab === 'env'
                ? 'border-primary text-primary font-semibold'
                : 'border-transparent text-muted-foreground hover:text-foreground'
            }`}
          >
            <KeyRound className="h-3.5 w-3.5" />
            Opção 2: Projeto Próprio no Vercel (.env)
          </button>
        </div>

        {/* Tab 1: Quick Add Domain in Firebase Console */}
        {activeTab === 'quick' && (
          <div className="space-y-4 text-xs leading-relaxed">
            <div className="space-y-3 rounded-xl border border-border bg-muted/20 p-4">
              <div className="flex items-start gap-2.5">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-primary text-[11px] font-bold text-primary-foreground">
                  1
                </span>
                <div className="flex-1">
                  <p className="font-semibold text-foreground">
                    Acesse a página de Domínios Autorizados no Firebase
                  </p>
                  <p className="text-muted-foreground mt-0.5">
                    Abra o Firebase Console do projeto <code className="font-mono text-[11px] bg-muted px-1.5 py-0.5 rounded">{firebaseProjectId}</code>:
                  </p>
                  <a
                    href={firebaseSettingsUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="mt-2 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-medium shadow-xs transition-colors cursor-pointer"
                  >
                    <span>Abrir Configurações do Firebase Console</span>
                    <ExternalLink className="h-3.5 w-3.5" />
                  </a>
                </div>
              </div>

              <div className="flex items-start gap-2.5 pt-2 border-t border-border/60">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-primary text-[11px] font-bold text-primary-foreground">
                  2
                </span>
                <div className="flex-1">
                  <p className="font-semibold text-foreground">
                    Adicione o domínio nas configurações
                  </p>
                  <p className="text-muted-foreground mt-0.5">
                    No Firebase, vá em <strong>Authentication</strong> &gt; aba <strong>Configurações</strong> (Settings) &gt; role até <strong>Domínios autorizados</strong> (Authorized domains).
                  </p>
                  <p className="text-muted-foreground mt-1">
                    Clique em <strong>Adicionar domínio</strong>, cole <strong className="font-mono text-foreground">{currentDomain}</strong> e clique em <strong>Salvar</strong>.
                  </p>
                </div>
              </div>

              <div className="flex items-start gap-2.5 pt-2 border-t border-border/60">
                <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-primary text-[11px] font-bold text-primary-foreground">
                  3
                </span>
                <div className="flex-1">
                  <p className="font-semibold text-foreground">
                    Pronto! Teste o login novamente
                  </p>
                  <p className="text-muted-foreground mt-0.5">
                    A liberação é instantânea. Basta clicar no botão abaixo para concluir o login com o Google.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Tab 2: Custom Firebase Project for Vercel */}
        {activeTab === 'env' && (
          <div className="space-y-4 text-xs leading-relaxed">
            <p className="text-muted-foreground">
              Se você prefere gerenciar seu próprio projeto Firebase para produção no Vercel:
            </p>

            <ol className="list-decimal list-inside space-y-1.5 text-muted-foreground pl-1">
              <li>Crie um projeto no <a href="https://console.firebase.google.com" target="_blank" rel="noopener noreferrer" className="text-primary hover:underline">Firebase Console</a>.</li>
              <li>Ative <strong>Authentication &gt; Google Sign-in</strong> e adicione seu domínio do Vercel nos domínios autorizados.</li>
              <li>Em <strong>Configurações do Projeto &gt; Aplicativos Web</strong>, copie as credenciais.</li>
              <li>No painel do Vercel, acesse <strong>Project Settings &gt; Environment Variables</strong> e adicione as variáveis abaixo:</li>
            </ol>

            <div className="relative rounded-lg border border-border bg-muted/60 p-3 font-mono text-[11px] text-foreground">
              <pre className="overflow-x-auto whitespace-pre">{sampleEnvText}</pre>
              <button
                type="button"
                onClick={() => copyToClipboard(sampleEnvText, 'env')}
                className="absolute top-2 right-2 inline-flex items-center gap-1 px-2 py-1 rounded bg-card border border-border text-[10px] font-medium text-foreground hover:bg-muted cursor-pointer shadow-2xs"
              >
                {copiedEnv ? (
                  <>
                    <Check className="h-3 w-3 text-success" />
                    <span className="text-success">Copiado!</span>
                  </>
                ) : (
                  <>
                    <Copy className="h-3 w-3" />
                    <span>Copiar Modelo</span>
                  </>
                )}
              </button>
            </div>
          </div>
        )}



        {/* Bottom Actions */}
        <div className="flex flex-col-reverse sm:flex-row sm:items-center sm:justify-end gap-2.5 pt-3 border-t border-border">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-medium rounded-lg border border-input bg-card hover:bg-accent text-foreground transition-colors cursor-pointer"
          >
            Fechar
          </button>
          {onRetryLogin && (
            <button
              type="button"
              disabled={isRetrying}
              onClick={onRetryLogin}
              className="inline-flex items-center justify-center gap-2 px-5 py-2 text-xs font-bold rounded-lg bg-primary hover:bg-primary/90 text-primary-foreground shadow transition-colors disabled:opacity-50 cursor-pointer"
            >
              {isRetrying ? (
                <>
                  <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                  <span>Conectando...</span>
                </>
              ) : (
                <>
                  <RefreshCw className="h-3.5 w-3.5" />
                  <span>Tentar Login Novamente</span>
                </>
              )}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
