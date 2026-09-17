import React, { useState } from 'react';
import ReactDOM from 'react-dom/client';
import GranolaNotepad from './components/GranolaNotepad';
import TeamsRecorderTab from '../TeamsRecorderTab';
import '../index.css';
import { Mic, Layers, ShieldCheck, User } from 'lucide-react';

function App() {
  const [viewMode, setViewMode] = useState<'NOTEPAD' | 'CLASSIC'>('NOTEPAD');
  const [userEmail, setUserEmail] = useState<string>('ankith.ravindran@mailplus.com.au');
  const [displayName, setDisplayName] = useState<string>('Ankith Ravindran');

  const apiBaseUrl = typeof window !== 'undefined' ? window.location.origin : 'http://localhost:5001';

  return (
    <div className="min-h-screen bg-[var(--bg-app-logged-in)] text-[var(--brand-ink)] flex flex-col font-sans">
      {/* Top Application Bar with Minutes.Plus Branding */}
      <header className="px-6 py-3 bg-[var(--bg-surface)] border-b border-[var(--border)] flex items-center justify-between text-xs shadow-sm">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-[var(--brand-primary)] flex items-center justify-center text-white shadow-md">
              <Mic className="w-4 h-4 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-black text-[var(--brand-primary)] text-base tracking-tight">Minutes.Plus</span>
                <span className="px-2 py-0.5 bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] font-bold rounded-full text-[10px] border border-[var(--brand-primary)]/20">
                  AI Notepad Engine
                </span>
              </div>
              <p className="text-[10px] text-[var(--brand-ink-soft)] -mt-0.5">Independent Meeting Recorder & AI Synthesis</p>
            </div>
          </div>

          <div className="h-5 w-px bg-[var(--border)] mx-1" />

          {/* Mode Switcher */}
          <div className="flex items-center bg-[var(--bg-cream)] p-1 rounded-lg border border-[var(--border)]">
            <button
              onClick={() => setViewMode('NOTEPAD')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md font-bold text-xs transition-all ${
                viewMode === 'NOTEPAD'
                  ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                  : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Mic className="w-3.5 h-3.5" />
              <span>AI Meeting Notepad</span>
            </button>

            <button
              onClick={() => setViewMode('CLASSIC')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md font-bold text-xs transition-all ${
                viewMode === 'CLASSIC'
                  ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                  : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>Recording Library & Hub</span>
            </button>
          </div>
        </div>

        {/* User Identity Pill */}
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 px-3 py-1.5 rounded-lg text-xs">
            <User className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span className="font-bold text-[var(--brand-primary)]">{displayName}</span>
            <span className="text-[var(--brand-ink-soft)] font-mono text-[11px]">({userEmail})</span>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 p-4 sm:p-6 flex flex-col justify-center">
        {viewMode === 'NOTEPAD' ? (
          <GranolaNotepad
            userEmail={userEmail}
            displayName={displayName}
            apiBaseUrl={apiBaseUrl}
          />
        ) : (
          <div className="max-w-6xl mx-auto w-full py-2">
            <TeamsRecorderTab />
          </div>
        )}
      </main>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
