import React, { useState } from 'react';
import ReactDOM from 'react-dom/client';
import GranolaNotepad from './components/GranolaNotepad';
import TeamsRecorderTab from '../TeamsRecorderTab';
import '../index.css';
import { Flame, Layers, ShieldCheck, User } from 'lucide-react';

function App() {
  const [viewMode, setViewMode] = useState<'GRANOLA' | 'CLASSIC'>('GRANOLA');
  const [userEmail, setUserEmail] = useState<string>('ankith.ravindran@mailplus.com.au');
  const [displayName, setDisplayName] = useState<string>('Ankith Ravindran');

  const apiBaseUrl = typeof window !== 'undefined' ? window.location.origin : 'http://localhost:5001';

  return (
    <div className="min-h-screen bg-[#090d14] text-slate-100 flex flex-col font-sans">
      {/* Top Application Bar */}
      <header className="px-6 py-2.5 bg-[#0e131d] border-b border-slate-800/80 flex items-center justify-between text-xs">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-gradient-to-tr from-amber-600 to-orange-500 flex items-center justify-center text-slate-950 font-black shadow-md shadow-amber-900/30">
              <Flame className="w-4 h-4 fill-current text-slate-950" />
            </div>
            <div>
              <span className="font-bold text-white text-sm tracking-tight">Granola</span>
              <span className="text-[10px] text-amber-400 font-mono ml-1 font-semibold">AI Notepad</span>
            </div>
          </div>

          <div className="h-4 w-px bg-slate-800 mx-1" />

          {/* Mode Switcher */}
          <div className="flex items-center bg-slate-900 p-0.5 rounded-lg border border-slate-800">
            <button
              onClick={() => setViewMode('GRANOLA')}
              className={`flex items-center gap-1.5 px-3 py-1 rounded-md font-semibold transition-all ${
                viewMode === 'GRANOLA'
                  ? 'bg-amber-500 text-slate-950 shadow-sm'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Flame className="w-3 h-3" />
              <span>Granola Notepad</span>
            </button>

            <button
              onClick={() => setViewMode('CLASSIC')}
              className={`flex items-center gap-1.5 px-3 py-1 rounded-md font-semibold transition-all ${
                viewMode === 'CLASSIC'
                  ? 'bg-slate-800 text-slate-200 shadow-sm'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Layers className="w-3 h-3" />
              <span>Classic Hub & Storage</span>
            </button>
          </div>
        </div>

        {/* User Identity Pill */}
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 bg-slate-900 px-3 py-1 rounded-full border border-slate-800 text-[11px] text-slate-300">
            <User className="w-3 h-3 text-amber-400" />
            <span className="font-medium">{displayName}</span>
            <span className="text-slate-600">({userEmail})</span>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 p-4 flex flex-col justify-center">
        {viewMode === 'GRANOLA' ? (
          <GranolaNotepad
            userEmail={userEmail}
            displayName={displayName}
            apiBaseUrl={apiBaseUrl}
          />
        ) : (
          <div className="max-w-6xl mx-auto w-full py-6">
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
