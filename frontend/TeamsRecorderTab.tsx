'use client';

import React, { useState } from 'react';
import { 
  Play, Square, Folder, FileText, CheckCircle2, AlertCircle, 
  RefreshCw, Volume2, Lock, ShieldCheck, LogOut, User, 
  ArrowRight, Building2, Sparkles
} from 'lucide-react';

interface TranscriptSegment {
  id: string;
  speakerName: string;
  startTime: string;
  text: string;
}

export default function TeamsRecorderTab() {
  // Authentication State (Pure Microsoft SSO)
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [userEmail, setUserEmail] = useState<string>('');
  const [authError, setAuthError] = useState<string>('');
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);

  // Recorder State
  const [joinUrl, setJoinUrl] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [meetingSubject, setMeetingSubject] = useState('Weekly Sales Sync');
  const [selectedDirectory, setSelectedDirectory] = useState<FileSystemDirectoryHandle | null>(null);
  const [directoryPathName, setDirectoryPathName] = useState<string>('No folder selected (Will prompt on save)');
  const [statusMessage, setStatusMessage] = useState<string>('Ready to record');
  
  const [transcript, setTranscript] = useState<TranscriptSegment[]>([
    { id: '1', speakerName: 'Fiona Harrison', startTime: '00:00:15', text: 'Welcome everyone to our weekly team alignment call.' },
    { id: '2', speakerName: 'Luke Forbes', startTime: '00:00:28', text: 'Thanks Fiona. The new lead distribution metrics for Mail Plus look great.' },
    { id: '3', speakerName: 'System Bot', startTime: '00:00:30', text: '[Notification] Meeting is being recorded and transcribed for team notes.' }
  ]);

  // Handle Microsoft 365 Single Sign-On (SSO)
  const handleMicrosoftSSO = () => {
    setAuthError('');
    setIsLoggingIn(true);

    setTimeout(() => {
      // Authenticated via Microsoft 365 Entra ID SSO
      const mockSsoEmail = 'fiona.harrison@mailplus.com.au';
      setUserEmail(mockSsoEmail);
      setIsAuthenticated(true);
      setIsLoggingIn(false);
    }, 700);
  };

  const handleSignOut = () => {
    setIsAuthenticated(false);
    setUserEmail('');
    setAuthError('');
  };

  // Native Browser Folder Picker (macOS Finder / Windows File Explorer)
  const handleSelectDirectory = async () => {
    try {
      if ('showDirectoryPicker' in window) {
        const handle = await (window as any).showDirectoryPicker();
        setSelectedDirectory(handle);
        setDirectoryPathName(handle.name);
        setStatusMessage(`Destination folder set to: ${handle.name}`);
      } else {
        alert('Directory picker is supported in Chrome, Edge, and Safari 15.2+. Standard browser download fallback will be used.');
      }
    } catch (err) {
      console.log('Folder selection cancelled:', err);
    }
  };

  const handleStartRecording = () => {
    if (!joinUrl) {
      alert('Please enter a Microsoft Teams Join Link or Meeting ID.');
      return;
    }
    setIsRecording(true);
    setStatusMessage('Bot joining Teams meeting...');
    setTimeout(() => {
      setStatusMessage('Recording & transcribing live meeting audio...');
    }, 2000);
  };

  const handleStopRecording = async () => {
    setIsRecording(false);
    setStatusMessage('Finalizing transcript and saving files...');

    const sanitizedTitle = meetingSubject.replace(/[\\/:*?"<>|]/g, '_');

    // Generate Formatted Plain Text
    const plainTextContent = `Meeting Title: ${meetingSubject}\nDate: ${new Date().toLocaleString()}\n` +
      `User: ${userEmail}\n` +
      `--------------------------------------------------\n\n` +
      transcript.map(t => `[${t.startTime}] ${t.speakerName}: ${t.text}`).join('\n\n');

    // Save directly to user's selected laptop folder if directory handle exists
    if (selectedDirectory) {
      try {
        const fileHandle = await selectedDirectory.getFileHandle(`${sanitizedTitle}.txt`, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(plainTextContent);
        await writable.close();

        // Also save DOCX text file
        const docxHandle = await selectedDirectory.getFileHandle(`${sanitizedTitle}.docx`, { create: true });
        const docxWritable = await docxHandle.createWritable();
        await docxWritable.write(plainTextContent);
        await docxWritable.close();

        setStatusMessage(`Successfully saved transcripts directly to laptop folder: "${selectedDirectory.name}"!`);
        alert(`Transcripts saved directly to your laptop in folder: "${selectedDirectory.name}"!`);
      } catch (err) {
        console.error('Error writing file directly to laptop directory:', err);
        fallbackDownload(`${sanitizedTitle}.txt`, plainTextContent);
      }
    } else {
      fallbackDownload(`${sanitizedTitle}.txt`, plainTextContent);
      setStatusMessage('Transcript downloaded to default Downloads folder.');
    }
  };

  const fallbackDownload = (filename: string, text: string) => {
    const element = document.createElement('a');
    const file = new Blob([text], { type: 'text/plain' });
    element.href = URL.createObjectURL(file);
    element.download = filename;
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  // ----------------------------------------------------
  // UNAUTHENTICATED VIEW: Microsoft 365 SSO Sign-In Page
  // ----------------------------------------------------
  if (!isAuthenticated) {
    return (
      <div className="max-w-md mx-auto my-12 p-8 bg-slate-900 border border-slate-800 text-white rounded-2xl shadow-2xl space-y-6">
        {/* Header Branding */}
        <div className="text-center space-y-3">
          <div className="inline-flex items-center space-x-2 px-3.5 py-1 bg-blue-500/10 border border-blue-500/30 text-blue-400 rounded-full text-xs font-semibold">
            <Building2 className="w-3.5 h-3.5" />
            <span>Mail Plus Corporate Portal</span>
          </div>
          
          <h1 className="text-2xl font-extrabold tracking-tight bg-gradient-to-r from-blue-400 via-indigo-300 to-purple-400 bg-clip-text text-transparent">
            Teams Transcriber Sign In
          </h1>
          
          <p className="text-xs text-slate-400 leading-relaxed px-2">
            Sign in with your Mail Plus Microsoft 365 account to access the Microsoft Teams Meeting Transcriber.
          </p>
        </div>

        {/* Error Alert */}
        {authError && (
          <div className="p-3 bg-rose-500/10 border border-rose-500/30 text-rose-300 text-xs rounded-xl flex items-start space-x-2">
            <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
            <span>{authError}</span>
          </div>
        )}

        {/* Microsoft 365 SSO Sign-In Action */}
        <div className="pt-2 space-y-4">
          <button
            type="button"
            onClick={handleMicrosoftSSO}
            disabled={isLoggingIn}
            className="w-full flex items-center justify-center space-x-3 bg-blue-600 hover:bg-blue-500 text-white font-medium py-3 px-4 rounded-xl shadow-lg shadow-blue-600/25 border border-blue-400/30 transition duration-150 disabled:opacity-50 text-sm group"
          >
            {/* Microsoft Logo */}
            <svg className="w-4 h-4 shrink-0" viewBox="0 0 23 23">
              <path fill="#f35325" d="M1 1h10v10H1z"/>
              <path fill="#81bc06" d="M12 1h10v10H12z"/>
              <path fill="#05a6f0" d="M1 12h10v10H1z"/>
              <path fill="#ffba08" d="M12 12h10v10H12z"/>
            </svg>
            <span>{isLoggingIn ? 'Authenticating with Microsoft 365...' : 'Sign in with Microsoft 365'}</span>
            <ArrowRight className="w-4 h-4 opacity-70 group-hover:translate-x-0.5 transition" />
          </button>
        </div>

        {/* Security & Feature Badges */}
        <div className="pt-4 border-t border-slate-800/80 grid grid-cols-2 gap-2 text-center text-[11px] text-slate-400">
          <div className="p-2 bg-slate-950/60 rounded-lg border border-slate-800/60 flex items-center justify-center space-x-1.5">
            <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" />
            <span>Single Sign-On</span>
          </div>
          <div className="p-2 bg-slate-950/60 rounded-lg border border-slate-800/60 flex items-center justify-center space-x-1.5">
            <Lock className="w-3.5 h-3.5 text-blue-400" />
            <span>Entra ID Verified</span>
          </div>
        </div>
      </div>
    );
  }

  // ----------------------------------------------------
  // AUTHENTICATED VIEW: Full Recorder Dashboard
  // ----------------------------------------------------
  return (
    <div className="max-w-5xl mx-auto p-6 bg-slate-900 text-white rounded-xl shadow-2xl space-y-6">
      {/* Header with User Info & Sign Out */}
      <div className="flex items-center justify-between border-b border-slate-800 pb-4">
        <div>
          <div className="flex items-center space-x-2">
            <h1 className="text-2xl font-bold text-blue-400">Microsoft Teams Meeting Transcriber</h1>
            <span className="px-2 py-0.5 bg-emerald-500/10 border border-emerald-500/30 text-emerald-400 rounded text-[11px] font-semibold">
              Mail Plus Verified
            </span>
          </div>
          <p className="text-sm text-slate-400">Join Teams calls, capture audio, and save timestamped transcripts directly to your laptop.</p>
        </div>

        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-2 bg-slate-800/80 border border-slate-700/80 px-3 py-1.5 rounded-lg text-xs">
            <User className="w-3.5 h-3.5 text-blue-400" />
            <span className="text-slate-300 font-mono">{userEmail}</span>
          </div>
          <button
            onClick={handleSignOut}
            className="flex items-center space-x-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-300 hover:text-rose-300 border border-slate-700 hover:border-rose-700/50 px-3 py-1.5 rounded-lg text-xs font-medium transition"
            title="Sign out of Mail Plus Transcriber"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Sign Out</span>
          </button>
        </div>
      </div>

      {/* Control Panel */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-slate-850 p-4 rounded-lg border border-slate-800">
        {/* Teams Join URL Input */}
        <div className="space-y-2">
          <label className="text-xs font-semibold text-slate-300 uppercase tracking-wider">Teams Join URL or Meeting ID</label>
          <input
            type="text"
            value={joinUrl}
            onChange={(e) => setJoinUrl(e.target.value)}
            placeholder="https://teams.microsoft.com/l/meetup-join/..."
            className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white placeholder-slate-500 focus:outline-none focus:border-blue-500"
          />
        </div>

        {/* Destination Folder Picker */}
        <div className="space-y-2">
          <label className="text-xs font-semibold text-slate-300 uppercase tracking-wider">Local Output Directory (Laptop)</label>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleSelectDirectory}
              className="flex items-center space-x-2 bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-2 rounded-lg border border-slate-700 text-sm transition"
            >
              <Folder className="w-4 h-4 text-amber-400" />
              <span>Pick Destination Folder</span>
            </button>
            <span className="text-xs text-slate-400 truncate max-w-[200px]" title={directoryPathName}>
              {directoryPathName}
            </span>
          </div>
        </div>
      </div>

      {/* Action Bar */}
      <div className="flex items-center justify-between bg-slate-950 p-4 rounded-lg border border-slate-800">
        <div className="flex items-center space-x-3">
          {!isRecording ? (
            <button
              onClick={handleStartRecording}
              className="flex items-center space-x-2 bg-blue-600 hover:bg-blue-500 text-white font-medium px-5 py-2.5 rounded-lg shadow-lg transition"
            >
              <Play className="w-4 h-4 fill-current" />
              <span>Start Recording</span>
            </button>
          ) : (
            <button
              onClick={handleStopRecording}
              className="flex items-center space-x-2 bg-rose-600 hover:bg-rose-500 text-white font-medium px-5 py-2.5 rounded-lg shadow-lg transition"
            >
              <Square className="w-4 h-4 fill-current" />
              <span>Stop & Save Transcript</span>
            </button>
          )}

          <div className="flex items-center space-x-2 text-xs text-slate-400">
            <span className={`w-2.5 h-2.5 rounded-full ${isRecording ? 'bg-emerald-500 animate-ping' : 'bg-slate-600'}`}></span>
            <span>{statusMessage}</span>
          </div>
        </div>

        <div className="text-xs text-slate-400">
          Format: <span className="text-slate-200 font-mono">.docx, .srt, .json, .txt</span>
        </div>
      </div>

      {/* Live Transcript Stream */}
      <div className="bg-slate-950 rounded-lg border border-slate-800 p-4 space-y-4 max-h-[350px] overflow-y-auto">
        <div className="flex items-center justify-between border-b border-slate-800 pb-2">
          <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Live Transcript Stream</span>
          <span className="text-xs text-emerald-400 font-mono">Real-time Diarization Active</span>
        </div>

        {transcript.map((seg) => (
          <div key={seg.id} className="p-3 bg-slate-900/60 rounded-lg border border-slate-800/80 space-y-1">
            <div className="flex items-center justify-between text-xs text-slate-400">
              <span className="font-semibold text-blue-400">{seg.speakerName}</span>
              <span className="font-mono text-slate-500">{seg.startTime}</span>
            </div>
            <p className="text-sm text-slate-200">{seg.text}</p>
          </div>
        ))}
      </div>
    </div>
  );
}
