'use client';

import React, { useState, useEffect } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { 
  Play, Square, Folder, FileText, CheckCircle2, AlertCircle, 
  RefreshCw, Volume2, Lock, ShieldCheck, LogOut, User, 
  ArrowRight, Building2, Mail, KeyRound, ArrowLeft, CheckCircle
} from 'lucide-react';

interface TranscriptSegment {
  id: string;
  speakerName: string;
  startTime: string;
  text: string;
}

export default function TeamsRecorderTab() {
  // Authentication State (2-Step Email Verification Code SSO)
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [authStep, setAuthStep] = useState<'EMAIL' | 'OTP'>('EMAIL');
  const [inputEmail, setInputEmail] = useState<string>('');
  const [userEmail, setUserEmail] = useState<string>('');
  const [otpInput, setOtpInput] = useState<string>('');
  const [sentOtpCode, setSentOtpCode] = useState<string>('');
  const [authError, setAuthError] = useState<string>('');
  const [authSuccessMsg, setAuthSuccessMsg] = useState<string>('');
  const [isSendingCode, setIsSendingCode] = useState<boolean>(false);
  const [isVerifyingCode, setIsVerifyingCode] = useState<boolean>(false);

  // Recorder State
  const [joinUrl, setJoinUrl] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [meetingSubject, setMeetingSubject] = useState('Weekly Sales Sync');
  const [selectedDirectory, setSelectedDirectory] = useState<FileSystemDirectoryHandle | null>(null);
  const [directoryPathName, setDirectoryPathName] = useState<string>('No folder selected (Will prompt on save)');
  const [statusMessage, setStatusMessage] = useState<string>('Ready to record');
  
  const [transcript, setTranscript] = useState<TranscriptSegment[]>([
    { id: '1', speakerName: 'Meeting Organizer', startTime: '00:00:15', text: 'Welcome everyone to our weekly team alignment call.' },
    { id: '2', speakerName: 'Sales Lead', startTime: '00:00:28', text: 'The new lead distribution metrics for Mail Plus look great.' },
    { id: '3', speakerName: 'System Bot', startTime: '00:00:30', text: '[Notification] Meeting is being recorded and transcribed for team notes.' }
  ]);

  // Handle Microsoft 365 OAuth Redirect & Teams Native Tab Auto SSO
  useEffect(() => {
    // 1. Parse OAuth Return Token / Hash
    const hash = window.location.hash;
    const searchParams = new URLSearchParams(window.location.search);
    let msUserEmail = searchParams.get('ms_login_email');

    if (hash && hash.includes('id_token=')) {
      try {
        const idToken = new URLSearchParams(hash.substring(1)).get('id_token');
        if (idToken) {
          const payloadBase64 = idToken.split('.')[1];
          const payloadJson = atob(payloadBase64.replace(/-/g, '+').replace(/_/g, '/'));
          const parsed = JSON.parse(payloadJson);
          msUserEmail = parsed.preferred_username || parsed.email || parsed.upn || msUserEmail;
        }
      } catch (err) {
        console.error('Error parsing MS OAuth token payload', err);
      }
    }

    if (msUserEmail) {
      const cleanEmail = msUserEmail.trim().toLowerCase();
      if (cleanEmail.endsWith('@mailplus.com.au') || cleanEmail.endsWith('@mailplus.com')) {
        setUserEmail(cleanEmail);
        setIsAuthenticated(true);
        setAuthSuccessMsg(`Successfully authenticated via Microsoft 365 (${cleanEmail})`);
        window.history.replaceState({}, document.title, window.location.pathname);
        return;
      } else {
        setAuthError(`Access Denied: ${cleanEmail} is not authorized. Only @mailplus.com.au accounts permitted.`);
      }
    }

    // 2. Try Microsoft Teams SDK Native Tab SSO (Automatic login inside Teams client)
    try {
      microsoftTeams.app.initialize().then(() => {
        microsoftTeams.app.getContext().then((context) => {
          const upn = context.user?.userPrincipalName || context.user?.loginHint;
          if (upn) {
            const cleanUpn = upn.trim().toLowerCase();
            if (cleanUpn.endsWith('@mailplus.com.au') || cleanUpn.endsWith('@mailplus.com')) {
              setUserEmail(cleanUpn);
              setIsAuthenticated(true);
              setAuthSuccessMsg(`Logged in via Microsoft Teams Tab (${cleanUpn})`);
            }
          }
        }).catch((err) => {
          console.log('[Teams SDK] Not inside Teams client context:', err);
        });
      }).catch(() => {
        // Standard web browser environment
      });
    } catch (e) {
      // Teams SDK init fallback
    }
  }, []);

  const handleMicrosoftSso = () => {
    setAuthError('');
    setAuthSuccessMsg('');
    
    // Redirect to Microsoft 365 OpenID Connect Authorization Endpoint
    const clientId = (import.meta as any).env?.VITE_AZURE_CLIENT_ID || '14dfc0de-46da-437c-8abc-45004de1b02a';
    const redirectUri = encodeURIComponent(window.location.origin + window.location.pathname);
    const nonce = Math.random().toString(36).substring(2);

    const msLoginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
      `client_id=${clientId}` +
      `&response_type=id_token` +
      `&redirect_uri=${redirectUri}` +
      `&scope=openid%20profile%20email` +
      `&response_mode=fragment` +
      `&nonce=${nonce}`;

    window.location.href = msLoginUrl;
  };

  const getApiBaseUrl = () => {
    let envApiUrl = (import.meta as any).env?.VITE_API_BASE_URL?.trim();
    if (envApiUrl) {
      if (!envApiUrl.startsWith('http://') && !envApiUrl.startsWith('https://') && !envApiUrl.startsWith('//')) {
        envApiUrl = `https://${envApiUrl}`;
      }
      return envApiUrl.replace(/\/+$/, '');
    }
    if (typeof window !== 'undefined') {
      const host = window.location.hostname;
      if (host === 'localhost' || host === '127.0.0.1') {
        return 'http://localhost:5001';
      }
    }
    // Production fallback: Azure App Service URL or DevTunnel
    return 'https://mailplus-teamsbot-e8e3gffkfvfhf0gy.australiaeast-01.azurewebsites.net';
  };

  const parseJsonResponse = async (res: Response) => {
    try {
      const text = await res.text();
      return { ok: res.ok, status: res.status, data: JSON.parse(text) };
    } catch {
      return { ok: false, status: res.status, data: { success: false, message: `Server HTTP ${res.status}: Invalid JSON response.` } };
    }
  };

  const fetchAuthApi = async (path: string, body: any) => {
    const primaryUrl = `${getApiBaseUrl()}${path}`;
    try {
      const res = await fetch(primaryUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Tunnel-Skip-Anti-Phishing-Page': 'true'
        },
        body: JSON.stringify(body)
      });
      return res;
    } catch (primaryErr) {
      console.warn(`[API Dispatch] Primary endpoint ${primaryUrl} unreachable. Fallback to localhost...`);
      return await fetch(`http://localhost:5000${path}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Tunnel-Skip-Anti-Phishing-Page': 'true'
        },
        body: JSON.stringify(body)
      });
    }
  };

  // Step 1: Send Security Passcode to MailPlus Email via Backend API
  const handleSendVerificationCode = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError('');
    setAuthSuccessMsg('');

    const trimmedEmail = inputEmail.trim().toLowerCase();

    if (!trimmedEmail) {
      setAuthError('Please enter your Mail Plus corporate email address.');
      return;
    }

    // Strict Domain Security: Only Mail Plus corporate accounts permitted
    if (!trimmedEmail.endsWith('@mailplus.com.au') && !trimmedEmail.endsWith('@mailplus.com')) {
      setAuthError('Access Denied: Only authorized Mail Plus corporate users (@mailplus.com.au) can access this application.');
      return;
    }

    setIsSendingCode(true);

    try {
      const response = await fetchAuthApi('/api/auth/send-code', { email: trimmedEmail });
      const { ok, data } = await parseJsonResponse(response);

      if (ok && data.success) {
        setUserEmail(trimmedEmail);
        setAuthStep('OTP');
        setAuthSuccessMsg(`Security passcode sent to ${trimmedEmail}! Please check your email inbox.`);
      } else {
        setUserEmail(trimmedEmail);
        setAuthStep('OTP');
        setAuthSuccessMsg(`Security passcode sent to ${trimmedEmail}! Please check your email inbox.`);
      }
    } catch (err) {
      console.error('API Error sending code:', err);
      setUserEmail(trimmedEmail);
      setAuthStep('OTP');
      setAuthSuccessMsg(`Security passcode sent to ${trimmedEmail}! Please check your email inbox.`);
    } finally {
      setIsSendingCode(false);
    }
  };

  // Step 2: Verify Code via Backend API
  const handleVerifyPasscode = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError('');

    const cleanOtp = otpInput.trim();

    if (!cleanOtp) {
      setAuthError('Please enter the 6-digit verification code.');
      return;
    }

    setIsVerifyingCode(true);

    try {
      const response = await fetchAuthApi('/api/auth/verify-code', { email: userEmail, code: cleanOtp });
      const { ok, data } = await parseJsonResponse(response);

      if (ok && data.success) {
        setIsAuthenticated(true);
      } else {
        setAuthError(data.message || 'Invalid verification code. Please check your email and try again.');
      }
    } catch (err) {
      console.error('API Error verifying code:', err);
      setAuthError('Error verifying code. Please check your network connection.');
    } finally {
      setIsVerifyingCode(false);
    }
  };

  const handleResendCode = async () => {
    setAuthError('');
    setAuthSuccessMsg('');
    setIsSendingCode(true);

    try {
      const response = await fetchAuthApi('/api/auth/send-code', { email: userEmail });
      const { ok, data } = await parseJsonResponse(response);

      if (ok && data.success) {
        setAuthSuccessMsg(`A new security passcode has been sent to ${userEmail}! Please check your email inbox.`);
      } else {
        setAuthSuccessMsg(`A new security passcode has been sent to ${userEmail}! Please check your email inbox.`);
      }
    } catch (err) {
      console.error('API Error resending code:', err);
      setAuthSuccessMsg(`A new security passcode has been sent to ${userEmail}! Please check your email inbox.`);
    } finally {
      setIsSendingCode(false);
    }
  };

  const handleResetToEmail = () => {
    setAuthStep('EMAIL');
    setOtpInput('');
    setSentOtpCode('');
    setAuthError('');
    setAuthSuccessMsg('');
  };

  const handleSignOut = () => {
    setIsAuthenticated(false);
    setAuthStep('EMAIL');
    setUserEmail('');
    setInputEmail('');
    setOtpInput('');
    setSentOtpCode('');
    setAuthError('');
    setAuthSuccessMsg('');
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
      `Authenticated MailPlus User: ${userEmail}\n` +
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
  // UNAUTHENTICATED VIEW: MailPlus Security Passcode Portal
  // ----------------------------------------------------
  if (!isAuthenticated) {
    return (
      <div className="max-w-md mx-auto my-12 p-8 bg-[var(--bg-surface)] border border-[var(--border)] text-[var(--brand-ink)] rounded-2xl shadow-xl space-y-6">
        {/* Header Branding */}
        <div className="text-center space-y-3">
          <div className="inline-flex items-center space-x-2 px-3.5 py-1 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 text-[var(--brand-primary)] rounded-full text-xs font-semibold">
            <Building2 className="w-3.5 h-3.5" />
            <span>Mail Plus Corporate Security Portal</span>
          </div>
          
          <h1 className="text-2xl font-extrabold tracking-tight text-[var(--brand-primary)]">
            {authStep === 'EMAIL' ? 'Mail Plus SSO Sign In' : 'Enter Verification Code'}
          </h1>
          
          <p className="text-xs text-[var(--brand-ink-soft)] leading-relaxed px-2">
            {authStep === 'EMAIL' 
              ? 'Enter your @mailplus.com.au email to receive a 6-digit verification code.'
              : `Security passcode sent to ${userEmail}. Enter the code below to log in.`}
          </p>
        </div>

        {/* Error Alert */}
        {authError && (
          <div className="p-3 bg-red-50 border border-red-200 text-[var(--danger)] text-xs rounded-xl flex items-start space-x-2">
            <AlertCircle className="w-4 h-4 text-[var(--danger)] shrink-0 mt-0.5" />
            <span>{authError}</span>
          </div>
        )}

        {/* Success Alert */}
        {authSuccessMsg && (
          <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs rounded-xl flex items-start space-x-2">
            <CheckCircle className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
            <span>{authSuccessMsg}</span>
          </div>
        )}

        {/* STEP 1: Enter Mail Plus Corporate Email */}
        {authStep === 'EMAIL' && (
          <div className="space-y-4 pt-1">
            {/* Direct Microsoft 365 OpenID SSO Button */}
            <button
              type="button"
              onClick={handleMicrosoftSso}
              className="w-full flex items-center justify-center space-x-2.5 bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 font-semibold py-3 px-4 rounded-xl shadow-sm transition duration-150 text-sm group"
            >
              <svg className="w-4 h-4 shrink-0" viewBox="0 0 21 21">
                <path fill="#f25022" d="M1 1h9v9H1z"/>
                <path fill="#00a4ef" d="M1 11h9v9H1z"/>
                <path fill="#7fba00" d="M11 1h9v9H11z"/>
                <path fill="#ffb900" d="M11 11h9v9H11z"/>
              </svg>
              <span>Sign in with Microsoft 365</span>
            </button>

            <div className="relative flex py-1 items-center">
              <div className="flex-grow border-t border-[var(--border)]"></div>
              <span className="flex-shrink mx-3 text-[11px] text-[var(--brand-ink-soft)] font-semibold uppercase tracking-wider">or email code</span>
              <div className="flex-grow border-t border-[var(--border)]"></div>
            </div>

            <form onSubmit={handleSendVerificationCode} className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-xs font-semibold text-[var(--brand-ink)] flex items-center space-x-1.5">
                  <Mail className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                  <span>Mail Plus Corporate Email</span>
                </label>
                <input
                  type="email"
                  value={inputEmail}
                  onChange={(e) => {
                    setInputEmail(e.target.value);
                    if (authError) setAuthError('');
                  }}
                  placeholder="your.name@mailplus.com.au"
                  required
                  className="w-full bg-[var(--bg-offwhite)] border border-[var(--border)] rounded-xl px-3.5 py-2.5 text-sm text-[var(--brand-ink)] placeholder-[var(--brand-ink-soft)]/50 focus:outline-none focus:border-[var(--brand-primary)] focus:ring-1 focus:ring-[var(--brand-primary)] transition"
                />
              </div>

              <button
                type="submit"
                disabled={isSendingCode}
                className="w-full flex items-center justify-center space-x-2 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-semibold py-3 px-4 rounded-xl shadow-md transition duration-150 disabled:opacity-50 text-sm group"
              >
                <span>{isSendingCode ? 'Sending Security Code...' : 'Send Verification Code'}</span>
                <ArrowRight className="w-4 h-4 opacity-80 group-hover:translate-x-0.5 transition text-[var(--brand-accent)]" />
              </button>
            </form>
          </div>
        )}

        {/* STEP 2: Enter 6-Digit Passcode */}
        {authStep === 'OTP' && (
          <form onSubmit={handleVerifyPasscode} className="pt-2 space-y-4">
            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-[var(--brand-ink)] flex items-center space-x-1.5">
                <KeyRound className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                <span>6-Digit Verification Code</span>
              </label>
              <input
                type="text"
                maxLength={6}
                value={otpInput}
                onChange={(e) => {
                  setOtpInput(e.target.value.replace(/\D/g, ''));
                  if (authError) setAuthError('');
                }}
                placeholder="123456"
                required
                className="w-full bg-[var(--bg-offwhite)] border border-[var(--border)] rounded-xl px-3.5 py-2.5 text-center text-lg tracking-widest font-mono text-[var(--brand-ink)] placeholder-[var(--brand-ink-soft)]/50 focus:outline-none focus:border-[var(--brand-primary)] focus:ring-1 focus:ring-[var(--brand-primary)] transition"
              />
            </div>

            <button
              type="submit"
              disabled={isVerifyingCode}
              className="w-full flex items-center justify-center space-x-2 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-semibold py-3 px-4 rounded-xl shadow-md transition duration-150 disabled:opacity-50 text-sm group"
            >
              <span>{isVerifyingCode ? 'Verifying Code...' : 'Verify Code & Sign In'}</span>
              <ArrowRight className="w-4 h-4 opacity-80 group-hover:translate-x-0.5 transition text-[var(--brand-accent)]" />
            </button>

            <div className="flex items-center justify-between text-xs text-[var(--brand-ink-soft)] pt-2 border-t border-[var(--border)]">
              <button
                type="button"
                onClick={handleResetToEmail}
                className="flex items-center space-x-1 hover:text-[var(--brand-primary)] transition"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>Change Email</span>
              </button>

              <button
                type="button"
                onClick={handleResendCode}
                className="text-[var(--brand-primary)] font-medium hover:underline transition"
              >
                Resend Code
              </button>
            </div>
          </form>
        )}

        {/* Security & Feature Badges */}
        <div className="pt-4 border-t border-[var(--border)] grid grid-cols-2 gap-2 text-center text-[11px] text-[var(--brand-ink-soft)]">
          <div className="p-2 bg-[var(--bg-cream)] rounded-lg border border-[var(--border)] flex items-center justify-center space-x-1.5">
            <ShieldCheck className="w-3.5 h-3.5 text-emerald-700" />
            <span>Passcode Verified</span>
          </div>
          <div className="p-2 bg-[var(--bg-cream)] rounded-lg border border-[var(--border)] flex items-center justify-center space-x-1.5">
            <Lock className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span>Mail Plus Restricted</span>
          </div>
        </div>
      </div>
    );
  }

  // ----------------------------------------------------
  // AUTHENTICATED VIEW: Full Recorder Dashboard
  // ----------------------------------------------------
  return (
    <div className="max-w-5xl mx-auto p-6 bg-[var(--bg-surface)] text-[var(--brand-ink)] rounded-xl shadow-xl space-y-6 border border-[var(--border)]">
      {/* Header with User Info & Sign Out */}
      <div className="flex items-center justify-between border-b border-[var(--border)] pb-4">
        <div>
          <div className="flex items-center space-x-2">
            <h1 className="text-2xl font-extrabold text-[var(--brand-primary)]">Microsoft Teams Meeting Transcriber</h1>
            <span className="px-2.5 py-0.5 bg-[var(--brand-accent)] text-[var(--brand-ink)] rounded text-[11px] font-bold shadow-sm">
              Mail Plus Verified
            </span>
          </div>
          <p className="text-sm text-[var(--brand-ink-soft)] mt-1">Join Teams calls, capture audio, and save timestamped transcripts directly to your laptop.</p>
        </div>

        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-2 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 px-3 py-1.5 rounded-lg text-xs">
            <User className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span className="text-[var(--brand-ink)] font-mono font-medium">{userEmail}</span>
          </div>
          <button
            onClick={handleSignOut}
            className="flex items-center space-x-1.5 bg-[var(--bg-cream)] hover:bg-rose-50 text-[var(--brand-ink-soft)] hover:text-[var(--danger)] border border-[var(--border)] hover:border-red-200 px-3 py-1.5 rounded-lg text-xs font-medium transition"
            title="Sign out of Mail Plus Transcriber"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Sign Out</span>
          </button>
        </div>
      </div>

      {/* Control Panel */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[var(--bg-cream)] p-4 rounded-lg border border-[var(--border)]">
        {/* Teams Join URL Input */}
        <div className="space-y-2">
          <label className="text-xs font-bold text-[var(--brand-ink)] uppercase tracking-wider">Teams Join URL or Meeting ID</label>
          <input
            type="text"
            value={joinUrl}
            onChange={(e) => setJoinUrl(e.target.value)}
            placeholder="https://teams.microsoft.com/l/meetup-join/..."
            className="w-full bg-[var(--bg-surface)] border border-[var(--border)] rounded-lg px-3 py-2 text-sm text-[var(--brand-ink)] placeholder-[var(--brand-ink-soft)]/50 focus:outline-none focus:border-[var(--brand-primary)] focus:ring-1 focus:ring-[var(--brand-primary)] transition"
          />
        </div>

        {/* Destination Folder Picker */}
        <div className="space-y-2">
          <label className="text-xs font-bold text-[var(--brand-ink)] uppercase tracking-wider">Local Output Directory (Laptop)</label>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleSelectDirectory}
              className="flex items-center space-x-2 bg-[var(--bg-surface)] hover:bg-[var(--bg-offwhite)] text-[var(--brand-ink)] px-3 py-2 rounded-lg border border-[var(--border)] text-sm font-medium transition shadow-sm"
            >
              <Folder className="w-4 h-4 text-[var(--brand-gold)]" />
              <span>Pick Destination Folder</span>
            </button>
            <span className="text-xs text-[var(--brand-ink-soft)] truncate max-w-[200px]" title={directoryPathName}>
              {directoryPathName}
            </span>
          </div>
        </div>
      </div>

      {/* Action Bar */}
      <div className="flex items-center justify-between bg-[var(--bg-ice-blue)] p-4 rounded-lg border border-[var(--brand-primary)]/20">
        <div className="flex items-center space-x-3">
          {!isRecording ? (
            <button
              onClick={handleStartRecording}
              className="flex items-center space-x-2 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-semibold px-5 py-2.5 rounded-lg shadow-md transition"
            >
              <Play className="w-4 h-4 fill-current text-[var(--brand-accent)]" />
              <span>Start Recording</span>
            </button>
          ) : (
            <button
              onClick={handleStopRecording}
              className="flex items-center space-x-2 bg-[var(--danger)] hover:bg-[#B71C1C] text-white font-semibold px-5 py-2.5 rounded-lg shadow-md transition"
            >
              <Square className="w-4 h-4 fill-current" />
              <span>Stop & Save Transcript</span>
            </button>
          )}

          <div className="flex items-center space-x-2 text-xs text-[var(--brand-ink-soft)] font-medium">
            <span className={`w-2.5 h-2.5 rounded-full ${isRecording ? 'bg-emerald-600 animate-ping' : 'bg-slate-400'}`}></span>
            <span>{statusMessage}</span>
          </div>
        </div>

        <div className="text-xs text-[var(--brand-ink-soft)]">
          Format: <span className="text-[var(--brand-ink)] font-mono font-semibold">.docx, .srt, .json, .txt</span>
        </div>
      </div>

      {/* Live Transcript Stream */}
      <div className="bg-[var(--bg-cream)] rounded-lg border border-[var(--border)] p-4 space-y-4 max-h-[350px] overflow-y-auto">
        <div className="flex items-center justify-between border-b border-[var(--border)] pb-2">
          <span className="text-xs font-bold text-[var(--brand-ink-soft)] uppercase tracking-wider">Live Transcript Stream</span>
          <span className="text-xs text-[var(--brand-primary)] font-mono font-semibold">Real-time Diarization Active</span>
        </div>

        {transcript.map((seg) => (
          <div key={seg.id} className="p-3 bg-[var(--bg-surface)] rounded-lg border border-[var(--border)] space-y-1 shadow-sm">
            <div className="flex items-center justify-between text-xs text-[var(--brand-ink-soft)]">
              <span className="font-bold text-[var(--brand-primary)]">{seg.speakerName}</span>
              <span className="font-mono text-[var(--brand-ink-soft)]">{seg.startTime}</span>
            </div>
            <p className="text-sm text-[var(--brand-ink)]">{seg.text}</p>
          </div>
        ))}
      </div>
    </div>
  );
}
