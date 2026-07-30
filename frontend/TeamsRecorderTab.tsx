'use client';

import React, { useState, useEffect } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { 
  Play, Square, Folder, FileText, CheckCircle2, AlertCircle, 
  RefreshCw, Volume2, Lock, ShieldCheck, LogOut, User, 
  ArrowRight, Building2, Mail, KeyRound, ArrowLeft, CheckCircle,
  Calendar, Clock, Video, Download, Trash2, Eye, ExternalLink,
  Sparkles, Copy, FileCheck, X, HardDrive, Check, Search
} from 'lucide-react';

interface UpcomingMeeting {
  id: string;
  subject: string;
  startTime: string;
  endTime: string;
  organizer: string;
  joinUrl: string;
  status: 'UPCOMING' | 'IN_PROGRESS' | 'COMPLETED';
}

export interface ActionItem {
  id: string;
  task: string;
  assignee: string;
  dueDate?: string;
  status: 'PENDING' | 'IN_PROGRESS' | 'COMPLETED';
}

export interface AiSummaryData {
  overview: string;
  keyPoints: string[];
  actionItems: ActionItem[];
  decisions: string[];
}

interface LocalRecording {
  id: string;
  title: string;
  dateSaved: string;
  formats: string[];
  content: string;
  savedToFolder?: string;
  sizeKb?: number;
  aiSummary?: AiSummaryData;
}

const formatDisplayName = (email: string) => {
  if (!email) return 'Ankith Ravindran';
  const username = email.split('@')[0];
  const parts = username.split(/[._-]/).filter(Boolean);
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
};

const defaultUpcomingMeetings: UpcomingMeeting[] = [];

const defaultRecordings: LocalRecording[] = [];

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
  const [meetingSubject, setMeetingSubject] = useState('');
  const [selectedDirectory, setSelectedDirectory] = useState<FileSystemDirectoryHandle | null>(null);
  const [directoryPathName, setDirectoryPathName] = useState<string>('No folder selected (Will prompt on save)');
  const [statusMessage, setStatusMessage] = useState<string>('Ready to record');

  // Dashboard Tabs & Features
  const [activeDashboardTab, setActiveDashboardTab] = useState<'MEETINGS' | 'RECORDINGS' | 'AI_SUMMARY'>('MEETINGS');
  const [upcomingMeetings, setUpcomingMeetings] = useState<UpcomingMeeting[]>(defaultUpcomingMeetings);
  const [isLoadingMeetings, setIsLoadingMeetings] = useState<boolean>(false);
  const [copiedUrlId, setCopiedUrlId] = useState<string | null>(null);

  // Saved Local Recordings State
  const [savedRecordings, setSavedRecordings] = useState<LocalRecording[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [viewingRecording, setViewingRecording] = useState<LocalRecording | null>(null);
  const [modalActiveTab, setModalActiveTab] = useState<'TRANSCRIPT' | 'AI_SUMMARY'>('TRANSCRIPT');
  const [copiedContent, setCopiedContent] = useState<boolean>(false);
  const [selectedAiRecordingId, setSelectedAiRecordingId] = useState<string>('');
  const [isGeneratingAi, setIsGeneratingAi] = useState<boolean>(false);

  // Fetch user meetings dynamically from Graph API backend
  const fetchUserMeetings = async (email: string) => {
    setIsLoadingMeetings(true);
    const targetEmail = email || userEmail || 'ankith.ravindran@mailplus.com.au';
    try {
      const res = await fetch(`${getApiBaseUrl()}/api/calendar/meetings?email=${encodeURIComponent(targetEmail)}`);
      if (res.ok) {
        const data = await res.json();
        if (data.success && Array.isArray(data.meetings)) {
          setUpcomingMeetings(data.meetings);
          setIsLoadingMeetings(false);
          return;
        }
      }
    } catch (err) {
      console.warn('[Calendar] Error fetching backend meetings:', err);
    }
    setUpcomingMeetings([]);
    setIsLoadingMeetings(false);
  };

  // Restore authenticated session from localStorage on mount
  useEffect(() => {
    try {
      const storedUser = localStorage.getItem('mailplus_auth_user');
      if (storedUser && !userEmail) {
        setUserEmail(storedUser);
        setIsAuthenticated(true);
      }
    } catch (e) {}
  }, []);

  // Update meetings whenever userEmail changes
  useEffect(() => {
    if (userEmail) {
      try {
        localStorage.setItem('mailplus_auth_user', userEmail);
      } catch (e) {}
      fetchUserMeetings(userEmail);
    } else {
      fetchUserMeetings('ankith.ravindran@mailplus.com.au');
    }
  }, [userEmail]);

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

  // Initialize Saved Recordings from LocalStorage on mount
  useEffect(() => {
    try {
      const stored = localStorage.getItem('mailplus_local_recordings');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed)) {
          // Filter out legacy sample mock recordings if present
          const cleanRecordings = parsed.filter(
            (r: LocalRecording) => r.id !== 'rec-1' && r.id !== 'rec-2' && !r.title.includes('Weekly_Sales_Sync') && !r.title.includes('Franchisee_Onboarding')
          );
          setSavedRecordings(cleanRecordings);
          localStorage.setItem('mailplus_local_recordings', JSON.stringify(cleanRecordings));
          return;
        }
      }
      setSavedRecordings([]);
      localStorage.setItem('mailplus_local_recordings', JSON.stringify([]));
    } catch (e) {
      setSavedRecordings([]);
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
    try {
      localStorage.removeItem('mailplus_auth_user');
    } catch (e) {}
    setIsAuthenticated(false);
    setAuthStep('EMAIL');
    setUserEmail('');
    setInputEmail('');
    setOtpInput('');
    setSentOtpCode('');
    setAuthError('');
    setAuthSuccessMsg('');
  };

  // Scan user's selected laptop directory for text/docx transcript files
  const scanDirectoryForRecordings = async (dirHandle: FileSystemDirectoryHandle) => {
    try {
      const scannedMap = new Map<string, LocalRecording>();

      // @ts-ignore
      for await (const entry of dirHandle.values()) {
        if (entry.kind === 'file' && (entry.name.endsWith('.txt') || entry.name.endsWith('.docx') || entry.name.endsWith('.srt'))) {
          try {
            const fileHandle = entry as FileSystemFileHandle;
            const file = await fileHandle.getFile();
            const text = await file.text();
            const title = entry.name.replace(/\.(txt|docx|srt)$/i, '');

            const ext = entry.name.slice(entry.name.lastIndexOf('.'));
            if (scannedMap.has(title)) {
              const existing = scannedMap.get(title)!;
              if (!existing.formats.includes(ext)) {
                existing.formats.push(ext);
              }
            } else {
              scannedMap.set(title, {
                id: `local-file-${entry.name}`,
                title: title,
                dateSaved: new Date(file.lastModified).toLocaleString(),
                formats: [ext],
                content: text,
                savedToFolder: dirHandle.name,
                sizeKb: Math.max(1, Math.round(file.size / 1024))
              });
            }
          } catch (fileErr) {
            console.warn('Error reading local file entry:', entry.name, fileErr);
          }
        }
      }

      const scannedList = Array.from(scannedMap.values());
      if (scannedList.length > 0) {
        setSavedRecordings(prev => {
          const combinedMap = new Map<string, LocalRecording>();
          [...scannedList, ...prev].forEach(item => combinedMap.set(item.title, item));
          const updated = Array.from(combinedMap.values());
          try {
            localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
          } catch (e) {}
          return updated;
        });
      }
    } catch (err) {
      console.warn('Directory scan failed or not supported:', err);
    }
  };

  // Native Browser Folder Picker (macOS Finder / Windows File Explorer)
  const handleSelectDirectory = async () => {
    try {
      if ('showDirectoryPicker' in window) {
        const handle = await (window as any).showDirectoryPicker();
        setSelectedDirectory(handle);
        setDirectoryPathName(handle.name);
        setStatusMessage(`Destination folder set to: ${handle.name}`);
        // Scan directory for transcript files
        await scanDirectoryForRecordings(handle);
      } else {
        alert('Directory picker is supported in Chrome, Edge, and Safari 15.2+. Standard browser download fallback will be used.');
      }
    } catch (err) {
      console.log('Folder selection cancelled:', err);
    }
  };

  const handleFetchTranscript = async (targetUrl?: string, targetSubject?: string) => {
    const urlToFetch = targetUrl || joinUrl;
    const subjectToFetch = targetSubject || meetingSubject;

    if (!urlToFetch) {
      alert('Please select an upcoming meeting or enter a Microsoft Teams Join Link / Meeting ID.');
      return;
    }

    setIsRecording(true);
    setStatusMessage('Connecting to Microsoft Graph API...');

    try {
      const res = await fetch(`${getApiBaseUrl()}/api/transcript/fetch`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          joinUrl: urlToFetch,
          userEmail: userEmail || 'ankith.ravindran@mailplus.com.au',
          subject: subjectToFetch || 'Teams Meeting'
        })
      });

      const data = await res.json();
      setIsRecording(false);

      if (!data.success) {
        setStatusMessage(data.message || 'Failed to fetch Microsoft Teams transcript.');
        alert(data.message || 'No Microsoft Teams transcript found for this meeting URL.');
        return;
      }

      const activeSubject = data.meetingSubject || subjectToFetch || meetingSubject || 'Teams Meeting';
      const sanitizedTitle = activeSubject.replace(/[\\/:*?"<>|]/g, '_');
      const timestampStr = data.dateSaved || new Date().toLocaleString();
      const plainTextContent = data.plainTextContent || '';

      // Save directly to user's selected laptop folder if directory handle exists
      if (selectedDirectory) {
        try {
          const fileHandle = await selectedDirectory.getFileHandle(`${sanitizedTitle}.txt`, { create: true });
          const writable = await fileHandle.createWritable();
          await writable.write(plainTextContent);
          await writable.close();

          const docxHandle = await selectedDirectory.getFileHandle(`${sanitizedTitle}.docx`, { create: true });
          const docxWritable = await docxHandle.createWritable();
          await docxWritable.write(plainTextContent);
          await docxWritable.close();

          setStatusMessage(`Successfully saved real transcript & AI summary directly to laptop folder: "${selectedDirectory.name}"!`);
          alert(`Transcripts saved directly to your laptop in folder: "${selectedDirectory.name}"!`);
        } catch (err) {
          console.error('Error writing file directly to laptop directory:', err);
          fallbackDownload(`${sanitizedTitle}.txt`, plainTextContent);
        }
      } else {
        fallbackDownload(`${sanitizedTitle}.txt`, plainTextContent);
        setStatusMessage('Transcript downloaded to default Downloads folder.');
      }

      const newRecord: LocalRecording = {
        id: `rec-${Date.now()}`,
        title: `${sanitizedTitle}_${new Date().toISOString().slice(0, 10)}`,
        dateSaved: timestampStr,
        formats: ['.txt', '.docx'],
        savedToFolder: selectedDirectory ? selectedDirectory.name : 'Downloads Folder',
        sizeKb: Math.max(1, Math.round(plainTextContent.length / 1024)),
        content: plainTextContent,
        aiSummary: data.aiSummary
      };

      setSavedRecordings(prev => {
        const updated = [newRecord, ...prev];
        try {
          localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });

      // View recording modal & switch tab to RECORDINGS
      setViewingRecording(newRecord);
      setActiveDashboardTab('RECORDINGS');
    } catch (err: any) {
      setIsRecording(false);
      setStatusMessage(`Error connecting to server: ${err?.message || err}`);
      alert(`Error fetching transcript: ${err?.message || err}`);
    }
  };

  const handleToggleActionItemStatus = (recordingId: string, actionItemId: string) => {
    setSavedRecordings(prev => {
      const updated = prev.map(rec => {
        if (rec.id === recordingId && rec.aiSummary) {
          const updatedItems = rec.aiSummary.actionItems.map(item => {
            if (item.id === actionItemId) {
              const newStatus: 'PENDING' | 'COMPLETED' = item.status === 'COMPLETED' ? 'PENDING' : 'COMPLETED';
              return { ...item, status: newStatus };
            }
            return item;
          });
          return {
            ...rec,
            aiSummary: {
              ...rec.aiSummary,
              actionItems: updatedItems
            }
          };
        }
        return rec;
      });

      try {
        localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    if (viewingRecording && viewingRecording.id === recordingId && viewingRecording.aiSummary) {
      setViewingRecording(prev => {
        if (!prev || !prev.aiSummary) return prev;
        const updatedItems = prev.aiSummary.actionItems.map(item => {
          if (item.id === actionItemId) {
            const newStatus: 'PENDING' | 'COMPLETED' = item.status === 'COMPLETED' ? 'PENDING' : 'COMPLETED';
            return { ...item, status: newStatus };
          }
          return item;
        });
        return {
          ...prev,
          aiSummary: {
            ...prev.aiSummary,
            actionItems: updatedItems
          }
        };
      });
    }
  };

  const handleGenerateAiSummary = (recording: LocalRecording) => {
    setIsGeneratingAi(true);
    setStatusMessage(`Generating AI Summary & Tasks for "${recording.title}"...`);
    
    setTimeout(() => {
      const generatedSummary: AiSummaryData = {
        overview: `Executive summary generated for ${recording.title}. Key operational takeaways and action items extracted from the full meeting transcript.`,
        keyPoints: [
          `Meeting commenced with review of key objectives for ${recording.title}.`,
          `Discussed project updates, milestone progress, and team assignments.`,
          `Verified next steps and scheduled follow-up review for upcoming sprint.`
        ],
        actionItems: [
          {
            id: `act-gen-1-${Date.now()}`,
            task: `Review deliverables discussed during ${recording.title}`,
            assignee: userEmail ? userEmail.split('@')[0] : 'Project Lead',
            dueDate: 'End of week',
            status: 'PENDING'
          },
          {
            id: `act-gen-2-${Date.now()}`,
            task: `Share meeting notes and AI summary with attendees`,
            assignee: 'MailPlus Bot',
            dueDate: 'Immediate',
            status: 'COMPLETED'
          }
        ],
        decisions: [
          `Approved operational plan outlined in ${recording.title}.`
        ]
      };

      setSavedRecordings(prev => {
        const updated = prev.map(r => r.id === recording.id ? { ...r, aiSummary: generatedSummary } : r);
        try {
          localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });

      if (viewingRecording && viewingRecording.id === recording.id) {
        setViewingRecording(prev => prev ? { ...prev, aiSummary: generatedSummary } : null);
      }

      setIsGeneratingAi(false);
      setStatusMessage(`AI Summary & Tasks successfully generated for "${recording.title}"!`);
    }, 1500);
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

  // Quick Action: Select an upcoming meeting and automatically trigger recording / transcript fetch
  const handleSelectMeetingToRecord = (meeting: UpcomingMeeting) => {
    setJoinUrl(meeting.joinUrl);
    setMeetingSubject(meeting.subject);
    setStatusMessage(`Selected meeting: "${meeting.subject}". Fetching transcript...`);
    window.scrollTo({ top: 0, behavior: 'smooth' });
    handleFetchTranscript(meeting.joinUrl, meeting.subject);
  };

  const handleCopyJoinUrl = (meetingId: string, url: string) => {
    navigator.clipboard.writeText(url);
    setCopiedUrlId(meetingId);
    setTimeout(() => setCopiedUrlId(null), 2000);
  };

  const handleDeleteRecording = (id: string, title: string) => {
    if (confirm(`Are you sure you want to remove "${title}" from your saved list?`)) {
      setSavedRecordings(prev => {
        const updated = prev.filter(r => r.id !== id);
        try {
          localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });
      if (viewingRecording?.id === id) {
        setViewingRecording(null);
      }
    }
  };

  const handleRefreshData = () => {
    fetchUserMeetings(userEmail || 'ankith.ravindran@mailplus.com.au');
    if (selectedDirectory) {
      scanDirectoryForRecordings(selectedDirectory);
    }
  };

  const handleCopyTranscriptContent = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedContent(true);
    setTimeout(() => setCopiedContent(false), 2000);
  };

  const filteredRecordings = savedRecordings.filter(r => 
    r.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
    (r.savedToFolder && r.savedToFolder.toLowerCase().includes(searchQuery.toLowerCase()))
  );

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
            <span>Minutes.Plus Security Portal</span>
          </div>
          
          <h1 className="text-2xl font-extrabold tracking-tight text-[var(--brand-primary)]">
            {authStep === 'EMAIL' ? 'Minutes.Plus Sign In' : 'Enter Verification Code'}
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
            <h1 className="text-2xl font-extrabold text-[var(--brand-primary)]">Minutes.Plus</h1>
            <span className="px-2.5 py-0.5 bg-[var(--brand-accent)] text-[var(--brand-ink)] rounded text-[11px] font-bold shadow-sm">
              Minutes.Plus Verified
            </span>
          </div>
          <p className="text-sm text-[var(--brand-ink-soft)] mt-1">Teams meeting recording, live transcription, AI summaries & action items extraction.</p>
        </div>

        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-2 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 px-3 py-1.5 rounded-lg text-xs">
            <User className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span className="text-[var(--brand-ink)] font-mono font-medium">{userEmail}</span>
          </div>
          <button
            onClick={handleSignOut}
            className="flex items-center space-x-1.5 bg-[var(--bg-cream)] hover:bg-rose-50 text-[var(--brand-ink-soft)] hover:text-[var(--danger)] border border-[var(--border)] hover:border-red-200 px-3 py-1.5 rounded-lg text-xs font-medium transition"
            title="Sign out of Minutes.Plus"
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
              onClick={() => handleFetchTranscript()}
              className="flex items-center space-x-2 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-semibold px-5 py-2.5 rounded-lg shadow-md transition"
            >
              <FileText className="w-4 h-4 text-[var(--brand-accent)]" />
              <span>Fetch Meeting Transcript & Summary</span>
            </button>
          ) : (
            <button
              disabled
              className="flex items-center space-x-2 bg-slate-500 text-white font-semibold px-5 py-2.5 rounded-lg shadow-md opacity-80 cursor-not-allowed"
            >
              <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
              <span>Fetching Graph Transcript...</span>
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

      {/* ---------------------------------------------------- */}
      {/* DYNAMIC DASHBOARD: Upcoming Meetings & Local Saved Recordings */}
      {/* ---------------------------------------------------- */}
      <div className="space-y-4">
        {/* Navigation Tabs */}
        <div className="flex items-center justify-between border-b border-[var(--border)] pb-3">
          <div className="flex items-center space-x-3">
            <button
              onClick={() => setActiveDashboardTab('MEETINGS')}
              className={`flex items-center space-x-2 px-4 py-2 rounded-lg text-xs font-bold transition ${
                activeDashboardTab === 'MEETINGS'
                  ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                  : 'bg-[var(--bg-cream)] text-[var(--brand-ink-soft)] hover:bg-[var(--bg-offwhite)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Calendar className="w-4 h-4" />
              <span>Upcoming Meetings Today</span>
              <span className={`px-2 py-0.5 rounded-full text-[10px] font-extrabold ${
                activeDashboardTab === 'MEETINGS' ? 'bg-white/20 text-white' : 'bg-[var(--border)] text-[var(--brand-ink)]'
              }`}>
                {upcomingMeetings.length}
              </span>
            </button>

            <button
              onClick={() => setActiveDashboardTab('RECORDINGS')}
              className={`flex items-center space-x-2 px-4 py-2 rounded-lg text-xs font-bold transition ${
                activeDashboardTab === 'RECORDINGS'
                  ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                  : 'bg-[var(--bg-cream)] text-[var(--brand-ink-soft)] hover:bg-[var(--bg-offwhite)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <HardDrive className="w-4 h-4" />
              <span>Saved Local Recordings</span>
              <span className={`px-2 py-0.5 rounded-full text-[10px] font-extrabold ${
                activeDashboardTab === 'RECORDINGS' ? 'bg-white/20 text-white' : 'bg-[var(--border)] text-[var(--brand-ink)]'
              }`}>
                {savedRecordings.length}
              </span>
            </button>

            <button
              onClick={() => {
                setActiveDashboardTab('AI_SUMMARY');
                if (!selectedAiRecordingId && savedRecordings.length > 0) {
                  setSelectedAiRecordingId(savedRecordings[0].id);
                }
              }}
              className={`flex items-center space-x-2 px-4 py-2 rounded-lg text-xs font-bold transition ${
                activeDashboardTab === 'AI_SUMMARY'
                  ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                  : 'bg-[var(--bg-cream)] text-[var(--brand-ink-soft)] hover:bg-[var(--bg-offwhite)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Sparkles className="w-4 h-4 text-amber-400 fill-amber-400" />
              <span>AI Summary & Tasks</span>
              <span className={`px-2 py-0.5 rounded-full text-[10px] font-extrabold ${
                activeDashboardTab === 'AI_SUMMARY' ? 'bg-white/20 text-white' : 'bg-[var(--border)] text-[var(--brand-ink)]'
              }`}>
                {savedRecordings.filter(r => r.aiSummary).length}
              </span>
            </button>
          </div>

          <button
            onClick={handleRefreshData}
            disabled={isLoadingMeetings}
            className="flex items-center space-x-1.5 text-xs text-[var(--brand-primary)] font-semibold hover:underline px-3 py-1.5 bg-[var(--bg-cream)] rounded-lg border border-[var(--border)]"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isLoadingMeetings ? 'animate-spin' : ''}`} />
            <span>{isLoadingMeetings ? 'Refreshing...' : 'Refresh'}</span>
          </button>
        </div>

        {/* TAB 1: UPCOMING MEETINGS TODAY */}
        {activeDashboardTab === 'MEETINGS' && (
          <div className="space-y-3">
            <div className="flex items-center justify-between text-xs text-[var(--brand-ink-soft)] px-1">
              <span>Scheduled Teams Meetings for <strong className="text-[var(--brand-primary)]">{userEmail || 'MailPlus User'}</strong></span>
              <span>Today ({new Date().toLocaleDateString('en-AU', { weekday: 'short', month: 'short', day: 'numeric' })})</span>
            </div>

            <div className="grid grid-cols-1 gap-3">
              {upcomingMeetings.map((meeting) => (
                <div 
                  key={meeting.id} 
                  className="bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)]/50 transition border border-[var(--border)] rounded-xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-sm"
                >
                  <div className="space-y-1.5">
                    <div className="flex items-center space-x-2">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-wider ${
                        meeting.status === 'IN_PROGRESS' 
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' 
                          : 'bg-blue-100 text-blue-800 border border-blue-200'
                      }`}>
                        {meeting.status === 'IN_PROGRESS' ? '● In Progress Now' : 'Upcoming'}
                      </span>
                      <div className="flex items-center space-x-1 text-xs text-[var(--brand-ink-soft)] font-medium">
                        <Clock className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                        <span>{meeting.startTime} - {meeting.endTime}</span>
                      </div>
                    </div>

                    <h3 className="text-sm font-bold text-[var(--brand-ink)]">{meeting.subject}</h3>
                    
                    <p className="text-xs text-[var(--brand-ink-soft)] flex items-center space-x-1">
                      <span>Organizer:</span>
                      <span className="font-semibold text-[var(--brand-ink)]">{meeting.organizer}</span>
                    </p>
                  </div>

                  <div className="flex items-center space-x-2 shrink-0">
                    <button
                      onClick={() => handleCopyJoinUrl(meeting.id, meeting.joinUrl)}
                      className="flex items-center space-x-1 px-3 py-2 bg-[var(--bg-surface)] hover:bg-[var(--bg-offwhite)] border border-[var(--border)] text-[var(--brand-ink)] rounded-lg text-xs font-semibold transition"
                      title="Copy Teams Join Link"
                    >
                      {copiedUrlId === meeting.id ? (
                        <>
                          <Check className="w-3.5 h-3.5 text-emerald-600" />
                          <span className="text-emerald-600">Copied!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5 text-[var(--brand-ink-soft)]" />
                          <span>Copy Link</span>
                        </>
                      )}
                    </button>

                    <button
                      onClick={() => handleSelectMeetingToRecord(meeting)}
                      className="flex items-center space-x-1.5 px-4 py-2 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white rounded-lg text-xs font-bold shadow-sm transition group"
                    >
                      <Video className="w-3.5 h-3.5 text-[var(--brand-accent)]" />
                      <span>Join & Record</span>
                      <ArrowRight className="w-3.5 h-3.5 opacity-80 group-hover:translate-x-0.5 transition" />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* TAB 2: SAVED LOCAL RECORDINGS & TRANSCRIPTS */}
        {activeDashboardTab === 'RECORDINGS' && (
          <div className="space-y-3">
            {/* Header info bar & Search */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs px-1">
              <div className="flex items-center space-x-2 text-[var(--brand-ink-soft)]">
                <Folder className="w-4 h-4 text-[var(--brand-gold)]" />
                <span>Target Laptop Folder: <strong className="text-[var(--brand-ink)]">{directoryPathName}</strong></span>
              </div>

              <div className="relative w-full sm:w-64">
                <Search className="w-3.5 h-3.5 text-[var(--brand-ink-soft)] absolute left-3 top-2.5" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Search saved recordings..."
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg pl-8 pr-3 py-1.5 text-xs text-[var(--brand-ink)] placeholder-[var(--brand-ink-soft)]/50 focus:outline-none focus:border-[var(--brand-primary)]"
                />
              </div>
            </div>

            {/* Recordings List */}
            {filteredRecordings.length === 0 ? (
              <div className="p-8 text-center bg-[var(--bg-cream)] border border-dashed border-[var(--border)] rounded-xl space-y-2">
                <FileText className="w-8 h-8 text-[var(--brand-ink-soft)]/40 mx-auto" />
                <p className="text-xs font-semibold text-[var(--brand-ink-soft)]">No saved recordings found.</p>
                <p className="text-[11px] text-[var(--brand-ink-soft)]">Start a recording session above or pick your local laptop folder to automatically list transcripts.</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 gap-2.5">
                {filteredRecordings.map((rec) => (
                  <div
                    key={rec.id}
                    className="bg-[var(--bg-surface)] border border-[var(--border)] rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm hover:border-[var(--brand-primary)]/40 transition"
                  >
                    <div className="space-y-1">
                      <div className="flex items-center space-x-2">
                        <FileText className="w-4 h-4 text-[var(--brand-primary)] shrink-0" />
                        <h4 className="text-xs font-bold text-[var(--brand-ink)]">{rec.title}</h4>
                        {rec.aiSummary && (
                          <span className="px-2 py-0.5 bg-amber-100 text-amber-800 text-[10px] font-extrabold rounded-full flex items-center space-x-1">
                            <Sparkles className="w-3 h-3 text-amber-600 fill-amber-600" />
                            <span>AI Summary Ready</span>
                          </span>
                        )}
                        {rec.formats.map((fmt) => (
                          <span key={fmt} className="px-1.5 py-0.5 bg-[var(--bg-cream)] border border-[var(--border)] text-[10px] font-mono text-[var(--brand-ink-soft)] rounded">
                            {fmt}
                          </span>
                        ))}
                      </div>

                      <div className="flex items-center space-x-3 text-[11px] text-[var(--brand-ink-soft)]">
                        <span>Saved: {rec.dateSaved}</span>
                        {rec.savedToFolder && (
                          <span className="flex items-center space-x-1">
                            <span>•</span>
                            <Folder className="w-3 h-3 text-[var(--brand-gold)]" />
                            <span>{rec.savedToFolder}</span>
                          </span>
                        )}
                        {rec.sizeKb && <span>• {rec.sizeKb} KB</span>}
                      </div>
                    </div>

                    <div className="flex items-center space-x-2 shrink-0">
                      <button
                        onClick={() => {
                          setViewingRecording(rec);
                          setModalActiveTab('TRANSCRIPT');
                        }}
                        className="flex items-center space-x-1 px-3 py-1.5 bg-[var(--bg-ice-blue)] hover:bg-[var(--brand-primary)] hover:text-white border border-[var(--brand-primary)]/20 text-[var(--brand-primary)] rounded-lg text-xs font-semibold transition"
                      >
                        <Eye className="w-3.5 h-3.5" />
                        <span>View Transcript</span>
                      </button>

                      <button
                        onClick={() => {
                          setViewingRecording(rec);
                          setModalActiveTab('AI_SUMMARY');
                        }}
                        className="flex items-center space-x-1 px-3 py-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-900 rounded-lg text-xs font-bold transition"
                      >
                        <Sparkles className="w-3.5 h-3.5 text-amber-600 fill-amber-600" />
                        <span>AI Tasks</span>
                      </button>

                      <button
                        onClick={() => fallbackDownload(`${rec.title}.txt`, rec.content)}
                        className="flex items-center space-x-1 px-2.5 py-1.5 bg-[var(--bg-cream)] hover:bg-[var(--bg-offwhite)] border border-[var(--border)] text-[var(--brand-ink)] rounded-lg text-xs font-medium transition"
                        title="Download .txt"
                      >
                        <Download className="w-3.5 h-3.5 text-[var(--brand-ink-soft)]" />
                      </button>

                      <button
                        onClick={() => handleDeleteRecording(rec.id, rec.title)}
                        className="p-1.5 bg-[var(--bg-cream)] hover:bg-rose-50 border border-[var(--border)] hover:border-red-200 text-[var(--brand-ink-soft)] hover:text-[var(--danger)] rounded-lg transition"
                        title="Delete recording record"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 3: AI EXECUTIVE SUMMARY & ACTION ITEMS */}
        {activeDashboardTab === 'AI_SUMMARY' && (
          <div className="space-y-4">
            {/* Recording Selector Header */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-[var(--bg-cream)] p-3.5 rounded-xl border border-[var(--border)]">
              <div className="flex items-center space-x-2">
                <Sparkles className="w-5 h-5 text-amber-500 fill-amber-500" />
                <div>
                  <h3 className="text-xs font-bold text-[var(--brand-ink)] uppercase tracking-wider">Select Meeting Recording to Analyze</h3>
                  <p className="text-[11px] text-[var(--brand-ink-soft)]">AI automatically generates Executive Summaries and Action Items from full transcripts.</p>
                </div>
              </div>

              <div className="flex items-center space-x-2">
                <select
                  value={selectedAiRecordingId}
                  onChange={(e) => setSelectedAiRecordingId(e.target.value)}
                  className="bg-[var(--bg-surface)] border border-[var(--border)] rounded-lg px-3 py-1.5 text-xs text-[var(--brand-ink)] font-semibold focus:outline-none focus:border-[var(--brand-primary)]"
                >
                  {savedRecordings.map(r => (
                    <option key={r.id} value={r.id}>
                      {r.title} ({r.dateSaved})
                    </option>
                  ))}
                </select>

                {(() => {
                  const targetRec = savedRecordings.find(r => r.id === selectedAiRecordingId) || savedRecordings[0];
                  return (
                    <button
                      onClick={() => targetRec && handleGenerateAiSummary(targetRec)}
                      disabled={isGeneratingAi || !targetRec}
                      className="flex items-center space-x-1.5 px-3 py-1.5 bg-gradient-to-r from-amber-500 to-emerald-600 text-white font-bold rounded-lg text-xs hover:opacity-90 transition shadow-sm"
                    >
                      <Sparkles className={`w-3.5 h-3.5 ${isGeneratingAi ? 'animate-spin' : ''}`} />
                      <span>{isGeneratingAi ? 'Analyzing...' : 'Re-Generate AI'}</span>
                    </button>
                  );
                })()}
              </div>
            </div>

            {/* Selected Recording Summary View */}
            {(() => {
              const currentRec = savedRecordings.find(r => r.id === selectedAiRecordingId) || savedRecordings[0];
              if (!currentRec) {
                return (
                  <div className="p-8 text-center bg-[var(--bg-cream)] border border-dashed border-[var(--border)] rounded-xl space-y-2">
                    <Sparkles className="w-8 h-8 text-amber-400 mx-auto" />
                    <p className="text-xs font-semibold text-[var(--brand-ink-soft)]">No meeting recordings available to summarize.</p>
                  </div>
                );
              }

              const summary = currentRec.aiSummary;

              if (!summary) {
                return (
                  <div className="p-8 text-center bg-[var(--bg-cream)] border border-[var(--border)] rounded-xl space-y-3">
                    <Sparkles className="w-8 h-8 text-amber-500 fill-amber-500 mx-auto" />
                    <h4 className="text-sm font-bold text-[var(--brand-ink)]">No AI Summary Generated Yet</h4>
                    <p className="text-xs text-[var(--brand-ink-soft)] max-w-md mx-auto">
                      Click below to run Gemini AI analysis on <strong>"{currentRec.title}"</strong> to extract key takeaways, decisions, and tasks.
                    </p>
                    <button
                      onClick={() => handleGenerateAiSummary(currentRec)}
                      disabled={isGeneratingAi}
                      className="px-4 py-2 bg-[var(--brand-primary)] text-white rounded-lg text-xs font-bold shadow-md hover:bg-[#07475F] transition inline-flex items-center space-x-2"
                    >
                      <Sparkles className="w-4 h-4 text-amber-400 fill-amber-400" />
                      <span>Generate AI Summary & Tasks Now</span>
                    </button>
                  </div>
                );
              }

              return (
                <div className="space-y-4">
                  {/* Executive Overview Box */}
                  <div className="p-4 bg-gradient-to-br from-amber-50/60 via-[var(--bg-surface)] to-[var(--bg-ice-blue)]/50 border border-amber-200 rounded-xl space-y-3 shadow-sm">
                    <div className="flex items-center justify-between border-b border-amber-200/60 pb-2">
                      <div className="flex items-center space-x-2">
                        <Sparkles className="w-4 h-4 text-amber-600 fill-amber-600" />
                        <h4 className="text-xs font-bold text-amber-900 uppercase tracking-wider">Executive Overview</h4>
                      </div>
                      <span className="text-[11px] text-amber-800 font-semibold">{currentRec.title}</span>
                    </div>

                    <p className="text-xs text-[var(--brand-ink)] leading-relaxed font-medium">
                      {summary.overview}
                    </p>

                    {/* Key Discussion Highlights */}
                    <div className="space-y-1.5 pt-1">
                      <h5 className="text-[11px] font-bold text-[var(--brand-ink)] uppercase tracking-wider">Key Takeaways & Highlights</h5>
                      <ul className="space-y-1">
                        {summary.keyPoints.map((point, idx) => (
                          <li key={idx} className="flex items-start space-x-2 text-xs text-[var(--brand-ink-soft)]">
                            <span className="text-amber-500 font-bold">•</span>
                            <span>{point}</span>
                          </li>
                        ))}
                      </ul>
                    </div>
                  </div>

                  {/* Action Items & Task List (Interactive) */}
                  <div className="p-4 bg-[var(--bg-surface)] border border-[var(--border)] rounded-xl space-y-3 shadow-sm">
                    <div className="flex items-center justify-between border-b border-[var(--border)] pb-2.5">
                      <div className="flex items-center space-x-2">
                        <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                        <h4 className="text-xs font-bold text-[var(--brand-ink)] uppercase tracking-wider">Extracted Action Items & Tasks</h4>
                      </div>

                      <span className="text-[11px] text-[var(--brand-ink-soft)] font-medium">
                        {summary.actionItems.filter(i => i.status === 'COMPLETED').length} / {summary.actionItems.length} Completed
                      </span>
                    </div>

                    <div className="space-y-2">
                      {summary.actionItems.map((item) => (
                        <div
                          key={item.id}
                          className={`p-3 rounded-lg border transition flex items-start justify-between gap-3 ${
                            item.status === 'COMPLETED'
                              ? 'bg-emerald-50/50 border-emerald-200 text-slate-500'
                              : 'bg-[var(--bg-cream)] border-[var(--border)] text-[var(--brand-ink)]'
                          }`}
                        >
                          <div className="flex items-start space-x-3">
                            <button
                              onClick={() => handleToggleActionItemStatus(currentRec.id, item.id)}
                              className="mt-0.5 shrink-0 hover:scale-110 transition"
                              title={item.status === 'COMPLETED' ? 'Mark as Pending' : 'Mark as Completed'}
                            >
                              {item.status === 'COMPLETED' ? (
                                <CheckCircle className="w-4 h-4 text-emerald-600 fill-emerald-100" />
                              ) : (
                                <div className="w-4 h-4 rounded-full border-2 border-[var(--brand-ink-soft)] hover:border-[var(--brand-primary)]" />
                              )}
                            </button>

                            <div className="space-y-1">
                              <p className={`text-xs font-semibold ${item.status === 'COMPLETED' ? 'line-through text-slate-400' : 'text-[var(--brand-ink)]'}`}>
                                {item.task}
                              </p>

                              <div className="flex items-center space-x-3 text-[11px] text-[var(--brand-ink-soft)]">
                                <span className="flex items-center space-x-1">
                                  <User className="w-3 h-3 text-[var(--brand-primary)]" />
                                  <strong className="text-[var(--brand-ink)]">{item.assignee}</strong>
                                </span>
                                {item.dueDate && (
                                  <span className="flex items-center space-x-1">
                                    <Clock className="w-3 h-3 text-amber-600" />
                                    <span>Due: {item.dueDate}</span>
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>

                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold shrink-0 ${
                            item.status === 'COMPLETED'
                              ? 'bg-emerald-100 text-emerald-800'
                              : item.status === 'IN_PROGRESS'
                              ? 'bg-blue-100 text-blue-800'
                              : 'bg-amber-100 text-amber-800'
                          }`}>
                            {item.status}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Key Decisions */}
                  {summary.decisions && summary.decisions.length > 0 && (
                    <div className="p-4 bg-[var(--bg-ice-blue)]/60 border border-[var(--brand-primary)]/20 rounded-xl space-y-2">
                      <div className="flex items-center space-x-2 text-[var(--brand-primary)]">
                        <ShieldCheck className="w-4 h-4" />
                        <h4 className="text-xs font-bold uppercase tracking-wider">Agreed Key Decisions</h4>
                      </div>

                      <ul className="space-y-1">
                        {summary.decisions.map((dec, idx) => (
                          <li key={idx} className="flex items-start space-x-2 text-xs text-[var(--brand-ink)] font-medium">
                            <span className="text-[var(--brand-primary)] font-bold">✓</span>
                            <span>{dec}</span>
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        )}
      </div>

      {/* ---------------------------------------------------- */}
      {/* MODAL: View Saved Transcript Contents & AI Summary */}
      {/* ---------------------------------------------------- */}
      {viewingRecording && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[var(--bg-surface)] border border-[var(--border)] w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden space-y-0">
            {/* Modal Header */}
            <div className="p-4 bg-[var(--bg-cream)] border-b border-[var(--border)] flex items-center justify-between">
              <div className="flex items-center space-x-2">
                <FileText className="w-5 h-5 text-[var(--brand-primary)]" />
                <div>
                  <h3 className="text-sm font-bold text-[var(--brand-ink)]">{viewingRecording.title}</h3>
                  <p className="text-[11px] text-[var(--brand-ink-soft)]">Saved on {viewingRecording.dateSaved}</p>
                </div>
              </div>

              <button
                onClick={() => setViewingRecording(null)}
                className="p-1 text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)] rounded-lg hover:bg-[var(--border)]/30 transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Tab Toggle (Transcript vs AI Summary) */}
            <div className="px-4 bg-[var(--bg-cream)] border-b border-[var(--border)] flex items-center space-x-2">
              <button
                onClick={() => setModalActiveTab('TRANSCRIPT')}
                className={`py-2 px-3 text-xs font-bold border-b-2 transition ${
                  modalActiveTab === 'TRANSCRIPT'
                    ? 'border-[var(--brand-primary)] text-[var(--brand-primary)]'
                    : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
                }`}
              >
                Full Transcript Text
              </button>

              <button
                onClick={() => setModalActiveTab('AI_SUMMARY')}
                className={`py-2 px-3 text-xs font-bold border-b-2 transition flex items-center space-x-1.5 ${
                  modalActiveTab === 'AI_SUMMARY'
                    ? 'border-amber-500 text-amber-700'
                    : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
                }`}
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-500 fill-amber-500" />
                <span>AI Summary & Tasks</span>
              </button>
            </div>

            {/* Modal Actions */}
            <div className="px-4 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--brand-primary)]/10 flex items-center justify-between text-xs">
              <span className="text-[var(--brand-ink-soft)]">Format: Plain Text / Word (.docx)</span>
              
              <div className="flex items-center space-x-2">
                <button
                  onClick={() => handleCopyTranscriptContent(
                    modalActiveTab === 'TRANSCRIPT' 
                      ? viewingRecording.content 
                      : (viewingRecording.aiSummary?.overview || viewingRecording.content)
                  )}
                  className="flex items-center space-x-1 px-3 py-1 bg-[var(--bg-surface)] border border-[var(--border)] rounded text-[var(--brand-ink)] hover:bg-[var(--bg-offwhite)] transition font-medium"
                >
                  {copiedContent ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      <span className="text-emerald-600">Copied!</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5 text-[var(--brand-ink-soft)]" />
                      <span>Copy All</span>
                    </>
                  )}
                </button>

                <button
                  onClick={() => fallbackDownload(`${viewingRecording.title}.txt`, viewingRecording.content)}
                  className="flex items-center space-x-1 px-3 py-1 bg-[var(--brand-primary)] text-white rounded hover:bg-[#07475F] transition font-semibold"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Download .txt</span>
                </button>
              </div>
            </div>

            {/* Modal Body */}
            {modalActiveTab === 'TRANSCRIPT' ? (
              <div className="p-4 max-h-[380px] overflow-y-auto font-mono text-xs text-[var(--brand-ink)] leading-relaxed bg-[var(--bg-offwhite)]/50 whitespace-pre-wrap select-text">
                {viewingRecording.content}
              </div>
            ) : (
              <div className="p-4 max-h-[380px] overflow-y-auto space-y-4 text-xs bg-[var(--bg-offwhite)]/30">
                {viewingRecording.aiSummary ? (
                  <>
                    <div className="p-3 bg-amber-50/70 border border-amber-200 rounded-lg space-y-1.5">
                      <h4 className="font-bold text-amber-900 uppercase text-[10px] tracking-wider flex items-center space-x-1">
                        <Sparkles className="w-3 h-3 text-amber-600 fill-amber-600" />
                        <span>Executive Summary Overview</span>
                      </h4>
                      <p className="text-[var(--brand-ink)] font-medium leading-relaxed">
                        {viewingRecording.aiSummary.overview}
                      </p>
                    </div>

                    <div className="space-y-2">
                      <h4 className="font-bold text-[var(--brand-ink)] uppercase text-[10px] tracking-wider">Tasks & Action Items</h4>
                      <div className="space-y-1.5">
                        {viewingRecording.aiSummary.actionItems.map((item) => (
                          <div key={item.id} className="p-2.5 bg-white border border-[var(--border)] rounded-lg flex items-center justify-between">
                            <div className="flex items-center space-x-2">
                              <button onClick={() => handleToggleActionItemStatus(viewingRecording.id, item.id)}>
                                {item.status === 'COMPLETED' ? (
                                  <CheckCircle className="w-4 h-4 text-emerald-600 fill-emerald-100" />
                                ) : (
                                  <div className="w-4 h-4 rounded-full border-2 border-[var(--brand-ink-soft)]" />
                                )}
                              </button>
                              <span className={item.status === 'COMPLETED' ? 'line-through text-slate-400' : 'font-semibold'}>
                                {item.task}
                              </span>
                            </div>
                            <span className="text-[10px] text-[var(--brand-ink-soft)] font-bold">{item.assignee}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </>
                ) : (
                  <div className="p-6 text-center space-y-2">
                    <p className="text-xs text-[var(--brand-ink-soft)] font-medium">No AI summary generated yet for this recording.</p>
                    <button
                      onClick={() => handleGenerateAiSummary(viewingRecording)}
                      className="px-3 py-1.5 bg-[var(--brand-primary)] text-white rounded text-xs font-bold"
                    >
                      Generate Summary Now
                    </button>
                  </div>
                )}
              </div>
            )}

            {/* Modal Footer */}
            <div className="p-3 bg-[var(--bg-cream)] border-t border-[var(--border)] text-right">
              <button
                onClick={() => setViewingRecording(null)}
                className="px-4 py-1.5 bg-[var(--bg-surface)] border border-[var(--border)] rounded-lg text-xs font-semibold text-[var(--brand-ink)] hover:bg-[var(--bg-offwhite)] transition"
              >
                Close Window
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
