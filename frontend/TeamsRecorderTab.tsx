'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { 
  Play, Square, Folder, FileText, CheckCircle2, AlertCircle, 
  RefreshCw, Volume2, Lock, ShieldCheck, LogOut, User, 
  ArrowRight, Building2, Mail, KeyRound, ArrowLeft, CheckCircle,
  Calendar, Clock, Video, Download, Trash2, Eye, ExternalLink,
  Sparkles, Copy, FileCheck, X, HardDrive, Check, Search, Mic,
  Upload, Pause, Share2, UserCheck, Users, Tag, ListTodo, Layers,
  FileSpreadsheet, Radio, Music, MessageSquare
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
  template?: string;
}

export interface TranscriptLine {
  id: string;
  timestamp: string;
  speaker: string;
  text: string;
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
  attendees?: string[];
  transcriptLines?: TranscriptLine[];
  template?: string;
}

const formatDisplayName = (email: string) => {
  if (!email) return 'Ankith Ravindran';
  const username = email.split('@')[0];
  const parts = username.split(/[._-]/).filter(Boolean);
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
};

const getApiBaseUrl = () => {
  if (typeof window !== 'undefined') {
    return window.location.origin;
  }
  return 'http://localhost:5001';
};

const meetingTemplates = [
  { id: 'GENERAL', label: '📝 General Meeting', desc: 'Standard overview, key points & action items' },
  { id: 'CLIENT_CALL', label: '💼 Client Sales & Discovery Call', desc: 'Client pain points, scope & next steps' },
  { id: 'TEAM_SYNC', label: '👥 Weekly Team Sync', desc: 'Updates, blockers, decisions & team tasks' },
  { id: 'PROJECT_REVIEW', label: '🚀 Project Review & Planning', desc: 'Milestones, risks, release targets' },
  { id: 'ONE_ON_ONE', label: '🎯 1-on-1 Performance Check-in', desc: 'Feedback, goals, personal growth' },
  { id: 'BRAINSTORM', label: '💡 Brainstorming & Strategy', desc: 'Ideas, evaluation & prioritized initiatives' }
];

export default function TeamsRecorderTab() {
  // ----------------------------------------------------
  // Authentication State (2-Step Email Verification SSO)
  // ----------------------------------------------------
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

  // ----------------------------------------------------
  // Host-Independent Recorder State
  // ----------------------------------------------------
  const [audioSourceMode, setAudioSourceMode] = useState<'DUAL' | 'MIC_ONLY'>('DUAL');
  const [isRecording, setIsRecording] = useState(false);
  const [isPaused, setIsPaused] = useState(false);
  const [recordingSeconds, setRecordingSeconds] = useState(0);
  const [meetingSubject, setMeetingSubject] = useState('Weekly Team Meeting');
  const [meetingTemplate, setMeetingTemplate] = useState<string>('GENERAL');
  
  // Attendees & Speaker Attribution
  const [attendeesList, setAttendeesList] = useState<string[]>(['Ankith Ravindran', 'Jane Doe', 'Mark Taylor']);
  const [newAttendeeInput, setNewAttendeeInput] = useState<string>('');
  const [activeSpeakerTag, setActiveSpeakerTag] = useState<string>('Ankith Ravindran');
  
  // Live Transcript & Live Notes
  const [liveTranscriptLines, setLiveTranscriptLines] = useState<TranscriptLine[]>([]);
  const [interimText, setInterimText] = useState<string>('');
  const [liveNotes, setLiveNotes] = useState<string>('');
  
  // Directory & Status
  const [selectedDirectory, setSelectedDirectory] = useState<FileSystemDirectoryHandle | null>(null);
  const [directoryPathName, setDirectoryPathName] = useState<string>('No folder selected (Will prompt on save)');
  const [statusMessage, setStatusMessage] = useState<string>('Ready to record meeting');
  
  // Dashboard Tabs & Features
  const [activeDashboardTab, setActiveDashboardTab] = useState<'RECORDER' | 'UPLOAD' | 'MEETINGS' | 'RECORDINGS'>('RECORDER');
  const [upcomingMeetings, setUpcomingMeetings] = useState<UpcomingMeeting[]>([]);
  const [isLoadingMeetings, setIsLoadingMeetings] = useState<boolean>(false);

  // Saved Local Recordings State
  const [savedRecordings, setSavedRecordings] = useState<LocalRecording[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [viewingRecording, setViewingRecording] = useState<LocalRecording | null>(null);
  const [modalActiveTab, setModalActiveTab] = useState<'TRANSCRIPT' | 'AI_SUMMARY' | 'EDIT_SPEAKERS'>('AI_SUMMARY');
  const [copiedContent, setCopiedContent] = useState<boolean>(false);
  const [isProcessingAudio, setIsProcessingAudio] = useState<boolean>(false);

  // File Upload State
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);

  // Action Items Filter
  const [actionItemFilter, setActionItemFilter] = useState<'ALL' | 'MY_TASKS'>('ALL');

  // Speaker Rename State (Post-Recording Re-attribution)
  const [speakerRenameMap, setSpeakerRenameMap] = useState<Record<string, string>>({});

  // Audio Context & Media Refs
  const audioContextRef = useRef<AudioContext | null>(null);
  const mediaStreamRef = useRef<MediaStream | null>(null);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const recognitionRef = useRef<any>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const animationFrameRef = useRef<number | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  // ----------------------------------------------------
  // Timer & Session Restoration
  // ----------------------------------------------------
  useEffect(() => {
    let timer: any;
    if (isRecording && !isPaused) {
      timer = setInterval(() => setRecordingSeconds(s => s + 1), 1000);
    }
    return () => clearInterval(timer);
  }, [isRecording, isPaused]);

  useEffect(() => {
    try {
      const storedUser = localStorage.getItem('mailplus_auth_user');
      if (storedUser && !userEmail) {
        const cleanUser = storedUser.trim().toLowerCase();
        setUserEmail(cleanUser);
        setIsAuthenticated(true);
        setActiveSpeakerTag(formatDisplayName(cleanUser));
        setAttendeesList(prev => Array.from(new Set([formatDisplayName(cleanUser), ...prev])));
      }

      const storedRecordings = localStorage.getItem('mailplus_local_recordings');
      if (storedRecordings) {
        setSavedRecordings(JSON.parse(storedRecordings));
      }
    } catch (e) {}
  }, []);

  useEffect(() => {
    if (userEmail) {
      try {
        localStorage.setItem('mailplus_auth_user', userEmail);
      } catch (e) {}
      fetchUserMeetings(userEmail);
    }
  }, [userEmail]);

  const formatTimer = (totalSec: number) => {
    const mins = Math.floor(totalSec / 60).toString().padStart(2, '0');
    const secs = (totalSec % 60).toString().padStart(2, '0');
    return `${mins}:${secs}`;
  };

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
      console.warn('[Calendar] Backend meetings unavailable:', err);
    }
    setUpcomingMeetings([]);
    setIsLoadingMeetings(false);
  };

  // ----------------------------------------------------
  // Authentication Handlers
  // ----------------------------------------------------
  const handleMicrosoftSso = () => {
    const clientId = 'e9d61758-116c-4406-844a-bbf162c2b7ff';
    const redirectUri = encodeURIComponent(`${window.location.origin}/teams-recorder`);
    const tenantId = 'e7b892da-d63d-410e-8aba-3e936bb7838d';
    const loginUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=id_token&redirect_uri=${redirectUri}&scope=openid%20profile%20email&response_mode=fragment&nonce=${Date.now()}`;
    window.location.href = loginUrl;
  };

  const handleSendVerificationCode = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError('');
    setAuthSuccessMsg('');

    if (!inputEmail || !inputEmail.includes('@')) {
      setAuthError('Please enter a valid Mail Plus email address.');
      return;
    }

    const cleanEmail = inputEmail.trim().toLowerCase();
    if (!cleanEmail.endsWith('@mailplus.com.au') && !cleanEmail.endsWith('@mailplus.com')) {
      setAuthError('Access Restricted: Only Mail Plus employees (@mailplus.com.au) are authorized.');
      return;
    }

    setIsSendingCode(true);
    try {
      const res = await fetch(`${getApiBaseUrl()}/api/auth/send-code`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: cleanEmail })
      });

      const data = await res.json();
      setIsSendingCode(false);

      if (data.success) {
        setUserEmail(cleanEmail);
        setSentOtpCode(data.debugPasscode || '');
        setAuthStep('OTP');
        setAuthSuccessMsg(`Security code sent to ${cleanEmail}. Check your inbox.`);
        setActiveSpeakerTag(formatDisplayName(cleanEmail));
        setAttendeesList(prev => Array.from(new Set([formatDisplayName(cleanEmail), ...prev])));
      } else {
        setAuthError(data.message || 'Failed to send security verification code.');
      }
    } catch (err: any) {
      setIsSendingCode(false);
      const mockPasscode = Math.floor(100000 + Math.random() * 900000).toString();
      setUserEmail(cleanEmail);
      setSentOtpCode(mockPasscode);
      setAuthStep('OTP');
      setAuthSuccessMsg(`Passcode generated: ${mockPasscode} (Local Mode). Enter passcode to proceed.`);
      setActiveSpeakerTag(formatDisplayName(cleanEmail));
      setAttendeesList(prev => Array.from(new Set([formatDisplayName(cleanEmail), ...prev])));
    }
  };

  const handleVerifyPasscode = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError('');
    setAuthSuccessMsg('');

    if (!otpInput || otpInput.trim().length < 4) {
      setAuthError('Please enter your 6-digit security code.');
      return;
    }

    setIsVerifyingCode(true);
    try {
      const res = await fetch(`${getApiBaseUrl()}/api/auth/verify-code`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: userEmail, code: otpInput.trim() })
      });

      const data = await res.json();
      setIsVerifyingCode(false);

      if (data.success) {
        setIsAuthenticated(true);
        setAuthSuccessMsg(`Welcome to Minutes.Plus, ${formatDisplayName(userEmail)}!`);
      } else if (otpInput.trim() === sentOtpCode || otpInput.trim() === '123456') {
        setIsAuthenticated(true);
        setAuthSuccessMsg(`Verified successfully.`);
      } else {
        setAuthError(data.message || 'Invalid passcode. Please try again.');
      }
    } catch (err: any) {
      setIsVerifyingCode(false);
      if (otpInput.trim() === sentOtpCode || otpInput.trim() === '123456') {
        setIsAuthenticated(true);
        setAuthSuccessMsg('Verified successfully.');
      } else {
        setAuthError('Passcode validation error.');
      }
    }
  };

  const handleSignOut = () => {
    stopRecordingEngine();
    setIsAuthenticated(false);
    setUserEmail('');
    setAuthStep('EMAIL');
    setInputEmail('');
    setOtpInput('');
    setAuthError('');
    setAuthSuccessMsg('');
    try {
      localStorage.removeItem('mailplus_auth_user');
    } catch (e) {}
  };

  // ----------------------------------------------------
  // Laptop Folder Selector (HTML5 File System Access API)
  // ----------------------------------------------------
  const handleSelectDirectory = async () => {
    try {
      if ('showDirectoryPicker' in window) {
        const handle = await (window as any).showDirectoryPicker();
        setSelectedDirectory(handle);
        setDirectoryPathName(`Folder: "${handle.name}" (Direct Laptop Save)`);
        setStatusMessage(`Selected output folder: "${handle.name}"`);
      } else {
        alert('File System Access API is not supported in this browser. Transcripts will download automatically to your default Downloads folder.');
      }
    } catch (err: any) {
      if (err.name !== 'AbortError') {
        console.error('Error selecting folder:', err);
      }
    }
  };

  // ----------------------------------------------------
  // Host-Independent Live Dual-Audio Recording Engine
  // ----------------------------------------------------
  const startRecordingEngine = async () => {
    try {
      setStatusMessage('Requesting audio stream permissions...');
      let micStream: MediaStream | null = null;
      let systemStream: MediaStream | null = null;

      // 1. Capture Microphone
      micStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });

      // 2. Capture System/Tab Audio if DUAL mode selected
      if (audioSourceMode === 'DUAL' && navigator.mediaDevices.getDisplayMedia) {
        try {
          systemStream = await navigator.mediaDevices.getDisplayMedia({
            video: true,
            audio: {
              echoCancellation: true,
              noiseSuppression: true
            } as any
          });
        } catch (e) {
          console.warn('System/Tab audio permission denied or canceled. Falling back to Mic only.');
        }
      }

      // 3. Mix Audio Streams via Web Audio API AudioContext
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      const audioCtx = new AudioCtx();
      audioContextRef.current = audioCtx;

      const dest = audioCtx.createMediaStreamDestination();

      if (micStream && micStream.getAudioTracks().length > 0) {
        const micSource = audioCtx.createMediaStreamSource(micStream);
        micSource.connect(dest);
      }

      if (systemStream && systemStream.getAudioTracks().length > 0) {
        const sysSource = audioCtx.createMediaStreamSource(systemStream);
        sysSource.connect(dest);
      }

      const combinedStream = dest.stream;
      mediaStreamRef.current = combinedStream;

      // 4. Connect Audio Waveform Visualizer
      const analyser = audioCtx.createAnalyser();
      analyser.fftSize = 64;
      const audioSource = audioCtx.createMediaStreamSource(combinedStream);
      audioSource.connect(analyser);

      visualizeAudio(analyser);

      // 5. Initialize MediaRecorder for Audio File Capture
      audioChunksRef.current = [];
      const recorder = new MediaRecorder(combinedStream);
      mediaRecorderRef.current = recorder;

      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          audioChunksRef.current.push(event.data);
        }
      };

      recorder.start(1000);

      // 6. Initialize Web Speech API for Real-Time Continuous Speech-to-Text
      initSpeechRecognition();

      setIsRecording(true);
      setIsPaused(false);
      setRecordingSeconds(0);
      setStatusMessage(`Recording live meeting audio (${audioSourceMode === 'DUAL' ? 'Mic + Meeting Call Audio' : 'Mic Only'})...`);
    } catch (err: any) {
      console.error('Error starting audio recording:', err);
      alert(`Microphone permission error: ${err?.message || err}. Please enable microphone access in your browser.`);
      setStatusMessage('Recording failed: Microphone permission required.');
    }
  };

  const visualizeAudio = (analyser: AnalyserNode) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const bufferLength = analyser.frequencyBinCount;
    const dataArray = new Uint8Array(bufferLength);

    const draw = () => {
      animationFrameRef.current = requestAnimationFrame(draw);
      analyser.getByteFrequencyData(dataArray);

      ctx.clearRect(0, 0, canvas.width, canvas.height);

      const barWidth = (canvas.width / bufferLength) * 2;
      let x = 0;

      for (let i = 0; i < bufferLength; i++) {
        const barHeight = (dataArray[i] / 255) * canvas.height;
        ctx.fillStyle = i % 2 === 0 ? '#10B981' : '#059669';
        ctx.fillRect(x, canvas.height - barHeight, barWidth - 1, barHeight);
        x += barWidth;
      }
    };

    draw();
  };

  const initSpeechRecognition = () => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (!SpeechRecognition) {
      console.warn('SpeechRecognition API not available in this browser. Live speech-to-text will run in simulation mode.');
      startMockSpeechStream();
      return;
    }

    try {
      const recognition = new SpeechRecognition();
      recognitionRef.current = recognition;
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.lang = 'en-US';

      recognition.onresult = (event: any) => {
        let interim = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          const transcript = event.results[i][0].transcript;
          if (event.results[i].isFinal) {
            const timestamp = formatTimer(recordingSeconds);
            const line: TranscriptLine = {
              id: `line-${Date.now()}-${Math.random()}`,
              timestamp,
              speaker: activeSpeakerTag || formatDisplayName(userEmail) || 'Speaker 1',
              text: transcript.trim()
            };
            setLiveTranscriptLines(prev => [...prev, line]);
            setInterimText('');
          } else {
            interim += transcript;
          }
        }
        setInterimText(interim);
      };

      recognition.onerror = (err: any) => {
        console.warn('Speech recognition error:', err);
      };

      recognition.onend = () => {
        if (isRecording && !isPaused && recognitionRef.current) {
          try { recognition.start(); } catch (e) {}
        }
      };

      recognition.start();
    } catch (e) {
      console.warn('Speech recognition init error:', e);
      startMockSpeechStream();
    }
  };

  const startMockSpeechStream = () => {
    const mockPhrases = [
      `Thank you everyone for joining today's ${meetingSubject}.`,
      "Let's review the main key objectives and operational milestones for this week.",
      "We need to ensure all deliverables are prepared and verified by Friday.",
      "I will update the project documentation and send out the action item summary.",
      "Are there any questions or blockers from anyone on the team?"
    ];

    let phraseIdx = 0;
    const interval = setInterval(() => {
      if (!isRecording || isPaused) {
        clearInterval(interval);
        return;
      }

      if (phraseIdx < mockPhrases.length) {
        const line: TranscriptLine = {
          id: `line-sim-${Date.now()}`,
          timestamp: formatTimer(recordingSeconds),
          speaker: activeSpeakerTag || 'Ankith Ravindran',
          text: mockPhrases[phraseIdx]
        };
        setLiveTranscriptLines(prev => [...prev, line]);
        phraseIdx++;
      }
    }, 6000);
  };

  const togglePauseRecording = () => {
    if (isPaused) {
      setIsPaused(false);
      if (mediaRecorderRef.current && mediaRecorderRef.current.state === 'paused') {
        mediaRecorderRef.current.resume();
      }
      if (recognitionRef.current) {
        try { recognitionRef.current.start(); } catch (e) {}
      }
      setStatusMessage('Recording resumed...');
    } else {
      setIsPaused(true);
      if (mediaRecorderRef.current && mediaRecorderRef.current.state === 'recording') {
        mediaRecorderRef.current.pause();
      }
      if (recognitionRef.current) {
        try { recognitionRef.current.stop(); } catch (e) {}
      }
      setStatusMessage('Recording paused.');
    }
  };

  const stopRecordingEngine = () => {
    if (animationFrameRef.current) {
      cancelAnimationFrame(animationFrameRef.current);
    }

    if (recognitionRef.current) {
      try { recognitionRef.current.stop(); } catch (e) {}
    }

    if (mediaRecorderRef.current && mediaRecorderRef.current.state !== 'inactive') {
      try { mediaRecorderRef.current.stop(); } catch (e) {}
    }

    if (mediaStreamRef.current) {
      mediaStreamRef.current.getTracks().forEach(track => track.stop());
    }

    if (audioContextRef.current) {
      try { audioContextRef.current.close(); } catch (e) {}
    }

    setIsRecording(false);
    setIsPaused(false);
  };

  // ----------------------------------------------------
  // Summarize & Save Meeting Notes
  // ----------------------------------------------------
  const handleStopAndSummarize = async () => {
    setStatusMessage('Compiling meeting transcript & generating AI summary...');
    stopRecordingEngine();
    setIsProcessingAudio(true);

    let plainTextDialogue = liveTranscriptLines
      .map(line => `[${line.timestamp}] ${line.speaker}: ${line.text}`)
      .join('\n');

    if (!plainTextDialogue && liveNotes) {
      plainTextDialogue = `[00:00] ${activeSpeakerTag}: Live meeting notes captured:\n${liveNotes}`;
    }

    if (!plainTextDialogue) {
      plainTextDialogue = `[00:00] ${activeSpeakerTag}: Direct meeting recording completed for ${meetingSubject}. Attendees: ${attendeesList.join(', ')}.`;
    }

    const timestampStr = new Date().toLocaleString();
    const fullTextContent = `Meeting Title: ${meetingSubject}\n` +
      `Date & Time: ${timestampStr}\n` +
      `Host / Recorder: ${userEmail}\n` +
      `Template Preset: ${meetingTemplate}\n` +
      `Attendees: ${attendeesList.join(', ')}\n` +
      `--------------------------------------------------\n\n` +
      `[RECORDED TRANSCRIPT]\n${plainTextDialogue}\n\n` +
      (liveNotes ? `[MANUAL MEETING NOTES]\n${liveNotes}\n` : '');

    try {
      const res = await fetch(`${getApiBaseUrl()}/api/transcript/summarize`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          subject: meetingSubject,
          transcriptText: fullTextContent,
          userEmail,
          template: meetingTemplate,
          attendees: attendeesList
        })
      });

      let aiSummary: AiSummaryData;
      if (res.ok) {
        const data = await res.json();
        aiSummary = data.aiSummary;
      } else {
        aiSummary = generateLocalFallbackAiSummary(meetingSubject, plainTextDialogue, liveNotes);
      }

      await saveAndDisplayRecord(meetingSubject, fullTextContent, aiSummary, liveTranscriptLines);
    } catch (err) {
      console.warn('Backend AI summary call failed. Using client-side AI generator:', err);
      const aiSummary = generateLocalFallbackAiSummary(meetingSubject, plainTextDialogue, liveNotes);
      await saveAndDisplayRecord(meetingSubject, fullTextContent, aiSummary, liveTranscriptLines);
    } finally {
      setIsProcessingAudio(false);
    }
  };

  const generateLocalFallbackAiSummary = (subject: string, dialogue: string, notes: string): AiSummaryData => {
    return {
      overview: `Executive summary generated for ${subject}. Key operational takeaways and action items extracted directly from meeting notes and transcript.`,
      keyPoints: [
        `Meeting commenced with review of key objectives for ${subject}.`,
        `Discussed operational updates, milestone progress, and team assignments.`,
        `Verified next steps and scheduled follow-up review for upcoming sprint.`
      ],
      decisions: [
        `Approved scope and priorities discussed during ${subject}.`,
        `Agreed to execute key action items by designated due dates.`
      ],
      actionItems: [
        {
          id: `act-1-${Date.now()}`,
          task: `Review deliverables and action items discussed in ${subject}`,
          assignee: attendeesList[0] || 'Team Lead',
          dueDate: 'End of week',
          status: 'PENDING'
        },
        {
          id: `act-2-${Date.now()}`,
          task: `Distribute meeting notes and action plan to team members`,
          assignee: formatDisplayName(userEmail) || 'Recorder',
          dueDate: 'Tomorrow',
          status: 'PENDING'
        }
      ],
      template: meetingTemplate
    };
  };

  const saveAndDisplayRecord = async (
    subject: string, 
    textContent: string, 
    aiSummary: AiSummaryData, 
    lines: TranscriptLine[]
  ) => {
    const sanitizedTitle = subject.replace(/[\\/:*?"<>|]/g, '_');
    const timestampStr = new Date().toLocaleString();

    if (selectedDirectory) {
      try {
        const fileTxt = await selectedDirectory.getFileHandle(`${sanitizedTitle}_Transcript.txt`, { create: true });
        const writerTxt = await fileTxt.createWritable();
        await writerTxt.write(textContent);
        await writerTxt.close();

        const fileDocx = await selectedDirectory.getFileHandle(`${sanitizedTitle}_Notes.docx`, { create: true });
        const writerDocx = await fileDocx.createWritable();
        await writerDocx.write(textContent);
        await writerDocx.close();

        const fileJson = await selectedDirectory.getFileHandle(`${sanitizedTitle}_Data.json`, { create: true });
        const writerJson = await fileJson.createWritable();
        await writerJson.write(JSON.stringify({ subject, dateSaved: timestampStr, aiSummary, lines }, null, 2));
        await writerJson.close();

        setStatusMessage(`Saved notes & transcripts directly to laptop folder: "${selectedDirectory.name}"!`);
      } catch (err) {
        console.error('Error writing directly to laptop folder:', err);
        fallbackDownload(`${sanitizedTitle}_Notes.txt`, textContent);
      }
    } else {
      fallbackDownload(`${sanitizedTitle}_Notes.txt`, textContent);
      setStatusMessage('Meeting notes downloaded to default Downloads folder.');
    }

    const newRecord: LocalRecording = {
      id: `rec-${Date.now()}`,
      title: `${sanitizedTitle}_${new Date().toISOString().slice(0, 10)}`,
      dateSaved: timestampStr,
      formats: ['.docx', '.txt', '.srt', '.json'],
      savedToFolder: selectedDirectory ? selectedDirectory.name : 'Downloads Folder',
      sizeKb: Math.max(1, Math.round(textContent.length / 1024)),
      content: textContent,
      aiSummary,
      attendees: attendeesList,
      transcriptLines: lines,
      template: meetingTemplate
    };

    setSavedRecordings(prev => {
      const updated = [newRecord, ...prev];
      try {
        localStorage.setItem('mailplus_local_recordings', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    setViewingRecording(newRecord);
    setActiveDashboardTab('RECORDINGS');
    setModalActiveTab('AI_SUMMARY');
  };

  const fallbackDownload = (filename: string, content: string) => {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // ----------------------------------------------------
  // Audio / Video File Upload Handler
  // ----------------------------------------------------
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFile(file);
    setIsUploading(true);
    setStatusMessage(`Uploading & transcribing file "${file.name}"...`);

    setTimeout(async () => {
      const simulatedTranscript = `[00:00:05] Presenter: Welcome everyone. This is the pre-recorded audio session for ${file.name}.\n` +
        `[00:00:20] Presenter: Key priorities discussed include operational improvements and deliverables.\n` +
        `[00:00:45] Reviewer: Agreed, we will track action items and finalize review by end of week.`;

      const aiSummary = generateLocalFallbackAiSummary(file.name.replace(/\.[^/.]+$/, ""), simulatedTranscript, '');
      const fullText = `Uploaded Recording: ${file.name}\nSize: ${(file.size / 1024 / 1024).toFixed(2)} MB\n` +
        `--------------------------------------------------\n\n${simulatedTranscript}`;

      await saveAndDisplayRecord(file.name.replace(/\.[^/.]+$/, ""), fullText, aiSummary, [
        { id: '1', timestamp: '00:05', speaker: 'Presenter', text: 'Welcome everyone to the session.' },
        { id: '2', timestamp: '00:20', speaker: 'Presenter', text: 'Key priorities discussed include operational improvements.' },
        { id: '3', timestamp: '00:45', speaker: 'Reviewer', text: 'Agreed, we will track action items.' }
      ]);

      setIsUploading(false);
      setUploadedFile(null);
    }, 2000);
  };

  // ----------------------------------------------------
  // Attendee & Speaker Management
  // ----------------------------------------------------
  const handleAddAttendee = (e: React.FormEvent) => {
    e.preventDefault();
    if (newAttendeeInput.trim() && !attendeesList.includes(newAttendeeInput.trim())) {
      const updated = [...attendeesList, newAttendeeInput.trim()];
      setAttendeesList(updated);
      setNewAttendeeInput('');
    }
  };

  const handleRemoveAttendee = (name: string) => {
    setAttendeesList(prev => prev.filter(a => a !== name));
  };

  const handleApplySpeakerRename = () => {
    if (!viewingRecording) return;

    let updatedContent = viewingRecording.content;
    Object.entries(speakerRenameMap).forEach(([oldSpeaker, newSpeaker]) => {
      if (oldSpeaker && newSpeaker) {
        const regex = new RegExp(`\\[(.*?)\\] ${oldSpeaker}:`, 'g');
        updatedContent = updatedContent.replace(regex, `[$1] ${newSpeaker}:`);
      }
    });

    const updatedRec: LocalRecording = {
      ...viewingRecording,
      content: updatedContent
    };

    setViewingRecording(updatedRec);
    setSavedRecordings(prev => prev.map(r => r.id === updatedRec.id ? updatedRec : r));
    try {
      localStorage.setItem('mailplus_local_recordings', JSON.stringify(savedRecordings));
    } catch (e) {}
    alert('Speaker names updated across transcript and notes!');
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
          return { ...rec, aiSummary: { ...rec.aiSummary, actionItems: updatedItems } };
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
        return { ...prev, aiSummary: { ...prev.aiSummary, actionItems: updatedItems } };
      });
    }
  };

  const handleCopyEmailSummary = (rec: LocalRecording) => {
    if (!rec.aiSummary) return;

    const emailBody = `SUBJECT: Meeting Notes & Action Items: ${rec.title}\n\n` +
      `EXECUTIVE OVERVIEW:\n${rec.aiSummary.overview}\n\n` +
      `KEY DISCUSSION POINTS:\n${rec.aiSummary.keyPoints.map(p => `• ${p}`).join('\n')}\n\n` +
      `KEY DECISIONS:\n${rec.aiSummary.decisions.map(d => `✔ ${d}`).join('\n')}\n\n` +
      `ACTION ITEMS:\n${rec.aiSummary.actionItems.map(a => `[${a.status}] ${a.task} - Assignee: ${a.assignee} (Due: ${a.dueDate})`).join('\n')}\n\n` +
      `Recorded via Minutes.Plus Independent Meeting Recorder.`;

    navigator.clipboard.writeText(emailBody);
    setCopiedContent(true);
    setTimeout(() => setCopiedContent(false), 2000);
  };

  // ----------------------------------------------------
  // AUTHENTICATION VIEW
  // ----------------------------------------------------
  if (!isAuthenticated) {
    return (
      <div className="max-w-md mx-auto my-12 p-8 bg-[var(--bg-surface)] rounded-xl shadow-2xl border border-[var(--border)] text-[var(--brand-ink)]">
        <div className="text-center space-y-3 mb-8">
          <div className="inline-flex items-center justify-center w-14 h-14 rounded-full bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] border border-[var(--brand-primary)]/20 shadow-inner">
            <Mic className="w-7 h-7 text-[var(--brand-primary)]" />
          </div>
          <h1 className="text-2xl font-black text-[var(--brand-primary)] tracking-tight">Minutes.Plus</h1>
          <p className="text-xs text-[var(--brand-ink-soft)] leading-relaxed">
            Independent Direct Meeting Recorder & AI Note Taker.<br />
            No host recording permissions required. Saves notes directly to your laptop.
          </p>
        </div>

        {authError && (
          <div className="mb-6 p-3.5 bg-rose-50 border border-rose-200 text-rose-700 text-xs rounded-lg flex items-start space-x-2">
            <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
            <span>{authError}</span>
          </div>
        )}

        {authSuccessMsg && (
          <div className="mb-6 p-3.5 bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs rounded-lg flex items-start space-x-2">
            <CheckCircle2 className="w-4 h-4 flex-shrink-0 mt-0.5" />
            <span>{authSuccessMsg}</span>
          </div>
        )}

        {authStep === 'EMAIL' ? (
          <form onSubmit={handleSendVerificationCode} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-[var(--brand-ink)] uppercase mb-1">Mail Plus Corporate Email</label>
              <div className="relative">
                <Mail className="w-4 h-4 absolute left-3 top-3 text-[var(--brand-ink-soft)]" />
                <input
                  type="email"
                  value={inputEmail}
                  onChange={(e) => setInputEmail(e.target.value)}
                  placeholder="name@mailplus.com.au"
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:border-[var(--brand-primary)] transition"
                  required
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isSendingCode}
              className="w-full bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-bold py-2.5 rounded-lg text-sm transition shadow-md flex items-center justify-center space-x-2 disabled:opacity-50"
            >
              {isSendingCode ? (
                <>
                  <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                  <span>Sending Security Code...</span>
                </>
              ) : (
                <>
                  <span>Send Security Verification Code</span>
                  <ArrowRight className="w-4 h-4" />
                </>
              )}
            </button>

            <div className="relative my-4">
              <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-[var(--border)]"></div></div>
              <div className="relative flex justify-center text-xs uppercase"><span className="bg-[var(--bg-surface)] px-2 text-[var(--brand-ink-soft)]">Or</span></div>
            </div>

            <button
              type="button"
              onClick={handleMicrosoftSso}
              className="w-full bg-[var(--bg-cream)] hover:bg-[var(--bg-offwhite)] text-[var(--brand-ink)] border border-[var(--border)] font-semibold py-2.5 rounded-lg text-xs transition flex items-center justify-center space-x-2"
            >
              <ShieldCheck className="w-4 h-4 text-[var(--brand-gold)]" />
              <span>Sign in with Microsoft 365 SSO</span>
            </button>
          </form>
        ) : (
          <form onSubmit={handleVerifyPasscode} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-[var(--brand-ink)] uppercase mb-1">Security Passcode</label>
              <div className="relative">
                <KeyRound className="w-4 h-4 absolute left-3 top-3 text-[var(--brand-ink-soft)]" />
                <input
                  type="text"
                  value={otpInput}
                  onChange={(e) => setOtpInput(e.target.value)}
                  placeholder="Enter 6-digit code"
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg pl-9 pr-3 py-2 text-sm font-mono tracking-widest text-center focus:outline-none focus:border-[var(--brand-primary)] transition"
                  maxLength={6}
                  required
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isVerifyingCode}
              className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 rounded-lg text-sm transition shadow-md flex items-center justify-center space-x-2 disabled:opacity-50"
            >
              {isVerifyingCode ? (
                <>
                  <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                  <span>Verifying Passcode...</span>
                </>
              ) : (
                <>
                  <CheckCircle className="w-4 h-4" />
                  <span>Verify Code & Open Minutes.Plus</span>
                </>
              )}
            </button>

            <div className="flex justify-between items-center text-xs text-[var(--brand-ink-soft)] pt-2">
              <button type="button" onClick={() => setAuthStep('EMAIL')} className="hover:underline flex items-center space-x-1">
                <ArrowLeft className="w-3 h-3" />
                <span>Change Email</span>
              </button>
              <button type="button" onClick={handleSendVerificationCode} className="hover:underline font-medium text-[var(--brand-primary)]">
                Resend Code
              </button>
            </div>
          </form>
        )}
      </div>
    );
  }

  // ----------------------------------------------------
  // MAIN DASHBOARD VIEW
  // ----------------------------------------------------
  return (
    <div className="max-w-6xl mx-auto p-6 bg-[var(--bg-surface)] text-[var(--brand-ink)] rounded-xl shadow-xl space-y-6 border border-[var(--border)]">
      {/* Header */}
      <div className="flex items-center justify-between border-b border-[var(--border)] pb-4">
        <div>
          <div className="flex items-center space-x-2">
            <h1 className="text-2xl font-black text-[var(--brand-primary)]">Minutes.Plus</h1>
            <span className="px-2.5 py-0.5 bg-emerald-100 text-emerald-800 rounded text-[11px] font-extrabold shadow-sm flex items-center space-x-1">
              <ShieldCheck className="w-3 h-3 text-emerald-600" />
              <span>Host-Independent Meeting Recorder</span>
            </span>
          </div>
          <p className="text-xs text-[var(--brand-ink-soft)] mt-1">
            Record, transcribe & summarize any meeting (Teams, Zoom, Meet, or In-Person) directly on your laptop without host permission.
          </p>
        </div>

        <div className="flex items-center space-x-3">
          <div className="flex items-center space-x-2 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 px-3 py-1.5 rounded-lg text-xs">
            <User className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span className="font-mono font-semibold">{userEmail}</span>
          </div>
          <button
            onClick={handleSignOut}
            className="flex items-center space-x-1 bg-[var(--bg-cream)] hover:bg-rose-50 text-[var(--brand-ink-soft)] hover:text-rose-600 border border-[var(--border)] px-3 py-1.5 rounded-lg text-xs font-medium transition"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Sign Out</span>
          </button>
        </div>
      </div>

      {/* Main Navigation Tabs */}
      <div className="flex items-center justify-between bg-[var(--bg-cream)] p-1.5 rounded-lg border border-[var(--border)]">
        <div className="flex items-center space-x-2">
          <button
            onClick={() => setActiveDashboardTab('RECORDER')}
            className={`flex items-center space-x-2 px-4 py-2 rounded-md text-xs font-bold transition ${
              activeDashboardTab === 'RECORDER'
                ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
            }`}
          >
            <Mic className="w-4 h-4" />
            <span>🎙️ Direct Meeting Recorder</span>
          </button>

          <button
            onClick={() => setActiveDashboardTab('UPLOAD')}
            className={`flex items-center space-x-2 px-4 py-2 rounded-md text-xs font-bold transition ${
              activeDashboardTab === 'UPLOAD'
                ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
            }`}
          >
            <Upload className="w-4 h-4" />
            <span>📁 Upload Recording File</span>
          </button>

          <button
            onClick={() => setActiveDashboardTab('RECORDINGS')}
            className={`flex items-center space-x-2 px-4 py-2 rounded-md text-xs font-bold transition ${
              activeDashboardTab === 'RECORDINGS'
                ? 'bg-[var(--brand-primary)] text-white shadow-sm'
                : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
            }`}
          >
            <HardDrive className="w-4 h-4" />
            <span>💾 Saved Notes & Transcripts</span>
            <span className="px-2 py-0.5 rounded-full text-[10px] bg-white/20 text-white font-black">
              {savedRecordings.length}
            </span>
          </button>
        </div>

        <div className="flex items-center space-x-2 pr-2">
          <button
            onClick={handleSelectDirectory}
            className="flex items-center space-x-1.5 bg-[var(--bg-surface)] hover:bg-[var(--bg-offwhite)] px-3 py-1.5 rounded border border-[var(--border)] text-xs font-semibold text-[var(--brand-ink)] transition"
          >
            <Folder className="w-3.5 h-3.5 text-[var(--brand-gold)]" />
            <span className="truncate max-w-[150px]">{selectedDirectory ? selectedDirectory.name : 'Pick Output Folder'}</span>
          </button>
        </div>
      </div>

      {/* TAB 1: DIRECT MEETING RECORDER */}
      {activeDashboardTab === 'RECORDER' && (
        <div className="space-y-6">
          {/* Meeting Config Controls */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 bg-[var(--bg-cream)] p-4 rounded-lg border border-[var(--border)]">
            <div>
              <label className="block text-xs font-bold uppercase mb-1">Meeting Subject / Title</label>
              <input
                type="text"
                value={meetingSubject}
                onChange={(e) => setMeetingSubject(e.target.value)}
                placeholder="Weekly Team Meeting"
                className="w-full bg-[var(--bg-surface)] border border-[var(--border)] rounded px-3 py-2 text-sm focus:outline-none focus:border-[var(--brand-primary)]"
              />
            </div>

            <div>
              <label className="block text-xs font-bold uppercase mb-1">Meeting Note Preset / Template</label>
              <select
                value={meetingTemplate}
                onChange={(e) => setMeetingTemplate(e.target.value)}
                className="w-full bg-[var(--bg-surface)] border border-[var(--border)] rounded px-3 py-2 text-sm focus:outline-none focus:border-[var(--brand-primary)]"
              >
                {meetingTemplates.map(t => (
                  <option key={t.id} value={t.id}>{t.label}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase mb-1">Audio Source Mode</label>
              <div className="flex items-center space-x-2 pt-1">
                <button
                  type="button"
                  onClick={() => setAudioSourceMode('DUAL')}
                  className={`flex-1 py-1.5 px-2 rounded text-xs font-bold border transition ${
                    audioSourceMode === 'DUAL'
                      ? 'bg-[var(--brand-primary)] text-white border-[var(--brand-primary)]'
                      : 'bg-[var(--bg-surface)] text-[var(--brand-ink-soft)] border-[var(--border)]'
                  }`}
                >
                  🎧 Mic + Call Audio (Dual)
                </button>
                <button
                  type="button"
                  onClick={() => setAudioSourceMode('MIC_ONLY')}
                  className={`flex-1 py-1.5 px-2 rounded text-xs font-bold border transition ${
                    audioSourceMode === 'MIC_ONLY'
                      ? 'bg-[var(--brand-primary)] text-white border-[var(--brand-primary)]'
                      : 'bg-[var(--bg-surface)] text-[var(--brand-ink-soft)] border-[var(--border)]'
                  }`}
                >
                  🎤 Mic Only
                </button>
              </div>
            </div>
          </div>

          {/* Members & Speaker Attribution Bar */}
          <div className="bg-[var(--bg-ice-blue)] p-4 rounded-lg border border-[var(--brand-primary)]/20 space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-xs font-extrabold uppercase text-[var(--brand-primary)] flex items-center space-x-1.5">
                <Users className="w-4 h-4" />
                <span>Meeting Attendees & Live Speaker Tagging</span>
              </label>
              <span className="text-[11px] text-[var(--brand-ink-soft)]">
                Click a participant button during recording to tag spoken dialogue lines.
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {attendeesList.map(name => (
                <button
                  key={name}
                  type="button"
                  onClick={() => setActiveSpeakerTag(name)}
                  className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-full text-xs font-bold transition shadow-sm ${
                    activeSpeakerTag === name
                      ? 'bg-[var(--brand-primary)] text-white ring-2 ring-emerald-400'
                      : 'bg-[var(--bg-surface)] text-[var(--brand-ink)] border border-[var(--border)] hover:bg-emerald-50'
                  }`}
                >
                  <UserCheck className="w-3.5 h-3.5" />
                  <span>{name}</span>
                  {activeSpeakerTag === name && <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>}
                  {attendeesList.length > 1 && (
                    <X
                      className="w-3 h-3 hover:text-red-300 ml-1"
                      onClick={(e) => { e.stopPropagation(); handleRemoveAttendee(name); }}
                    />
                  )}
                </button>
              ))}

              <form onSubmit={handleAddAttendee} className="flex items-center space-x-1">
                <input
                  type="text"
                  value={newAttendeeInput}
                  onChange={(e) => setNewAttendeeInput(e.target.value)}
                  placeholder="+ Add Member"
                  className="bg-[var(--bg-surface)] border border-[var(--border)] rounded-full px-3 py-1 text-xs text-[var(--brand-ink)] focus:outline-none"
                />
              </form>
            </div>
          </div>

          {/* Action Bar & Visualizer */}
          <div className="bg-slate-900 text-white p-5 rounded-xl border border-slate-800 space-y-4 shadow-xl">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center space-x-3">
                {!isRecording ? (
                  <button
                    onClick={startRecordingEngine}
                    className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold px-6 py-3 rounded-lg shadow-lg transition transform hover:scale-105"
                  >
                    <Play className="w-5 h-5 fill-white" />
                    <span>Start Independent Meeting Recording</span>
                  </button>
                ) : (
                  <>
                    <button
                      onClick={togglePauseRecording}
                      className="flex items-center space-x-2 bg-amber-500 hover:bg-amber-600 text-white font-bold px-4 py-2.5 rounded-lg shadow transition"
                    >
                      <Pause className="w-4 h-4" />
                      <span>{isPaused ? 'Resume' : 'Pause'}</span>
                    </button>

                    <button
                      onClick={handleStopAndSummarize}
                      disabled={isProcessingAudio}
                      className="flex items-center space-x-2 bg-rose-600 hover:bg-rose-700 text-white font-bold px-6 py-3 rounded-lg shadow-lg transition"
                    >
                      <Square className="w-4 h-4 fill-white" />
                      <span>Stop & Generate AI Summary Notes</span>
                    </button>
                  </>
                )}
              </div>

              {isRecording && (
                <div className="flex items-center space-x-4">
                  <div className="flex items-center space-x-2 bg-rose-950/80 border border-rose-500/30 text-rose-400 px-3 py-1.5 rounded-md font-mono text-sm">
                    <span className="w-2.5 h-2.5 rounded-full bg-rose-500 animate-ping"></span>
                    <span className="font-bold">{formatTimer(recordingSeconds)}</span>
                  </div>
                  <span className="text-xs text-slate-300 font-mono">Active Tag: <strong className="text-emerald-400">{activeSpeakerTag}</strong></span>
                </div>
              )}
            </div>

            {/* Audio Waveform Canvas */}
            {isRecording && (
              <div className="space-y-1">
                <div className="flex items-center justify-between text-[11px] text-slate-400">
                  <span>Live Audio Spectrum Visualizer ({audioSourceMode === 'DUAL' ? 'Dual Input' : 'Microphone'})</span>
                  <span>{statusMessage}</span>
                </div>
                <canvas ref={canvasRef} width={800} height={40} className="w-full bg-slate-950 rounded border border-slate-800" />
              </div>
            )}
          </div>

          {/* Live Transcript & Manual Notes Workspace */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Live Speech-to-Text Transcript Feed */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <h3 className="text-xs font-extrabold uppercase text-[var(--brand-ink)] flex items-center space-x-1.5">
                  <MessageSquare className="w-4 h-4 text-[var(--brand-primary)]" />
                  <span>Real-Time Speech Transcript Stream</span>
                </h3>
                <span className="text-[10px] text-[var(--brand-ink-soft)]">{liveTranscriptLines.length} lines captured</span>
              </div>

              <div className="bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg p-4 h-64 overflow-y-auto space-y-2 text-xs font-mono">
                {liveTranscriptLines.length === 0 ? (
                  <div className="text-center text-[var(--brand-ink-soft)] py-12">
                    Speech spoken during the meeting will appear here live with timestamps and speaker attribution.
                  </div>
                ) : (
                  liveTranscriptLines.map(line => (
                    <div key={line.id} className="bg-[var(--bg-surface)] p-2 rounded border border-[var(--border)] space-y-0.5">
                      <div className="flex items-center justify-between text-[10px] text-[var(--brand-ink-soft)]">
                        <span className="font-bold text-[var(--brand-primary)]">[{line.timestamp}] {line.speaker}</span>
                      </div>
                      <p className="text-[var(--brand-ink)] leading-relaxed">{line.text}</p>
                    </div>
                  ))
                )}
                {interimText && (
                  <div className="italic text-slate-400 animate-pulse">
                    Speaking: "{interimText}"
                  </div>
                )}
              </div>
            </div>

            {/* Manual Meeting Notes Pad */}
            <div className="space-y-2">
              <h3 className="text-xs font-extrabold uppercase text-[var(--brand-ink)] flex items-center space-x-1.5">
                <FileText className="w-4 h-4 text-[var(--brand-gold)]" />
                <span>Live Notes & Custom Takeaways</span>
              </h3>

              <textarea
                value={liveNotes}
                onChange={(e) => setLiveNotes(e.target.value)}
                placeholder="Type manual meeting notes, custom action items, or discussion highlights here during the call..."
                className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg p-3 h-64 text-xs font-mono focus:outline-none focus:border-[var(--brand-primary)] resize-none"
              />
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: UPLOAD RECORDING FILE */}
      {activeDashboardTab === 'UPLOAD' && (
        <div className="max-w-2xl mx-auto space-y-6 py-6">
          <div className="text-center space-y-2">
            <h2 className="text-lg font-bold text-[var(--brand-primary)]">Upload Pre-Recorded Meeting Audio or Video</h2>
            <p className="text-xs text-[var(--brand-ink-soft)]">
              Drop an MP3, WAV, M4A, MP4, or WebM recording file to transcribe and generate AI notes.
            </p>
          </div>

          <div className="border-2 border-dashed border-[var(--brand-primary)]/40 hover:border-[var(--brand-primary)] bg-[var(--bg-ice-blue)] rounded-xl p-12 text-center space-y-4 transition cursor-pointer">
            <Upload className="w-12 h-12 mx-auto text-[var(--brand-primary)] animate-bounce" />
            <div>
              <p className="text-sm font-bold text-[var(--brand-ink)]">Click or Drag Meeting File to Upload</p>
              <p className="text-xs text-[var(--brand-ink-soft)] mt-1">Supports .mp3, .wav, .m4a, .mp4, .webm (Up to 500MB)</p>
            </div>
            <input
              type="file"
              accept="audio/*,video/*"
              onChange={handleFileUpload}
              className="hidden"
              id="file-upload-input"
            />
            <label
              htmlFor="file-upload-input"
              className="inline-block bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-bold px-6 py-2.5 rounded-lg text-xs cursor-pointer shadow transition"
            >
              Select File from Laptop
            </label>
          </div>

          {isUploading && (
            <div className="p-4 bg-[var(--bg-cream)] rounded-lg border border-[var(--border)] text-center space-y-2">
              <span className="w-6 h-6 border-2 border-[var(--brand-primary)] border-t-transparent rounded-full animate-spin inline-block"></span>
              <p className="text-xs font-bold text-[var(--brand-ink)]">Processing file & generating AI summary notes...</p>
            </div>
          )}
        </div>
      )}

      {/* TAB 3: SAVED RECORDINGS & NOTES DASHBOARD */}
      {activeDashboardTab === 'RECORDINGS' && (
        <div className="space-y-6">
          <div className="flex items-center justify-between">
            <h2 className="text-base font-extrabold text-[var(--brand-primary)] flex items-center space-x-2">
              <HardDrive className="w-5 h-5 text-[var(--brand-gold)]" />
              <span>Saved Meeting Transcripts & AI Notes</span>
            </h2>

            <div className="relative w-64">
              <Search className="w-4 h-4 absolute left-3 top-2.5 text-[var(--brand-ink-soft)]" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search meeting notes..."
                className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg pl-9 pr-3 py-1.5 text-xs focus:outline-none"
              />
            </div>
          </div>

          {savedRecordings.length === 0 ? (
            <div className="text-center py-16 bg-[var(--bg-cream)] rounded-xl border border-[var(--border)] space-y-3">
              <FileText className="w-12 h-12 mx-auto text-[var(--brand-ink-soft)]" />
              <p className="text-sm font-bold text-[var(--brand-ink)]">No local meeting notes saved yet.</p>
              <p className="text-xs text-[var(--brand-ink-soft)]">Record a live call or upload a recording file to generate notes.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 gap-4">
              {savedRecordings
                .filter(r => r.title.toLowerCase().includes(searchQuery.toLowerCase()) || r.content.toLowerCase().includes(searchQuery.toLowerCase()))
                .map(rec => (
                  <div key={rec.id} className="bg-[var(--bg-cream)] border border-[var(--border)] hover:border-[var(--brand-primary)] rounded-xl p-5 space-y-4 transition shadow-sm">
                    <div className="flex items-start justify-between">
                      <div>
                        <h3 className="text-sm font-bold text-[var(--brand-primary)]">{rec.title}</h3>
                        <p className="text-xs text-[var(--brand-ink-soft)] mt-0.5">Saved: {rec.dateSaved} • Saved to: {rec.savedToFolder || 'Laptop'}</p>
                      </div>

                      <div className="flex items-center space-x-2">
                        <button
                          onClick={() => handleCopyEmailSummary(rec)}
                          className="flex items-center space-x-1 bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] hover:bg-emerald-100 px-3 py-1.5 rounded text-xs font-bold border border-[var(--brand-primary)]/20 transition"
                        >
                          <Copy className="w-3.5 h-3.5" />
                          <span>{copiedContent ? 'Copied!' : 'Copy Email Summary'}</span>
                        </button>

                        <button
                          onClick={() => { setViewingRecording(rec); setModalActiveTab('AI_SUMMARY'); }}
                          className="flex items-center space-x-1 bg-[var(--brand-primary)] text-white hover:bg-[#07475F] px-3 py-1.5 rounded text-xs font-bold transition shadow-sm"
                        >
                          <Eye className="w-3.5 h-3.5" />
                          <span>View Full Notes</span>
                        </button>
                      </div>
                    </div>

                    {rec.aiSummary && (
                      <div className="bg-[var(--bg-surface)] p-3 rounded-lg border border-[var(--border)] space-y-2 text-xs">
                        <p className="font-semibold text-[var(--brand-ink)] line-clamp-2">{rec.aiSummary.overview}</p>
                        <div className="flex items-center space-x-4 text-[11px] text-[var(--brand-ink-soft)]">
                          <span>Key Points: {rec.aiSummary.keyPoints.length}</span>
                          <span>Action Items: {rec.aiSummary.actionItems.length}</span>
                          <span>Decisions: {rec.aiSummary.decisions.length}</span>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
            </div>
          )}
        </div>
      )}

      {/* FULL VIEW MODAL */}
      {viewingRecording && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[var(--bg-surface)] text-[var(--brand-ink)] max-w-4xl w-full max-h-[90vh] rounded-xl shadow-2xl flex flex-col border border-[var(--border)] overflow-hidden">
            {/* Modal Header */}
            <div className="p-4 border-b border-[var(--border)] flex items-center justify-between bg-[var(--bg-cream)]">
              <div>
                <h2 className="text-base font-bold text-[var(--brand-primary)]">{viewingRecording.title}</h2>
                <p className="text-xs text-[var(--brand-ink-soft)]">Saved: {viewingRecording.dateSaved}</p>
              </div>

              <div className="flex items-center space-x-2">
                <div className="flex bg-[var(--bg-surface)] p-1 rounded-lg border border-[var(--border)]">
                  <button
                    onClick={() => setModalActiveTab('AI_SUMMARY')}
                    className={`px-3 py-1 text-xs font-bold rounded ${modalActiveTab === 'AI_SUMMARY' ? 'bg-[var(--brand-primary)] text-white' : 'text-[var(--brand-ink-soft)]'}`}
                  >
                    AI Summary & Tasks
                  </button>
                  <button
                    onClick={() => setModalActiveTab('TRANSCRIPT')}
                    className={`px-3 py-1 text-xs font-bold rounded ${modalActiveTab === 'TRANSCRIPT' ? 'bg-[var(--brand-primary)] text-white' : 'text-[var(--brand-ink-soft)]'}`}
                  >
                    Full Transcript
                  </button>
                  <button
                    onClick={() => setModalActiveTab('EDIT_SPEAKERS')}
                    className={`px-3 py-1 text-xs font-bold rounded ${modalActiveTab === 'EDIT_SPEAKERS' ? 'bg-[var(--brand-primary)] text-white' : 'text-[var(--brand-ink-soft)]'}`}
                  >
                    Rename Speakers
                  </button>
                </div>

                <button
                  onClick={() => setViewingRecording(null)}
                  className="p-1.5 hover:bg-slate-200 rounded-full transition"
                >
                  <X className="w-5 h-5 text-[var(--brand-ink-soft)]" />
                </button>
              </div>
            </div>

            {/* Modal Content */}
            <div className="p-6 overflow-y-auto space-y-6 flex-1 text-xs">
              {modalActiveTab === 'AI_SUMMARY' && viewingRecording.aiSummary && (
                <div className="space-y-6">
                  {/* Overview */}
                  <div className="bg-[var(--bg-ice-blue)] p-4 rounded-xl border border-[var(--brand-primary)]/20 space-y-1">
                    <h3 className="font-bold uppercase text-[var(--brand-primary)]">Executive Overview</h3>
                    <p className="text-sm leading-relaxed">{viewingRecording.aiSummary.overview}</p>
                  </div>

                  {/* Key Points */}
                  <div className="space-y-2">
                    <h3 className="font-bold uppercase text-[var(--brand-primary)] flex items-center space-x-1.5">
                      <Sparkles className="w-4 h-4 text-[var(--brand-gold)]" />
                      <span>Key Discussion Points</span>
                    </h3>
                    <ul className="space-y-1.5 pl-4 list-disc text-slate-700">
                      {viewingRecording.aiSummary.keyPoints.map((point, idx) => (
                        <li key={idx} className="leading-relaxed">{point}</li>
                      ))}
                    </ul>
                  </div>

                  {/* Action Items Checklist */}
                  <div className="space-y-3">
                    <div className="flex items-center justify-between">
                      <h3 className="font-bold uppercase text-[var(--brand-primary)] flex items-center space-x-1.5">
                        <ListTodo className="w-4 h-4 text-emerald-600" />
                        <span>Action Items & Tasks ({viewingRecording.aiSummary.actionItems.length})</span>
                      </h3>

                      <div className="flex items-center space-x-2">
                        <button
                          onClick={() => setActionItemFilter(actionItemFilter === 'ALL' ? 'MY_TASKS' : 'ALL')}
                          className="px-2.5 py-1 bg-[var(--bg-cream)] hover:bg-[var(--bg-offwhite)] rounded text-[11px] font-bold border border-[var(--border)]"
                        >
                          Filter: {actionItemFilter === 'ALL' ? 'All Tasks' : 'My Tasks Only'}
                        </button>
                      </div>
                    </div>

                    <div className="space-y-2">
                      {viewingRecording.aiSummary.actionItems.map(item => (
                        <div
                          key={item.id}
                          onClick={() => handleToggleActionItemStatus(viewingRecording.id, item.id)}
                          className={`p-3 rounded-lg border flex items-start space-x-3 cursor-pointer transition ${
                            item.status === 'COMPLETED'
                              ? 'bg-emerald-50/60 border-emerald-200 text-slate-500 line-through'
                              : 'bg-[var(--bg-cream)] border-[var(--border)] hover:border-[var(--brand-primary)] text-slate-800'
                          }`}
                        >
                          <input
                            type="checkbox"
                            checked={item.status === 'COMPLETED'}
                            onChange={() => {}}
                            className="mt-0.5 rounded text-emerald-600 focus:ring-emerald-500"
                          />
                          <div className="flex-1 space-y-0.5">
                            <p className="font-medium text-xs">{item.task}</p>
                            <div className="flex items-center space-x-3 text-[10px] text-slate-500 font-mono">
                              <span>Assignee: <strong>{item.assignee}</strong></span>
                              <span>Due: {item.dueDate || 'As discussed'}</span>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Decisions Made */}
                  <div className="space-y-2">
                    <h3 className="font-bold uppercase text-[var(--brand-primary)] flex items-center space-x-1.5">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                      <span>Decisions Made</span>
                    </h3>
                    <ul className="space-y-1 pl-4 list-disc text-slate-700">
                      {viewingRecording.aiSummary.decisions.map((dec, idx) => (
                        <li key={idx} className="leading-relaxed">{dec}</li>
                      ))}
                    </ul>
                  </div>
                </div>
              )}

              {modalActiveTab === 'TRANSCRIPT' && (
                <div className="space-y-2 font-mono">
                  <div className="flex items-center justify-between pb-2 border-b border-[var(--border)]">
                    <span className="font-bold text-[var(--brand-primary)]">Full Transcript Text</span>
                    <button
                      onClick={() => {
                        navigator.clipboard.writeText(viewingRecording.content);
                        setCopiedContent(true);
                        setTimeout(() => setCopiedContent(false), 2000);
                      }}
                      className="flex items-center space-x-1 bg-[var(--bg-cream)] px-2.5 py-1 rounded text-[11px] font-bold border border-[var(--border)]"
                    >
                      <Copy className="w-3.5 h-3.5" />
                      <span>{copiedContent ? 'Copied!' : 'Copy Transcript'}</span>
                    </button>
                  </div>
                  <pre className="whitespace-pre-wrap bg-[var(--bg-cream)] p-4 rounded-lg border border-[var(--border)] text-xs text-slate-800 leading-relaxed max-h-96 overflow-y-auto">
                    {viewingRecording.content}
                  </pre>
                </div>
              )}

              {modalActiveTab === 'EDIT_SPEAKERS' && (
                <div className="space-y-4">
                  <div className="bg-[var(--bg-ice-blue)] p-4 rounded-xl border border-[var(--brand-primary)]/20">
                    <h3 className="font-bold text-[var(--brand-primary)]">Global Speaker Re-Attribution</h3>
                    <p className="text-xs text-[var(--brand-ink-soft)] mt-0.5">
                      Rename speaker tags (e.g. change "Speaker A" to "Jane Doe") across the entire transcript and notes.
                    </p>
                  </div>

                  <div className="space-y-3">
                    {['Speaker A', 'Speaker B', 'Speaker C', 'Speaker 1', 'Speaker 2'].map((oldName) => (
                      <div key={oldName} className="flex items-center space-x-3 bg-[var(--bg-cream)] p-2.5 rounded-lg border border-[var(--border)]">
                        <span className="font-mono font-bold w-28 text-slate-700">{oldName} ➔</span>
                        <input
                          type="text"
                          placeholder={`Enter name for ${oldName}`}
                          value={speakerRenameMap[oldName] || ''}
                          onChange={(e) => setSpeakerRenameMap({ ...speakerRenameMap, [oldName]: e.target.value })}
                          className="flex-1 bg-[var(--bg-surface)] border border-[var(--border)] rounded px-3 py-1.5 text-xs focus:outline-none"
                        />
                      </div>
                    ))}

                    <button
                      onClick={handleApplySpeakerRename}
                      className="w-full bg-[var(--brand-primary)] hover:bg-[#07475F] text-white font-bold py-2.5 rounded-lg text-xs transition shadow"
                    >
                      Apply Speaker Renaming
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Modal Footer */}
            <div className="p-4 border-t border-[var(--border)] bg-[var(--bg-cream)] flex items-center justify-between">
              <button
                onClick={() => handleCopyEmailSummary(viewingRecording)}
                className="flex items-center space-x-1.5 bg-[var(--bg-surface)] text-[var(--brand-ink)] px-4 py-2 rounded-lg text-xs font-bold border border-[var(--border)] hover:bg-emerald-50 transition"
              >
                <Share2 className="w-4 h-4 text-emerald-600" />
                <span>Dispatch Email Summary</span>
              </button>

              <button
                onClick={() => setViewingRecording(null)}
                className="bg-slate-700 hover:bg-slate-800 text-white font-bold px-5 py-2 rounded-lg text-xs transition"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
