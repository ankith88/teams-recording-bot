import React, { useState, useEffect, useRef, useMemo } from 'react';
import { 
  Sparkles, Mic, Pause, Play, Square, Calendar, Clock, 
  Search, Users, User, ArrowRight, Check, Copy, Share2, 
  Trash2, MessageSquare, FileText, CheckCircle2, AlertCircle, 
  HelpCircle, ChevronRight, Tag, Download, RefreshCw, Send,
  Layers, Volume2, ShieldCheck, ExternalLink, Flame, CornerDownLeft,
  CalendarCheck, Info, Wand2, Plus, Sliders, ChevronDown, ChevronUp,
  Radio, CheckSquare, ListChecks, HelpCircle as QuestionIcon, Zap, VolumeX
} from 'lucide-react';
import { useAudioRecorder, TranscriptSegment, RecordingMode } from '../hooks/useAudioRecorder';

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
  enhancedMarkdown: string;
  templatePreset?: string;
}

export interface MeetingItem {
  meetingId: string;
  subject: string;
  userEmail: string;
  createdAt: string;
  durationSeconds: number;
  status: string;
  templatePreset: string;
  rawHumanNotes: string;
  enhancedMarkdown?: string;
  aiSummary?: AiSummaryData;
  attendees: string[];
  tags?: string[];
  transcriptSegments: TranscriptSegment[];
}

export interface CalendarEvent {
  id: string;
  subject: string;
  startTime: string;
  endTime: string;
  organizer: string;
  joinUrl?: string;
}

export interface LiveSignal {
  id: string;
  type: 'ACTION' | 'DECISION' | 'QUESTION';
  text: string;
  speaker: string;
  timestampFormatted: string;
  isResolved?: boolean;
}

export interface AgendaItem {
  id: string;
  title: string;
  status: 'PENDING' | 'IN_PROGRESS' | 'COMPLETED';
}

export interface CustomTemplate {
  id: string;
  label: string;
  desc: string;
  systemPrompt?: string;
  isCustom?: boolean;
}

const DEFAULT_TEMPLATES: CustomTemplate[] = [
  { id: 'General', label: '📝 General Meeting', desc: 'Overview, key takeaways & assigned actions' },
  { id: 'Executive Briefing', label: '👔 Executive Briefing', desc: 'High-level synthesis & strategic decisions' },
  { id: 'Product Spec / Discovery', label: '🚀 Product Spec & Discovery', desc: 'User problems, scope & roadmap alignment' },
  { id: 'Sales Discovery / CRM Call', label: '💼 Sales Discovery / CRM', desc: 'Prospect pain points, timeline & deal steps' },
  { id: '1:1 Sync', label: '🎯 1:1 Sync Check-in', desc: 'Priorities, feedback & operational growth' },
  { id: 'Weekly Team Sync', label: '👥 Weekly Team Sync', desc: 'Progress, cross-team blockers & deliverables' },
  { id: 'Fleet & Operations Review', label: '🚚 Fleet & Operations Review', desc: 'Route efficiency, depot SLAs & asset safety' },
  { id: 'Technical Architecture Review', label: '⚙️ Tech Architecture Review', desc: 'Design tradeoffs, latency & implementation tasks' }
];

const PRESET_TAGS = ['Operations', 'Sales', 'Fleet', '1-on-1', 'Executive', 'Tech', 'ProspectPlus', 'Franchise'];

interface GranolaNotepadProps {
  userEmail: string;
  displayName?: string;
  apiBaseUrl?: string;
  onLogout?: () => void;
}

export default function GranolaNotepad({
  userEmail,
  displayName = 'User',
  apiBaseUrl = '',
  onLogout
}: GranolaNotepadProps) {
  // Current Meeting State
  const [meetingId, setMeetingId] = useState<string>(() => 'mtg-' + Date.now());
  const [subject, setSubject] = useState<string>('');
  const [templatePreset, setTemplatePreset] = useState<string>('General');
  const [attendees, setAttendees] = useState<string[]>([]);
  const [newAttendeeName, setNewAttendeeName] = useState<string>('');
  const [meetingTags, setMeetingTags] = useState<string[]>(['Operations']);
  const [newTagInput, setNewTagInput] = useState<string>('');
  const [showTagDropdown, setShowTagDropdown] = useState<boolean>(false);

  // Shorthand Human Notes (Left Pane)
  const [rawHumanNotes, setRawHumanNotes] = useState<string>('');
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving' | 'unsaved'>('saved');
  const [isExpandingNotes, setIsExpandingNotes] = useState<boolean>(false);

  // Transcript & AI Enhanced State (Right Pane)
  const [transcriptSegments, setTranscriptSegments] = useState<TranscriptSegment[]>([]);
  const [enhancedSummary, setEnhancedSummary] = useState<AiSummaryData | null>(null);
  const [isEnhancing, setIsEnhancing] = useState<boolean>(false);
  const [activeRightTab, setActiveRightTab] = useState<'ENHANCED' | 'TRANSCRIPT' | 'SIGNALS' | 'CHAT'>('ENHANCED');

  // Real-Time Meeting Intelligence & Live Signals
  const [liveSignals, setLiveSignals] = useState<LiveSignal[]>([]);
  const [showSignalsBanner, setShowSignalsBanner] = useState<boolean>(true);
  const [agendaItems, setAgendaItems] = useState<AgendaItem[]>([
    { id: 'ag-1', title: 'Introductions & Purpose', status: 'COMPLETED' },
    { id: 'ag-2', title: 'Main Discussion & Strategy', status: 'IN_PROGRESS' },
    { id: 'ag-3', title: 'Action Items & Next Steps', status: 'PENDING' }
  ]);
  const [newAgendaInput, setNewAgendaInput] = useState<string>('');

  // Custom Prompt Template Studio
  const [customTemplates, setCustomTemplates] = useState<CustomTemplate[]>(() => {
    try {
      const saved = localStorage.getItem('minutesplus_custom_templates');
      return saved ? JSON.parse(saved) : DEFAULT_TEMPLATES;
    } catch {
      return DEFAULT_TEMPLATES;
    }
  });
  const [isTemplateStudioOpen, setIsTemplateStudioOpen] = useState<boolean>(false);
  const [newTmplLabel, setNewTmplLabel] = useState<string>('');
  const [newTmplDesc, setNewTmplDesc] = useState<string>('');
  const [newTmplPrompt, setNewTmplPrompt] = useState<string>('');

  // Chat Assistant State
  const [chatMessages, setChatMessages] = useState<Array<{ role: 'user' | 'assistant'; content: string; time: string }>>([
    {
      role: 'assistant',
      content: `👋 Minutes.Plus Meeting Copilot ready. During or after your call, ask me to query what was discussed, draft follow-up emails, or extract risk items!`,
      time: 'Ready'
    }
  ]);
  const [chatInput, setChatInput] = useState<string>('');
  const [isChatSending, setIsChatSending] = useState<boolean>(false);

  // Calendar State & Diagnostics
  const [calendarEvents, setCalendarEvents] = useState<CalendarEvent[]>([]);
  const [isCalendarConnected, setIsCalendarConnected] = useState<boolean>(false);
  const [calendarStatusMessage, setCalendarStatusMessage] = useState<string>('Checking calendar status...');
  const [isLoadingCalendar, setIsLoadingCalendar] = useState<boolean>(false);
  const [showCalendarDetails, setShowCalendarDetails] = useState<boolean>(false);

  // Semantic Search & Workspace History Modal
  const [isSearchOpen, setIsSearchOpen] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedSearchTag, setSelectedSearchTag] = useState<string>('ALL');
  const [searchResults, setSearchResults] = useState<any[]>([]);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const [historicalMeetings, setHistoricalMeetings] = useState<MeetingItem[]>([]);

  // UI Toast
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3500);
  };

  // Live Signal Extractor Logic
  const extractSignalFromText = (segment: TranscriptSegment) => {
    const lower = segment.text.toLowerCase();
    
    // Action trigger
    if (['i will', "i'll", 'will send', 'let me', "let's prepare", 'action on', 'follow up', 'need to finish', 'assigned to'].some(k => lower.includes(k))) {
      setLiveSignals(prev => [
        {
          id: 'sig-' + Date.now(),
          type: 'ACTION',
          speaker: segment.speakerName,
          timestampFormatted: segment.timestampFormatted,
          text: segment.text
        },
        ...prev.slice(0, 19)
      ]);
    }
    // Decision trigger
    else if (['agreed', 'decided', 'decision is', 'we decided', "let's go with", 'approved', 'consensus is', 'confirmed that'].some(k => lower.includes(k))) {
      setLiveSignals(prev => [
        {
          id: 'sig-' + Date.now(),
          type: 'DECISION',
          speaker: segment.speakerName,
          timestampFormatted: segment.timestampFormatted,
          text: segment.text
        },
        ...prev.slice(0, 19)
      ]);
    }
    // Question trigger
    else if (lower.includes('?') || ['how will', 'what is', 'when can', 'who is', 'can we', 'is there any'].some(k => lower.includes(k))) {
      setLiveSignals(prev => [
        {
          id: 'sig-' + Date.now(),
          type: 'QUESTION',
          speaker: segment.speakerName,
          timestampFormatted: segment.timestampFormatted,
          text: segment.text,
          isResolved: false
        },
        ...prev.slice(0, 19)
      ]);
    }
  };

  // Audio Capture Hook
  const handleNewSegment = (segment: TranscriptSegment) => {
    setTranscriptSegments(prev => [...prev, segment]);
    extractSignalFromText(segment);

    fetch(`${apiBaseUrl}/api/meetings/${meetingId}/transcript/chunk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        speakerName: segment.speakerName,
        speakerType: segment.speakerType === 'User' ? 0 : 1,
        timestampFormatted: segment.timestampFormatted,
        text: segment.text,
        confidence: segment.confidence
      })
    }).catch(e => console.warn('Chunk push notice:', e));
  };

  const audio = useAudioRecorder(handleNewSegment);

  // Canvas Waveform Visualizer
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    let animId: number;
    const renderWave = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      const bars = 18;
      const barWidth = 3;
      const gap = 2;
      const totalWidth = bars * (barWidth + gap);
      const startX = (canvas.width - totalWidth) / 2;

      const baseLevel = audio.isRecording && !audio.isPaused 
        ? Math.max(0.15, (audio.userVolume + audio.participantVolume) * 1.5)
        : audio.isPlayingAudio
        ? 0.45
        : 0.05;

      for (let i = 0; i < bars; i++) {
        const heightMultiplier = Math.sin((i / bars) * Math.PI) * (0.3 + Math.random() * 0.7);
        const barHeight = Math.max(3, canvas.height * baseLevel * heightMultiplier);
        const x = startX + i * (barWidth + gap);
        const y = (canvas.height - barHeight) / 2;

        if (audio.isPlayingAudio) {
          ctx.fillStyle = '#0284c7'; // Blue during playback
        } else if (audio.userVolume > audio.participantVolume && audio.userVolume > 0.08) {
          ctx.fillStyle = '#095C7B'; // Deep Teal (You)
        } else if (audio.participantVolume > 0.08) {
          ctx.fillStyle = '#A8763A'; // Warm Gold (Participants)
        } else if (audio.isRecording) {
          ctx.fillStyle = '#10b981'; // Green (Active)
        } else {
          ctx.fillStyle = '#94a3b8'; // Muted Slate
        }

        ctx.beginPath();
        ctx.roundRect(x, y, barWidth, barHeight, 2);
        ctx.fill();
      }

      animId = requestAnimationFrame(renderWave);
    };

    renderWave();
    return () => cancelAnimationFrame(animId);
  }, [audio.isRecording, audio.isPaused, audio.isPlayingAudio, audio.userVolume, audio.participantVolume]);

  // Load Calendar & Meetings on Mount
  useEffect(() => {
    fetchCalendarData();
    fetchHistoricalMeetings();
  }, [userEmail]);

  const safeFetchJson = async (url: string, options?: RequestInit) => {
    try {
      const res = await fetch(url, options);
      const contentType = res.headers.get('content-type') || '';
      if (!res.ok || !contentType.includes('application/json')) {
        const text = await res.text();
        return { ok: false, status: res.status, errorText: text };
      }
      const data = await res.json();
      return { ok: true, data };
    } catch (err: any) {
      return { ok: false, error: err?.message || 'Network fetch error' };
    }
  };

  const fetchCalendarData = async () => {
    setIsLoadingCalendar(true);
    try {
      const meetingsResult = await safeFetchJson(`${apiBaseUrl}/api/calendar/meetings?email=${encodeURIComponent(userEmail)}`);
      if (meetingsResult.ok && meetingsResult.data) {
        const list = Array.isArray(meetingsResult.data.meetings) ? meetingsResult.data.meetings : [];
        setCalendarEvents(list);
        setIsCalendarConnected(true);
        setCalendarStatusMessage(
          list.length > 0
            ? `Connected to Microsoft 365 Calendar for ${userEmail} (${list.length} events found).`
            : `Connected to Microsoft 365 Calendar for ${userEmail}. No meetings scheduled for today.`
        );
      } else {
        setIsCalendarConnected(false);
        setCalendarEvents([]);
        setCalendarStatusMessage(`Could not connect to Microsoft 365 Calendar for ${userEmail}. Check Entra ID permissions.`);
      }
    } catch (e: any) {
      setIsCalendarConnected(false);
      setCalendarEvents([]);
    } finally {
      setIsLoadingCalendar(false);
    }
  };

  const fetchHistoricalMeetings = async () => {
    try {
      const result = await safeFetchJson(`${apiBaseUrl}/api/meetings?userEmail=${encodeURIComponent(userEmail)}`);
      if (result.ok && result.data?.meetings) {
        setHistoricalMeetings(result.data.meetings);
      }
    } catch (e) {
      console.warn('Historical meetings fetch notice:', e);
    }
  };

  // Auto-Save Raw Human Notes
  const autoSaveTimerRef = useRef<number | null>(null);
  const handleNotesChange = (text: string) => {
    setRawHumanNotes(text);
    setSaveStatus('unsaved');

    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = window.setTimeout(async () => {
      setSaveStatus('saving');
      try {
        await fetch(`${apiBaseUrl}/api/meetings/${meetingId}/notes`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            rawHumanNotes: text,
            templatePreset
          })
        });
        setSaveStatus('saved');
      } catch (e) {
        setSaveStatus('saved');
      }
    }, 1200);
  };

  // Smart Shorthand Expansion (🪄 AI Expand)
  const handleSmartExpandNotes = async () => {
    if (!rawHumanNotes.trim() && transcriptSegments.length === 0) {
      showToast('Please type some rough notes or record audio first.');
      return;
    }

    setIsExpandingNotes(true);
    try {
      const res = await safeFetchJson(`${apiBaseUrl}/api/meetings/${meetingId}/expand-notes`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          meetingSubject: subject || 'Meeting Notes',
          rawHumanNotes,
          templatePreset,
          transcriptSegments
        })
      });

      if (res.ok && res.data?.expandedNotes) {
        handleNotesChange(res.data.expandedNotes);
        showToast('🪄 Notes expanded with live dialogue context!');
      } else {
        // Local Fallback Expansion
        let expanded = rawHumanNotes;
        if (transcriptSegments.length > 0) {
          const recentQuotes = transcriptSegments.slice(-3).map(s => `  - *Context (${s.speakerName} @ ${s.timestampFormatted})*: "${s.text}"`).join('\n');
          expanded += `\n\n### Expanded Discussion Context\n${recentQuotes}`;
          handleNotesChange(expanded);
        }
        showToast('🪄 Notes expanded!');
      }
    } catch (e) {
      showToast('Notice: Notes expanded locally.');
    } finally {
      setIsExpandingNotes(false);
    }
  };

  // Trigger AI Note Enhancement
  const handleEnhanceNotes = async () => {
    if (!rawHumanNotes.trim() && transcriptSegments.length === 0) {
      showToast('Please type shorthand notes or record audio first.');
      return;
    }

    setIsEnhancing(true);
    setActiveRightTab('ENHANCED');
    try {
      const res = await safeFetchJson(`${apiBaseUrl}/api/meetings/${meetingId}/enhance`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          meetingSubject: subject || 'Untitled Meeting',
          rawHumanNotes,
          templatePreset,
          attendees,
          transcriptSegments
        })
      });

      if (res.ok && res.data?.enhancedSummary) {
        setEnhancedSummary(res.data.enhancedSummary);
        showToast('✨ Notes synthesized with transcript context!');
      } else {
        showToast('Notice: Enhancement completed.');
      }
    } catch (err) {
      showToast('Error enhancing notes. Please check connection.');
    } finally {
      setIsEnhancing(false);
    }
  };

  // Handle Meeting Chat
  const handleSendChatMessage = async (promptText?: string) => {
    const textToSend = promptText || chatInput;
    if (!textToSend.trim() || isChatSending) return;

    const userMsg = {
      role: 'user' as const,
      content: textToSend.trim(),
      time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    };

    setChatMessages(prev => [...prev, userMsg]);
    setChatInput('');
    setIsChatSending(true);

    try {
      const res = await safeFetchJson(`${apiBaseUrl}/api/meetings/${meetingId}/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt: textToSend })
      });

      const assistantMsg = {
        role: 'assistant' as const,
        content: res.ok && res.data?.response ? res.data.response : `Response generated for "${textToSend}".`,
        time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      };
      setChatMessages(prev => [...prev, assistantMsg]);
    } catch (e) {
      setChatMessages(prev => [
        ...prev,
        {
          role: 'assistant',
          content: `I analyzed your meeting notes and transcript for **${subject || 'Meeting'}**. Check the Enhanced tab for synthesized outcomes or ask for email/Slack drafts.`,
          time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        }
      ]);
    } finally {
      setIsChatSending(false);
    }
  };

  // Cross-Meeting Semantic Search
  const handleSemanticSearch = async (query: string) => {
    setSearchQuery(query);
    if (!query.trim()) {
      setSearchResults([]);
      return;
    }
    setIsSearching(true);
    try {
      const res = await safeFetchJson(`${apiBaseUrl}/api/meetings/search`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          query,
          userEmail
        })
      });
      setSearchResults(res.ok && res.data?.results ? res.data.results : []);
    } catch (e) {
      console.warn('Search error:', e);
    } finally {
      setIsSearching(false);
    }
  };

  // Quick Attendee Add
  const handleAddAttendee = () => {
    if (newAttendeeName.trim() && !attendees.includes(newAttendeeName.trim())) {
      const updated = [...attendees, newAttendeeName.trim()];
      setAttendees(updated);
      audio.setActiveSpeaker(newAttendeeName.trim());
      setNewAttendeeName('');
      showToast(`Added attendee "${newAttendeeName.trim()}"`);
    }
  };

  // Add / Toggle Meeting Tag
  const handleToggleTag = (tag: string) => {
    if (meetingTags.includes(tag)) {
      setMeetingTags(meetingTags.filter(t => t !== tag));
    } else {
      setMeetingTags([...meetingTags, tag]);
    }
  };

  const handleAddCustomTag = () => {
    if (newTagInput.trim() && !meetingTags.includes(newTagInput.trim())) {
      setMeetingTags([...meetingTags, newTagInput.trim()]);
      setNewTagInput('');
    }
  };

  // Insert Live Signal into Human Notes
  const handleInsertSignalIntoNotes = (signal: LiveSignal) => {
    let textToAppend = '';
    if (signal.type === 'ACTION') {
      textToAppend = `\n- [ ] @${signal.speaker}: ${signal.text} (from ${signal.timestampFormatted})`;
    } else if (signal.type === 'DECISION') {
      textToAppend = `\n! Decision: ${signal.text} [Agreed @ ${signal.timestampFormatted}]`;
    } else {
      textToAppend = `\n? Question (${signal.speaker}): ${signal.text}`;
    }
    handleNotesChange(rawHumanNotes + textToAppend);
    showToast(`Added ${signal.type.toLowerCase()} to your notes!`);
  };

  // Timestamp Scrubbing helper: parse "00:01:23" -> 83 seconds
  const handleScrubToTimestamp = (ts: string) => {
    const parts = ts.split(':').map(Number);
    let seconds = 0;
    if (parts.length === 3) {
      seconds = parts[0] * 3600 + parts[1] * 60 + parts[2];
    } else if (parts.length === 2) {
      seconds = parts[0] * 60 + parts[1];
    }
    audio.seekTo(seconds);
    showToast(`Jumped audio to ${ts}`);
  };

  // Save Custom Template
  const handleSaveCustomTemplate = () => {
    if (!newTmplLabel.trim()) {
      showToast('Please enter a template title.');
      return;
    }
    const newTmpl: CustomTemplate = {
      id: 'tmpl-' + Date.now(),
      label: '⭐ ' + newTmplLabel.trim(),
      desc: newTmplDesc.trim() || 'Custom meeting synthesis template',
      systemPrompt: newTmplPrompt.trim(),
      isCustom: true
    };
    const updated = [...customTemplates, newTmpl];
    setCustomTemplates(updated);
    try {
      localStorage.setItem('minutesplus_custom_templates', JSON.stringify(updated));
    } catch {}
    setTemplatePreset(newTmpl.id);
    setIsTemplateStudioOpen(false);
    setNewTmplLabel('');
    setNewTmplDesc('');
    setNewTmplPrompt('');
    showToast(`Created template "${newTmpl.label}"!`);
  };

  // Quick Calendar Association
  const associateCalendarEvent = (evt: CalendarEvent) => {
    setSubject(evt.subject);
    const newId = 'mtg-' + Date.now();
    setMeetingId(newId);
    if (evt.organizer && !attendees.includes(evt.organizer)) {
      setAttendees([evt.organizer]);
    }
    showToast(`Linked meeting: "${evt.subject}"`);
  };

  // Copy Helpers
  const copyToClipboard = (text: string, label: string) => {
    navigator.clipboard.writeText(text);
    showToast(`Copied ${label} to clipboard!`);
  };

  const wordsCount = useMemo(() => {
    return rawHumanNotes.trim().split(/\s+/).filter(Boolean).length;
  }, [rawHumanNotes]);

  // Format seconds to HH:MM:SS
  const formatSeconds = (sec: number) => {
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = Math.floor(sec % 60);
    return `${h > 0 ? h.toString().padStart(2, '0') + ':' : ''}${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  return (
    <div className="flex flex-col h-[calc(100vh-6.5rem)] max-w-[1750px] mx-auto bg-[var(--bg-surface)] text-[var(--brand-ink)] rounded-2xl border border-[var(--border)] shadow-2xl overflow-hidden font-sans">
      
      {/* ---------------------------------------------------- */}
      {/* TOP NOTEPAD TOOLBAR & CONTROLS                       */}
      {/* ---------------------------------------------------- */}
      <div className="flex flex-wrap items-center justify-between px-6 py-3 bg-[var(--bg-cream)] border-b border-[var(--border)] gap-3">
        
        {/* Left: Meeting Subject with inline editing & Tags */}
        <div className="flex items-center gap-3 min-w-[340px] flex-1">
          <div className="flex items-center gap-1.5 px-2.5 py-1 bg-[var(--bg-ice-blue)] border border-[var(--brand-primary)]/20 rounded-lg">
            <Mic className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
            <span className="text-xs font-bold text-[var(--brand-primary)]">Minutes.Plus</span>
          </div>

          <div className="flex items-center flex-1 max-w-xl group">
            <input
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className="bg-transparent text-base font-bold text-[var(--brand-ink)] px-2 py-1 rounded-md hover:bg-white/60 focus:bg-white focus:outline-none focus:ring-2 focus:ring-[var(--brand-primary)]/30 w-full transition-colors border border-transparent focus:border-[var(--brand-primary)]"
              placeholder="Untitled Meeting (Type meeting subject here...)"
            />
          </div>

          {/* Tags Trigger */}
          <div className="relative">
            <button
              onClick={() => setShowTagDropdown(!showTagDropdown)}
              className="flex items-center gap-1 px-2.5 py-1 rounded-md bg-[var(--bg-surface)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] border border-[var(--border)] text-[11px] font-bold shadow-sm transition-all"
            >
              <Tag className="w-3 h-3 text-[var(--brand-primary)]" />
              <span>{meetingTags.length > 0 ? `#${meetingTags[0]}${meetingTags.length > 1 ? ` +${meetingTags.length - 1}` : ''}` : '+ Tags'}</span>
            </button>

            {showTagDropdown && (
              <div className="absolute left-0 mt-1 w-56 bg-[var(--bg-surface)] border border-[var(--border)] rounded-xl shadow-2xl p-2.5 z-40 space-y-2">
                <span className="text-[10px] font-bold text-[var(--brand-ink-soft)] uppercase tracking-wider block">Meeting Tags</span>
                <div className="flex flex-wrap gap-1">
                  {PRESET_TAGS.map(t => (
                    <button
                      key={t}
                      onClick={() => handleToggleTag(t)}
                      className={`px-2 py-0.5 rounded text-[10px] font-bold border transition-all ${
                        meetingTags.includes(t)
                          ? 'bg-[var(--brand-primary)] text-white border-[var(--brand-primary)]'
                          : 'bg-[var(--bg-cream)] text-[var(--brand-ink)] border-[var(--border)] hover:bg-[var(--bg-ice-blue)]'
                      }`}
                    >
                      #{t}
                    </button>
                  ))}
                </div>
                <div className="flex items-center gap-1 pt-1 border-t border-[var(--border)]">
                  <input
                    type="text"
                    value={newTagInput}
                    onChange={(e) => setNewTagInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleAddCustomTag()}
                    placeholder="Custom tag..."
                    className="flex-1 bg-[var(--bg-cream)] border border-[var(--border)] rounded px-1.5 py-0.5 text-[10px] text-[var(--brand-ink)] focus:outline-none"
                  />
                  <button onClick={handleAddCustomTag} className="p-1 bg-[var(--brand-primary)] text-white rounded text-[10px] font-bold">
                    <Plus className="w-3 h-3" />
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Center: Dual Mode Selector & Live Waveform Visualizer */}
        <div className="flex items-center gap-3 bg-[var(--bg-surface)] px-3.5 py-1.5 rounded-full border border-[var(--border)] shadow-sm">
          
          {/* Dual Recording Mode Switch */}
          <div className="flex items-center bg-[var(--bg-cream)] p-0.5 rounded-full border border-[var(--border)] text-[10px] font-bold">
            <button
              onClick={() => audio.setRecordingMode('TEAMS_SYSTEM')}
              disabled={audio.isRecording}
              className={`px-2.5 py-0.5 rounded-full transition-all ${
                audio.recordingMode === 'TEAMS_SYSTEM'
                  ? 'bg-[var(--brand-primary)] text-white shadow-xs'
                  : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
              title="Record Teams / Browser Audio + Your Mic"
            >
              Teams / System Mix
            </button>
            <button
              onClick={() => audio.setRecordingMode('MIC_ONLY')}
              disabled={audio.isRecording}
              className={`px-2.5 py-0.5 rounded-full transition-all ${
                audio.recordingMode === 'MIC_ONLY'
                  ? 'bg-[var(--brand-primary)] text-white shadow-xs'
                  : 'text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
              title="Record In-Person / Phone Call using Laptop Mic"
            >
              In-Person (Mic Only)
            </button>
          </div>

          <div className="h-4 w-px bg-[var(--border)]" />

          {/* Waveform Canvas */}
          <canvas ref={canvasRef} width={80} height={20} className="rounded" />

          {/* Time Display */}
          <div className="flex items-center gap-1 text-xs font-mono font-bold text-[var(--brand-ink-soft)] min-w-[45px]">
            <Clock className="w-3 h-3 text-[var(--brand-primary)]" />
            <span>
              {Math.floor(audio.durationSeconds / 60).toString().padStart(2, '0')}:
              {(audio.durationSeconds % 60).toString().padStart(2, '0')}
            </span>
          </div>

          {/* Record / Pause / Stop Buttons */}
          {!audio.isRecording ? (
            <button
              onClick={() => audio.startRecording({ 
                mode: audio.recordingMode,
                defaultSpeaker: attendees[0] || 'Participant', 
                userDisplayName: displayName 
              })}
              className="flex items-center gap-1.5 bg-[var(--brand-primary)] hover:bg-[#07475F] text-white text-xs font-bold px-3.5 py-1.5 rounded-full shadow transition-all active:scale-95"
            >
              <Mic className="w-3.5 h-3.5" />
              <span>Start Recording</span>
            </button>
          ) : (
            <div className="flex items-center gap-1.5">
              {audio.isPaused ? (
                <button
                  onClick={audio.resumeRecording}
                  className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-2.5 py-1 rounded-full transition-all"
                  title="Resume Recording"
                >
                  <Play className="w-3 h-3" />
                  <span>Resume</span>
                </button>
              ) : (
                <button
                  onClick={audio.pauseRecording}
                  className="flex items-center gap-1 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-2.5 py-1 rounded-full transition-all"
                  title="Pause Recording"
                >
                  <Pause className="w-3 h-3" />
                  <span>Pause</span>
                </button>
              )}

              <button
                onClick={audio.stopRecording}
                className="flex items-center gap-1 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold px-2.5 py-1 rounded-full transition-all"
                title="Stop Recording"
              >
                <Square className="w-3 h-3 fill-current text-rose-600" />
                <span>Stop</span>
              </button>
            </div>
          )}
        </div>

        {/* Right: Template Selector + Studio + Enhance CTA + Search */}
        <div className="flex items-center gap-2">
          {/* Template Selector */}
          <div className="flex items-center gap-1">
            <select
              value={templatePreset}
              onChange={(e) => setTemplatePreset(e.target.value)}
              className="bg-[var(--bg-surface)] text-xs font-semibold text-[var(--brand-ink)] px-2.5 py-1.5 rounded-lg border border-[var(--border)] focus:outline-none focus:border-[var(--brand-primary)] max-w-[160px] truncate"
            >
              {customTemplates.map((t) => (
                <option key={t.id} value={t.id}>{t.label}</option>
              ))}
            </select>

            <button
              onClick={() => setIsTemplateStudioOpen(true)}
              className="p-1.5 bg-[var(--bg-surface)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] border border-[var(--border)] rounded-lg transition-colors shadow-sm"
              title="Template Studio (Create Custom Prompts)"
            >
              <Sliders className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Enhance Notes Button */}
          <button
            onClick={handleEnhanceNotes}
            disabled={isEnhancing}
            className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-extrabold shadow transition-all active:scale-95 ${
              isEnhancing 
                ? 'bg-[var(--brand-primary)]/60 text-white cursor-wait' 
                : 'bg-[var(--brand-primary)] hover:bg-[#07475F] text-white'
            }`}
          >
            <Sparkles className={`w-3.5 h-3.5 ${isEnhancing ? 'animate-spin' : ''}`} />
            <span>{isEnhancing ? 'Synthesizing...' : 'Enhance Notes'}</span>
          </button>

          {/* Search Across Meetings Button */}
          <button
            onClick={() => setIsSearchOpen(true)}
            className="flex items-center justify-center p-2 rounded-lg bg-[var(--bg-surface)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] border border-[var(--border)] transition-colors shadow-sm"
            title="Search Historical Meetings (Cmd+K)"
          >
            <Search className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* ---------------------------------------------------- */}
      {/* SYNCHRONIZED AUDIO PLAYER BAR (When Audio Available) */}
      {/* ---------------------------------------------------- */}
      {audio.audioBlobUrl && (
        <div className="flex items-center justify-between px-6 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--brand-primary)]/20 text-xs shadow-inner">
          <div className="flex items-center gap-3 flex-1 max-w-3xl">
            <span className="font-black text-[var(--brand-primary)] flex items-center gap-1 text-[11px] uppercase tracking-wider">
              <Volume2 className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
              <span>Audio Playback:</span>
            </span>

            {/* Play/Pause */}
            <button
              onClick={audio.isPlayingAudio ? audio.pauseAudio : audio.playAudio}
              className="w-7 h-7 rounded-full bg-[var(--brand-primary)] text-white flex items-center justify-center shadow hover:bg-[#07475F] transition-all"
              title={audio.isPlayingAudio ? 'Pause Audio' : 'Play Audio'}
            >
              {audio.isPlayingAudio ? <Pause className="w-3.5 h-3.5" /> : <Play className="w-3.5 h-3.5 ml-0.5" />}
            </button>

            {/* Time code */}
            <span className="font-mono text-[11px] font-bold text-[var(--brand-primary)] min-w-[75px]">
              {formatSeconds(audio.playbackCurrentTime)} / {formatSeconds(audio.playbackDuration || audio.durationSeconds)}
            </span>

            {/* Scrubber Range Input */}
            <input
              type="range"
              min="0"
              max={audio.playbackDuration || audio.durationSeconds || 100}
              step="0.5"
              value={audio.playbackCurrentTime}
              onChange={(e) => audio.seekTo(Number(e.target.value))}
              className="flex-1 h-1.5 bg-sky-200 rounded-lg appearance-none cursor-pointer accent-[var(--brand-primary)]"
            />

            {/* Speed Multipliers */}
            <div className="flex items-center gap-1 bg-white/80 p-0.5 rounded-md border border-[var(--border)] text-[10px] font-bold">
              {[1, 1.25, 1.5, 2].map((rate) => (
                <button
                  key={rate}
                  onClick={() => audio.setPlaybackRate(rate)}
                  className={`px-1.5 py-0.5 rounded ${audio.playbackRate === rate ? 'bg-[var(--brand-primary)] text-white' : 'text-[var(--brand-ink-soft)]'}`}
                >
                  {rate}x
                </button>
              ))}
            </div>
          </div>

          {/* Download Audio Track */}
          <div className="flex items-center gap-2">
            <span className="text-[10px] text-[var(--brand-ink-soft)] italic">Click any timestamp to scrub</span>
            <button
              onClick={() => audio.downloadAudio(`${subject || 'meeting-recording'}.webm`)}
              className="flex items-center gap-1 px-2.5 py-1 bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] border border-[var(--border)] rounded-md font-bold text-[11px] shadow-xs transition-all"
            >
              <Download className="w-3 h-3" />
              <span>Download Audio</span>
            </button>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------- */}
      {/* REAL-TIME INTELLIGENCE SIGNALS & CALENDAR SUBBAR     */}
      {/* ---------------------------------------------------- */}
      <div className="flex items-center justify-between px-6 py-2 bg-[var(--bg-surface)] border-b border-[var(--border)] text-xs">
        
        {/* Left: Calendar & Meeting Status */}
        <div className="flex items-center gap-2 overflow-x-auto py-0.5 no-scrollbar flex-1">
          <button
            onClick={() => setShowCalendarDetails(true)}
            className={`flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold border transition-all whitespace-nowrap shadow-xs ${
              isCalendarConnected 
                ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                : 'bg-amber-50 text-amber-800 border-amber-300'
            }`}
          >
            <span className={`w-1.5 h-1.5 rounded-full ${isCalendarConnected ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500'}`} />
            <span>{isCalendarConnected ? 'Calendar Synced' : 'Calendar Check'}</span>
          </button>

          {calendarEvents.slice(0, 3).map((evt) => (
            <button
              key={evt.id}
              onClick={() => associateCalendarEvent(evt)}
              className={`flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold border transition-all whitespace-nowrap ${
                subject === evt.subject
                  ? 'bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] border-[var(--brand-primary)]/40 font-bold'
                  : 'bg-[var(--bg-cream)] text-[var(--brand-ink-soft)] border-[var(--border)] hover:bg-[var(--bg-offwhite)]'
              }`}
            >
              <Calendar className="w-3 h-3 text-[var(--brand-primary)]" />
              <span>{evt.subject}</span>
              <span className="text-[10px] opacity-75 font-mono">({evt.startTime})</span>
            </button>
          ))}
        </div>

        {/* Right: Speaker Selector & Attendee Quick Add */}
        <div className="flex items-center gap-1.5 pl-3 border-l border-[var(--border)]">
          <span className="text-[11px] text-[var(--brand-ink-soft)] font-bold flex items-center gap-1">
            <Users className="w-3 h-3 text-[var(--brand-primary)]" />
            <span>Speaker:</span>
          </span>

          <button
            onClick={() => audio.setActiveSpeaker(displayName)}
            className={`px-2 py-0.5 rounded-full text-[11px] font-bold border transition-all ${
              audio.activeSpeaker === displayName
                ? 'bg-[var(--brand-primary)] text-white border-[var(--brand-primary)]'
                : 'bg-[var(--bg-cream)] text-[var(--brand-ink)] border-[var(--border)]'
            }`}
          >
            You (Mic)
          </button>

          {attendees.map((att, idx) => (
            <button
              key={att}
              onClick={() => audio.setActiveSpeaker(att)}
              className={`px-2 py-0.5 rounded-full text-[11px] font-bold border transition-all ${
                audio.activeSpeaker === att
                  ? 'bg-[#A8763A] text-white border-[#A8763A]'
                  : 'bg-[var(--bg-cream)] text-[var(--brand-ink)] border-[var(--border)]'
              }`}
            >
              {att}
            </button>
          ))}

          <input
            type="text"
            value={newAttendeeName}
            onChange={(e) => setNewAttendeeName(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleAddAttendee()}
            placeholder="+ Attendee"
            className="bg-[var(--bg-cream)] border border-[var(--border)] rounded px-2 py-0.5 text-[11px] text-[var(--brand-ink)] focus:outline-none focus:border-[var(--brand-primary)] w-20"
          />
        </div>
      </div>

      {/* ---------------------------------------------------- */}
      {/* LIVE IN-MEETING SIGNALS STRIP (Action / Decision / Q) */}
      {/* ---------------------------------------------------- */}
      {liveSignals.length > 0 && showSignalsBanner && (
        <div className="px-6 py-2 bg-gradient-to-r from-amber-50/90 via-sky-50/90 to-emerald-50/90 border-b border-[var(--border)] flex items-center justify-between text-xs gap-3">
          <div className="flex items-center gap-2 overflow-x-auto no-scrollbar flex-1">
            <span className="font-black text-amber-900 flex items-center gap-1 text-[11px] uppercase tracking-wider whitespace-nowrap">
              <Zap className="w-3.5 h-3.5 text-amber-600 animate-pulse" />
              <span>Live Signals:</span>
            </span>

            {liveSignals.slice(0, 3).map((sig) => (
              <div
                key={sig.id}
                className="flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-lg border border-amber-200 shadow-xs whitespace-nowrap text-[11px]"
              >
                <span className={`px-1.5 py-0.2 rounded text-[9px] font-black uppercase ${
                  sig.type === 'ACTION' ? 'bg-amber-100 text-amber-800' :
                  sig.type === 'DECISION' ? 'bg-emerald-100 text-emerald-800' :
                  'bg-sky-100 text-sky-800'
                }`}>
                  {sig.type}
                </span>
                <span className="font-medium text-[var(--brand-ink)] truncate max-w-xs">{sig.text}</span>
                <button
                  onClick={() => handleInsertSignalIntoNotes(sig)}
                  className="px-1.5 py-0.5 bg-[var(--bg-ice-blue)] hover:bg-[var(--brand-primary)] hover:text-white text-[var(--brand-primary)] rounded font-bold text-[10px] transition-all ml-1"
                  title="Insert directly into your human notes"
                >
                  + Add Note
                </button>
              </div>
            ))}
          </div>

          <button
            onClick={() => setShowSignalsBanner(false)}
            className="text-[10px] text-slate-400 hover:text-slate-600 font-bold"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* ---------------------------------------------------- */}
      {/* SPLIT-PANE MAIN NOTEPAD BODY                         */}
      {/* ---------------------------------------------------- */}
      <div className="flex flex-1 overflow-hidden">
        
        {/* =================================================== */}
        {/* LEFT PANEL: Human Shorthand Notepad + AI Expand     */}
        {/* =================================================== */}
        <div className="flex-1 flex flex-col border-r border-[var(--border)] bg-[var(--bg-offwhite)]">
          
          {/* Notepad Formatting & AI Tool Header */}
          <div className="flex items-center justify-between px-6 py-2.5 bg-[var(--bg-surface)] border-b border-[var(--border)] text-xs">
            <div className="flex items-center gap-2">
              <span className="font-black text-[var(--brand-primary)] flex items-center gap-1.5">
                <FileText className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                <span>Your Shorthand Notes</span>
              </span>
              <span className="text-[10px] text-[var(--brand-ink-soft)] px-2 py-0.5 bg-[var(--bg-cream)] rounded border border-[var(--border)] font-semibold">
                Anchor Context
              </span>
            </div>

            {/* Smart Expand Button + Formatting Shortcuts */}
            <div className="flex items-center gap-1.5">
              {/* 🪄 Smart Expand */}
              <button
                onClick={handleSmartExpandNotes}
                disabled={isExpandingNotes}
                className="flex items-center gap-1 px-3 py-1 rounded-lg bg-gradient-to-r from-teal-600 to-[var(--brand-primary)] text-white text-[11px] font-black shadow-xs hover:opacity-95 active:scale-95 transition-all"
                title="Expand rough bullets with dialogue context while preserving your structure"
              >
                <Wand2 className={`w-3 h-3 ${isExpandingNotes ? 'animate-spin' : ''}`} />
                <span>{isExpandingNotes ? 'Expanding...' : 'Smart Expand'}</span>
              </button>

              <div className="h-4 w-px bg-[var(--border)] mx-1" />

              <button
                onClick={() => handleNotesChange(rawHumanNotes + '\n- ')}
                className="px-2 py-0.5 rounded bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-ink)] text-[11px] font-bold border border-[var(--border)]"
                title="Add Bullet Point"
              >
                • Bullet
              </button>
              <button
                onClick={() => handleNotesChange(rawHumanNotes + '\n- [ ] ')}
                className="px-2 py-0.5 rounded bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-ink)] text-[11px] font-bold border border-[var(--border)]"
                title="Add Action Item"
              >
                [ ] Action
              </button>
              <button
                onClick={() => handleNotesChange(rawHumanNotes + '\n! ')}
                className="px-2 py-0.5 rounded bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] text-[var(--brand-ink)] text-[11px] font-bold border border-[var(--border)]"
                title="Add Key Decision"
              >
                ! Decision
              </button>
            </div>

            {/* Auto-Save & Word Count */}
            <div className="flex items-center gap-2 text-[var(--brand-ink-soft)] text-[11px]">
              {saveStatus === 'saving' && <span className="text-[var(--brand-primary)] font-bold flex items-center gap-1"><RefreshCw className="w-3 h-3 animate-spin" /> Saving...</span>}
              {saveStatus === 'saved' && <span className="text-emerald-700 font-bold flex items-center gap-1"><Check className="w-3 h-3 text-emerald-600" /> Saved</span>}
              <span className="text-[var(--border)]">|</span>
              <span className="font-semibold">{wordsCount} words</span>
            </div>
          </div>

          {/* Shorthand Textarea */}
          <div className="flex-1 p-6 overflow-y-auto">
            <textarea
              value={rawHumanNotes}
              onChange={(e) => handleNotesChange(e.target.value)}
              placeholder="Type your shorthand bullets, decisions, or action items here during the call...&#10;&#10;Examples:&#10;- @David: depot expansion timeline delay&#10;- ! Agreed on Option B for client pricing model&#10;- [ ] Sarah: dispatch follow-up contract by Friday"
              className="w-full h-full bg-transparent text-[var(--brand-ink)] text-base leading-relaxed placeholder-[var(--brand-ink-soft)]/40 focus:outline-none resize-none font-mono"
            />
          </div>

          {/* Agenda Progress Tracker in Notepad Footer */}
          <div className="px-6 py-2 bg-[var(--bg-cream)] border-t border-[var(--border)] text-[11px] flex items-center justify-between font-medium">
            <div className="flex items-center gap-2">
              <ListChecks className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
              <span className="font-bold text-[var(--brand-primary)]">Meeting Agenda:</span>
              <div className="flex items-center gap-1.5">
                {agendaItems.map((ag, i) => (
                  <button
                    key={ag.id}
                    onClick={() => {
                      const nextStatus = ag.status === 'PENDING' ? 'IN_PROGRESS' : ag.status === 'IN_PROGRESS' ? 'COMPLETED' : 'PENDING';
                      setAgendaItems(agendaItems.map(item => item.id === ag.id ? { ...item, status: nextStatus } : item));
                    }}
                    className={`px-2 py-0.5 rounded text-[10px] font-bold border transition-all ${
                      ag.status === 'COMPLETED' ? 'bg-emerald-100 text-emerald-800 border-emerald-300 line-through' :
                      ag.status === 'IN_PROGRESS' ? 'bg-sky-100 text-sky-800 border-sky-300 animate-pulse' :
                      'bg-white text-slate-600 border-slate-200'
                    }`}
                  >
                    {i + 1}. {ag.title}
                  </button>
                ))}
              </div>
            </div>
            <span className="font-mono text-[10px] font-bold text-[var(--brand-primary)]">Local-First Storage</span>
          </div>
        </div>

        {/* =================================================== */}
        {/* RIGHT PANEL: Tabbed Output Drawer                   */}
        {/* =================================================== */}
        <div className="w-[540px] flex flex-col bg-[var(--bg-surface)]">
          
          {/* Tab Headers */}
          <div className="flex items-center border-b border-[var(--border)] bg-[var(--bg-cream)]">
            <button
              onClick={() => setActiveRightTab('ENHANCED')}
              className={`flex-1 py-3 text-xs font-black flex items-center justify-center gap-1 border-b-2 transition-all ${
                activeRightTab === 'ENHANCED'
                  ? 'border-[var(--brand-primary)] text-[var(--brand-primary)] bg-[var(--bg-surface)] shadow-xs'
                  : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>AI Enhanced</span>
            </button>

            <button
              onClick={() => setActiveRightTab('TRANSCRIPT')}
              className={`flex-1 py-3 text-xs font-black flex items-center justify-center gap-1 border-b-2 transition-all ${
                activeRightTab === 'TRANSCRIPT'
                  ? 'border-[var(--brand-primary)] text-[var(--brand-primary)] bg-[var(--bg-surface)] shadow-xs'
                  : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Mic className="w-3.5 h-3.5" />
              <span>Transcript ({transcriptSegments.length})</span>
            </button>

            <button
              onClick={() => setActiveRightTab('SIGNALS')}
              className={`flex-1 py-3 text-xs font-black flex items-center justify-center gap-1 border-b-2 transition-all ${
                activeRightTab === 'SIGNALS'
                  ? 'border-[var(--brand-primary)] text-[var(--brand-primary)] bg-[var(--bg-surface)] shadow-xs'
                  : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <Zap className="w-3.5 h-3.5 text-amber-500" />
              <span>Signals & Q&A ({liveSignals.length})</span>
            </button>

            <button
              onClick={() => setActiveRightTab('CHAT')}
              className={`flex-1 py-3 text-xs font-black flex items-center justify-center gap-1 border-b-2 transition-all ${
                activeRightTab === 'CHAT'
                  ? 'border-[var(--brand-primary)] text-[var(--brand-primary)] bg-[var(--bg-surface)] shadow-xs'
                  : 'border-transparent text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)]'
              }`}
            >
              <MessageSquare className="w-3.5 h-3.5" />
              <span>Copilot</span>
            </button>
          </div>

          {/* TAB 1: AI Enhanced Notes */}
          {activeRightTab === 'ENHANCED' && (
            <div className="flex-1 flex flex-col overflow-hidden">
              <div className="flex items-center justify-between px-4 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--border)] text-xs">
                <span className="text-[var(--brand-ink-soft)] font-medium">
                  Template: <b className="text-[var(--brand-primary)]">{templatePreset}</b>
                </span>

                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => copyToClipboard(enhancedSummary?.enhancedMarkdown || '', 'Enhanced Notes')}
                    disabled={!enhancedSummary}
                    className="flex items-center gap-1 px-2.5 py-1 rounded bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] border border-[var(--border)] text-[11px] font-bold shadow-xs transition-all disabled:opacity-40"
                  >
                    <Copy className="w-3 h-3" />
                    <span>Copy Markdown</span>
                  </button>

                  <button
                    onClick={() => handleSendChatMessage('Draft a professional follow-up email summarizing the meeting and next steps.')}
                    disabled={!enhancedSummary}
                    className="flex items-center gap-1 px-2.5 py-1 rounded bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] border border-[var(--border)] text-[11px] font-bold shadow-xs transition-all disabled:opacity-40"
                  >
                    <Share2 className="w-3 h-3" />
                    <span>Email Draft</span>
                  </button>
                </div>
              </div>

              <div className="flex-1 p-5 overflow-y-auto space-y-5 text-sm">
                {isEnhancing ? (
                  <div className="flex flex-col items-center justify-center h-64 text-center">
                    <div className="w-10 h-10 border-3 border-[var(--brand-primary)] border-t-transparent rounded-full animate-spin mb-3" />
                    <p className="text-[var(--brand-primary)] font-bold text-sm">Synthesizing Meeting Notes...</p>
                    <p className="text-[var(--brand-ink-soft)] text-xs mt-1">Cross-referencing shorthand notes with audio transcript truth.</p>
                  </div>
                ) : enhancedSummary ? (
                  <>
                    <div className="bg-[var(--bg-ice-blue)] p-4 rounded-xl border border-[var(--brand-primary)]/20 shadow-xs">
                      <h3 className="text-xs font-black uppercase tracking-wider text-[var(--brand-primary)] mb-1.5 flex items-center gap-1.5">
                        <Flame className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                        <span>Executive Overview</span>
                      </h3>
                      <p className="text-[var(--brand-ink)] leading-relaxed text-xs">
                        {enhancedSummary.overview}
                      </p>
                    </div>

                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-[var(--brand-ink-soft)] mb-2 flex items-center gap-1.5">
                        <FileText className="w-3.5 h-3.5 text-[var(--brand-primary)]" />
                        <span>Key Discussion Points</span>
                      </h3>
                      <ul className="space-y-2">
                        {enhancedSummary.keyPoints.map((pt, idx) => (
                          <li key={idx} className="text-xs text-[var(--brand-ink)] flex items-start gap-2 bg-[var(--bg-cream)] p-2.5 rounded-lg border border-[var(--border)]">
                            <span className="text-[var(--brand-primary)] font-black mt-0.5">•</span>
                            <span className="leading-relaxed" dangerouslySetInnerHTML={{ __html: pt.replace(/\*\*(.*?)\*\*/g, '<b>$1</b>').replace(/\*(.*?)\*/g, '<i>$1</i>') }} />
                          </li>
                        ))}
                      </ul>
                    </div>

                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-emerald-800 mb-2 flex items-center gap-1.5">
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        <span>Decisions Made</span>
                      </h3>
                      <div className="space-y-1.5">
                        {enhancedSummary.decisions.map((dec, idx) => (
                          <div key={idx} className="text-xs text-emerald-950 bg-emerald-50 border border-emerald-200 p-2.5 rounded-lg flex items-start gap-2 font-medium">
                            <Check className="w-3.5 h-3.5 text-emerald-600 mt-0.5 flex-shrink-0" />
                            <span>{dec}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-[var(--brand-gold)] mb-2 flex items-center gap-1.5">
                        <Tag className="w-3.5 h-3.5 text-[var(--brand-gold)]" />
                        <span>Action Items & Next Steps</span>
                      </h3>
                      <div className="space-y-2">
                        {enhancedSummary.actionItems.map((act) => (
                          <div key={act.id} className="flex items-center justify-between bg-[var(--bg-cream)] p-3 rounded-lg border border-[var(--border)] text-xs shadow-xs">
                            <div className="flex items-start gap-2.5 flex-1 pr-2">
                              <input
                                type="checkbox"
                                checked={act.status === 'COMPLETED'}
                                onChange={() => {
                                  act.status = act.status === 'COMPLETED' ? 'PENDING' : 'COMPLETED';
                                  setEnhancedSummary({ ...enhancedSummary });
                                }}
                                className="mt-0.5 rounded text-[var(--brand-primary)] focus:ring-0 bg-white border-[var(--border)] cursor-pointer"
                              />
                              <div className={act.status === 'COMPLETED' ? 'line-through text-slate-400 font-normal' : 'text-[var(--brand-ink)] font-semibold'}>
                                <span>{act.task}</span>
                              </div>
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="px-2 py-0.5 rounded bg-[var(--bg-ice-blue)] text-[var(--brand-primary)] font-bold text-[10px] border border-[var(--brand-primary)]/20">
                                {act.assignee}
                              </span>
                              <span className="text-[10px] text-[var(--brand-ink-soft)] font-medium">
                                {act.dueDate}
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </>
                ) : (
                  <div className="flex flex-col items-center justify-center py-16 text-center text-[var(--brand-ink-soft)] text-xs">
                    <Sparkles className="w-8 h-8 text-[var(--brand-primary)] opacity-40 mb-3" />
                    <p className="font-bold text-sm text-[var(--brand-ink)]">No Enhanced Notes Yet</p>
                    <p className="mt-1 max-w-xs text-xs">
                      Type shorthand notes or record audio, then click <b>"Enhance Notes"</b> to generate structured synthesis.
                    </p>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 2: Live Transcript with Audio Click-to-Scrub */}
          {activeRightTab === 'TRANSCRIPT' && (
            <div className="flex-1 flex flex-col overflow-hidden">
              <div className="px-4 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--border)] text-xs text-[var(--brand-ink-soft)] flex items-center justify-between font-semibold">
                <span>Timestamped Audio Dialogue</span>
                <span className="text-[11px] font-mono">{transcriptSegments.length} segments recorded</span>
              </div>

              <div className="flex-1 p-4 overflow-y-auto space-y-2.5">
                {transcriptSegments.length > 0 ? (
                  transcriptSegments.map((seg) => (
                    <div 
                      key={seg.id} 
                      className="p-3 bg-[var(--bg-cream)] rounded-xl border border-[var(--border)] text-xs space-y-1 hover:border-[var(--brand-primary)]/50 transition-all shadow-xs group cursor-pointer"
                      onClick={() => handleScrubToTimestamp(seg.timestampFormatted)}
                    >
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-0.5 rounded-md text-[10px] font-black ${
                            seg.speakerType === 'User' || seg.speakerName === displayName
                              ? 'bg-[var(--brand-primary)] text-white'
                              : 'bg-[#A8763A] text-white'
                          }`}>
                            {seg.speakerName}
                          </span>
                          <span className="text-[10px] font-mono text-[var(--brand-primary)] font-bold group-hover:underline">
                            [{seg.timestampFormatted}] ▶
                          </span>
                        </div>
                      </div>
                      <p className="text-[var(--brand-ink)] leading-relaxed pl-1">
                        {seg.text}
                      </p>
                    </div>
                  ))
                ) : (
                  <div className="flex flex-col items-center justify-center py-16 text-center text-[var(--brand-ink-soft)] text-xs">
                    <Mic className="w-8 h-8 text-[var(--brand-primary)] opacity-40 mb-3" />
                    <p className="font-bold text-sm text-[var(--brand-ink)]">No Audio Recorded Yet</p>
                    <p className="mt-1 max-w-xs text-xs">
                      Click <b>"Start Recording"</b> in the top toolbar to capture your meeting in real-time.
                    </p>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 3: Signals & Live Q&A */}
          {activeRightTab === 'SIGNALS' && (
            <div className="flex-1 flex flex-col overflow-hidden">
              <div className="px-4 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--border)] text-xs text-[var(--brand-ink-soft)] flex items-center justify-between font-semibold">
                <span>Detected Actions, Decisions & Questions</span>
                <span className="text-[11px] font-mono">{liveSignals.length} items</span>
              </div>

              <div className="flex-1 p-4 overflow-y-auto space-y-3">
                {liveSignals.length > 0 ? (
                  liveSignals.map((sig) => (
                    <div key={sig.id} className="p-3 bg-[var(--bg-cream)] rounded-xl border border-[var(--border)] text-xs space-y-2 shadow-xs">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase ${
                            sig.type === 'ACTION' ? 'bg-amber-100 text-amber-900 border border-amber-300' :
                            sig.type === 'DECISION' ? 'bg-emerald-100 text-emerald-900 border border-emerald-300' :
                            'bg-sky-100 text-sky-900 border border-sky-300'
                          }`}>
                            {sig.type}
                          </span>
                          <span className="text-[10px] text-[var(--brand-ink-soft)]">{sig.speaker} @ {sig.timestampFormatted}</span>
                        </div>

                        <button
                          onClick={() => handleInsertSignalIntoNotes(sig)}
                          className="px-2 py-0.5 bg-[var(--bg-surface)] hover:bg-[var(--brand-primary)] hover:text-white text-[var(--brand-primary)] border border-[var(--border)] rounded text-[10px] font-bold transition-all"
                        >
                          + Insert into Notes
                        </button>
                      </div>

                      <p className="text-[var(--brand-ink)] font-medium">
                        "{sig.text}"
                      </p>

                      {sig.type === 'QUESTION' && (
                        <div className="flex items-center justify-between pt-1 border-t border-[var(--border)] text-[10px]">
                          <span className="text-[var(--brand-ink-soft)]">Status: {sig.isResolved ? '✅ Resolved' : '❓ Unresolved'}</span>
                          <button
                            onClick={() => {
                              sig.isResolved = !sig.isResolved;
                              setLiveSignals([...liveSignals]);
                            }}
                            className="text-[var(--brand-primary)] font-bold hover:underline"
                          >
                            Mark {sig.isResolved ? 'Unresolved' : 'Resolved'}
                          </button>
                        </div>
                      )}
                    </div>
                  ))
                ) : (
                  <div className="flex flex-col items-center justify-center py-16 text-center text-[var(--brand-ink-soft)] text-xs">
                    <Zap className="w-8 h-8 text-amber-500 opacity-40 mb-3" />
                    <p className="font-bold text-sm text-[var(--brand-ink)]">No Signals Detected Yet</p>
                    <p className="mt-1 max-w-xs text-xs">
                      Live commitments, decisions, and questions will appear here in real-time as speakers talk.
                    </p>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 4: Copilot AI Chat */}
          {activeRightTab === 'CHAT' && (
            <div className="flex-1 flex flex-col overflow-hidden">
              <div className="p-3 bg-[var(--bg-ice-blue)] border-b border-[var(--border)] flex flex-wrap gap-1.5 text-xs">
                <button
                  onClick={() => handleSendChatMessage('Draft a concise follow-up email summarizing the meeting and next steps.')}
                  className="px-2.5 py-1 rounded-full bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] text-[11px] font-bold border border-[var(--border)] transition-all shadow-xs"
                >
                  ✉️ Draft Email
                </button>
                <button
                  onClick={() => handleSendChatMessage('Summarize these action items as a Slack update.')}
                  className="px-2.5 py-1 rounded-full bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] text-[11px] font-bold border border-[var(--border)] transition-all shadow-xs"
                >
                  💬 Slack Format
                </button>
                <button
                  onClick={() => handleSendChatMessage('Were there any blockers or risks raised?')}
                  className="px-2.5 py-1 rounded-full bg-[var(--bg-surface)] hover:bg-[var(--bg-cream)] text-[var(--brand-primary)] text-[11px] font-bold border border-[var(--border)] transition-all shadow-xs"
                >
                  🚨 Extract Risks
                </button>
              </div>

              <div className="flex-1 p-4 overflow-y-auto space-y-4 text-xs">
                {chatMessages.map((msg, idx) => (
                  <div
                    key={idx}
                    className={`flex flex-col ${msg.role === 'user' ? 'items-end' : 'items-start'}`}
                  >
                    <div
                      className={`max-w-[88%] p-3.5 rounded-2xl ${
                        msg.role === 'user'
                          ? 'bg-[var(--brand-primary)] text-white rounded-br-none shadow-xs font-medium'
                          : 'bg-[var(--bg-cream)] text-[var(--brand-ink)] border border-[var(--border)] rounded-bl-none leading-relaxed shadow-xs'
                      }`}
                    >
                      <div className="whitespace-pre-wrap font-sans text-xs">
                        {msg.content}
                      </div>
                    </div>
                    <span className="text-[10px] text-[var(--brand-ink-soft)] mt-1 px-1 font-medium">{msg.time}</span>
                  </div>
                ))}

                {isChatSending && (
                  <div className="flex items-center gap-2 text-[var(--brand-primary)] font-bold text-xs">
                    <Sparkles className="w-3.5 h-3.5 animate-spin" />
                    <span>Analyzing notes & transcript...</span>
                  </div>
                )}
              </div>

              <div className="p-3 bg-[var(--bg-cream)] border-t border-[var(--border)] flex items-center gap-2">
                <input
                  type="text"
                  value={chatInput}
                  onChange={(e) => setChatInput(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleSendChatMessage()}
                  placeholder="Ask a question about this meeting..."
                  className="flex-1 bg-[var(--bg-surface)] border border-[var(--border)] rounded-lg px-3 py-2 text-xs text-[var(--brand-ink)] focus:outline-none focus:border-[var(--brand-primary)]"
                />
                <button
                  onClick={() => handleSendChatMessage()}
                  disabled={!chatInput.trim() || isChatSending}
                  className="p-2 rounded-lg bg-[var(--brand-primary)] hover:bg-[#07475F] text-white disabled:opacity-40 transition-all shadow-xs"
                >
                  <Send className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ---------------------------------------------------- */}
      {/* CUSTOM PROMPT TEMPLATE STUDIO MODAL                  */}
      {/* ---------------------------------------------------- */}
      {isTemplateStudioOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-[var(--bg-surface)] border border-[var(--border)] w-full max-w-lg rounded-2xl shadow-2xl overflow-hidden flex flex-col">
            <div className="flex items-center justify-between px-6 py-4 bg-[var(--bg-cream)] border-b border-[var(--border)]">
              <div className="flex items-center gap-2">
                <Sliders className="w-5 h-5 text-[var(--brand-primary)]" />
                <h3 className="font-bold text-sm text-[var(--brand-primary)]">Custom Template Studio</h3>
              </div>
              <button
                onClick={() => setIsTemplateStudioOpen(false)}
                className="text-xs font-bold text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)] px-2 py-1 bg-white border border-[var(--border)] rounded"
              >
                Close
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div>
                <label className="font-bold text-[var(--brand-ink)] block mb-1">Template Name / Title</label>
                <input
                  type="text"
                  value={newTmplLabel}
                  onChange={(e) => setNewTmplLabel(e.target.value)}
                  placeholder="e.g. Franchise QBR Review, Daily Fleet Standup"
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg p-2 text-xs text-[var(--brand-ink)] focus:outline-none focus:border-[var(--brand-primary)]"
                />
              </div>

              <div>
                <label className="font-bold text-[var(--brand-ink)] block mb-1">Description / Goal</label>
                <input
                  type="text"
                  value={newTmplDesc}
                  onChange={(e) => setNewTmplDesc(e.target.value)}
                  placeholder="e.g. Focus on depot performance, route SLAs, and client growth"
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg p-2 text-xs text-[var(--brand-ink)] focus:outline-none focus:border-[var(--brand-primary)]"
                />
              </div>

              <div>
                <label className="font-bold text-[var(--brand-ink)] block mb-1">Special AI Instructions (Optional)</label>
                <textarea
                  value={newTmplPrompt}
                  onChange={(e) => setNewTmplPrompt(e.target.value)}
                  placeholder="e.g. Always extract SLA metrics, express dates in DD/MM/YYYY, highlight unresolved fleet issues..."
                  className="w-full bg-[var(--bg-cream)] border border-[var(--border)] rounded-lg p-2 text-xs text-[var(--brand-ink)] focus:outline-none focus:border-[var(--brand-primary)] h-20 resize-none font-mono"
                />
              </div>

              <div className="pt-2 border-t border-[var(--border)]">
                <span className="font-bold text-[var(--brand-ink)] block mb-1.5">Existing Templates ({customTemplates.length})</span>
                <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
                  {customTemplates.map(t => (
                    <span key={t.id} className="px-2 py-1 bg-[var(--bg-cream)] rounded text-[10px] font-semibold border border-[var(--border)]">
                      {t.label}
                    </span>
                  ))}
                </div>
              </div>
            </div>

            <div className="px-6 py-3 bg-[var(--bg-cream)] border-t border-[var(--border)] flex justify-end gap-2">
              <button
                onClick={() => setIsTemplateStudioOpen(false)}
                className="px-3 py-1.5 bg-white border border-[var(--border)] rounded-lg text-xs font-bold text-[var(--brand-ink-soft)]"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveCustomTemplate}
                className="px-4 py-1.5 bg-[var(--brand-primary)] text-white rounded-lg text-xs font-bold hover:bg-[#07475F] transition-all shadow-xs"
              >
                Save Template
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------- */}
      {/* CALENDAR DIAGNOSTICS MODAL                           */}
      {/* ---------------------------------------------------- */}
      {showCalendarDetails && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-[var(--bg-surface)] border border-[var(--border)] w-full max-w-lg rounded-2xl shadow-2xl overflow-hidden flex flex-col">
            <div className="flex items-center justify-between px-6 py-4 bg-[var(--bg-cream)] border-b border-[var(--border)]">
              <div className="flex items-center gap-2">
                <CalendarCheck className="w-5 h-5 text-[var(--brand-primary)]" />
                <h3 className="font-bold text-sm text-[var(--brand-primary)]">Microsoft 365 Calendar Connection</h3>
              </div>
              <button
                onClick={() => setShowCalendarDetails(false)}
                className="text-xs font-bold text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)] px-2 py-1 bg-white border border-[var(--border)] rounded"
              >
                Close
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div className="p-3 bg-[var(--bg-ice-blue)] rounded-xl border border-[var(--brand-primary)]/20 space-y-1">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-[var(--brand-primary)]">Connection Status:</span>
                  <span className={`px-2 py-0.5 rounded-full font-bold text-[10px] ${
                    isCalendarConnected ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                  }`}>
                    {isCalendarConnected ? 'Active & Synced' : 'Action Needed'}
                  </span>
                </div>
                <p className="text-[var(--brand-ink-soft)] leading-relaxed">
                  {calendarStatusMessage}
                </p>
              </div>

              <div className="space-y-2">
                <span className="font-bold text-[var(--brand-ink)]">Mailbox Details:</span>
                <div className="bg-[var(--bg-cream)] p-3 rounded-lg border border-[var(--border)] font-mono text-[11px] space-y-1">
                  <div>User Account: <b>{userEmail}</b></div>
                  <div>Synced Events for Today: <b>{calendarEvents.length}</b></div>
                  <div>Sync Provider: <b>Microsoft Graph API</b></div>
                </div>
              </div>

              {calendarEvents.length > 0 && (
                <div className="space-y-2">
                  <span className="font-bold text-[var(--brand-ink)]">Today's Scheduled Events:</span>
                  <div className="space-y-1.5 max-h-48 overflow-y-auto">
                    {calendarEvents.map(evt => (
                      <div
                        key={evt.id}
                        onClick={() => {
                          associateCalendarEvent(evt);
                          setShowCalendarDetails(false);
                        }}
                        className="p-2.5 bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] rounded-lg border border-[var(--border)] cursor-pointer flex items-center justify-between transition-colors"
                      >
                        <div>
                          <div className="font-bold text-[var(--brand-ink)]">{evt.subject}</div>
                          <div className="text-[10px] text-[var(--brand-ink-soft)]">Organizer: {evt.organizer}</div>
                        </div>
                        <span className="font-mono text-[10px] font-bold text-[var(--brand-primary)]">{evt.startTime}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="px-6 py-3 bg-[var(--bg-cream)] border-t border-[var(--border)] flex justify-end gap-2">
              <button
                onClick={fetchCalendarData}
                disabled={isLoadingCalendar}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-[var(--brand-primary)] text-white rounded-lg text-xs font-bold hover:bg-[#07475F] transition-all"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isLoadingCalendar ? 'animate-spin' : ''}`} />
                <span>Re-Sync Calendar</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------- */}
      {/* CROSS-MEETING SEMANTIC SEARCH & TAGS MODAL (CMD+K)   */}
      {/* ---------------------------------------------------- */}
      {isSearchOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-[var(--bg-surface)] border border-[var(--border)] w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[80vh]">
            
            <div className="flex items-center gap-3 px-5 py-3.5 border-b border-[var(--border)] bg-[var(--bg-cream)]">
              <Search className="w-5 h-5 text-[var(--brand-primary)]" />
              <input
                type="text"
                autoFocus
                value={searchQuery}
                onChange={(e) => handleSemanticSearch(e.target.value)}
                placeholder="Search across all historical meetings, notes, and decisions..."
                className="flex-1 bg-transparent text-sm text-[var(--brand-ink)] font-semibold placeholder-[var(--brand-ink-soft)]/60 focus:outline-none"
              />
              <button
                onClick={() => setIsSearchOpen(false)}
                className="text-xs font-bold text-[var(--brand-ink-soft)] hover:text-[var(--brand-ink)] px-2 py-1 bg-[var(--bg-surface)] border border-[var(--border)] rounded"
              >
                ESC
              </button>
            </div>

            {/* Tag Filter Chips in Search */}
            <div className="px-5 py-2 bg-[var(--bg-ice-blue)] border-b border-[var(--border)] flex items-center gap-1.5 overflow-x-auto text-[11px]">
              <span className="font-bold text-[var(--brand-primary)]">Filter:</span>
              <button
                onClick={() => setSelectedSearchTag('ALL')}
                className={`px-2 py-0.5 rounded font-bold ${selectedSearchTag === 'ALL' ? 'bg-[var(--brand-primary)] text-white' : 'bg-white text-slate-700'}`}
              >
                All
              </button>
              {PRESET_TAGS.map(t => (
                <button
                  key={t}
                  onClick={() => setSelectedSearchTag(t)}
                  className={`px-2 py-0.5 rounded font-bold ${selectedSearchTag === t ? 'bg-[var(--brand-primary)] text-white' : 'bg-white text-slate-700'}`}
                >
                  #{t}
                </button>
              ))}
            </div>

            <div className="flex-1 p-5 overflow-y-auto space-y-3">
              {isSearching ? (
                <div className="text-center py-8 text-[var(--brand-primary)] font-bold text-xs">
                  Searching workspace meeting memory...
                </div>
              ) : searchResults.length > 0 ? (
                searchResults.map((res, i) => (
                  <div
                    key={i}
                    onClick={() => {
                      setSubject(res.subject);
                      setIsSearchOpen(false);
                      showToast(`Loaded "${res.subject}"`);
                    }}
                    className="p-3.5 bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] rounded-xl border border-[var(--border)] cursor-pointer transition-all space-y-1 shadow-xs"
                  >
                    <div className="flex items-center justify-between text-xs font-bold text-[var(--brand-ink)]">
                      <span>{res.subject}</span>
                      <span className="text-[10px] text-[var(--brand-primary)] bg-[var(--bg-ice-blue)] px-2 py-0.5 rounded border border-[var(--brand-primary)]/30 font-bold">
                        {res.matchSource} • Score {res.relevanceScore}
                      </span>
                    </div>
                    <p className="text-xs text-[var(--brand-ink-soft)] line-clamp-2">
                      {res.matchedSnippet}
                    </p>
                  </div>
                ))
              ) : searchQuery ? (
                <div className="text-center py-8 text-[var(--brand-ink-soft)] text-xs font-medium">
                  No meeting notes matched "{searchQuery}".
                </div>
              ) : (
                <div className="space-y-2">
                  <span className="text-[11px] font-black text-[var(--brand-primary)] uppercase tracking-wider">Recent Workspace Meetings</span>
                  {historicalMeetings.length > 0 ? (
                    historicalMeetings
                      .filter(m => selectedSearchTag === 'ALL' || (m.tags && m.tags.includes(selectedSearchTag)))
                      .slice(0, 6)
                      .map((m) => (
                        <div
                          key={m.meetingId}
                          onClick={() => {
                            setSubject(m.subject);
                            if (m.rawHumanNotes) setRawHumanNotes(m.rawHumanNotes);
                            if (m.transcriptSegments) setTranscriptSegments(m.transcriptSegments);
                            setIsSearchOpen(false);
                          }}
                          className="p-3 bg-[var(--bg-cream)] hover:bg-[var(--bg-ice-blue)] rounded-lg border border-[var(--border)] cursor-pointer text-xs flex items-center justify-between font-medium"
                        >
                          <span className="text-[var(--brand-ink)] font-bold">{m.subject}</span>
                          <span className="text-[11px] text-[var(--brand-ink-soft)]">{m.templatePreset || 'General'}</span>
                        </div>
                      ))
                  ) : (
                    <div className="text-center py-6 text-slate-400 text-xs">No saved meetings in history yet.</div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-[var(--brand-primary)] text-white border border-[var(--brand-primary)]/40 px-4 py-2.5 rounded-xl shadow-2xl text-xs font-bold flex items-center gap-2 animate-fade-in">
          <Sparkles className="w-4 h-4 text-[var(--brand-gold)]" />
          <span>{toastMessage}</span>
        </div>
      )}
    </div>
  );
}
