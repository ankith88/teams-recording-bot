import { useState, useRef, useEffect, useCallback } from 'react';

export interface TranscriptSegment {
  id: string;
  meetingId?: string;
  timestampFormatted: string;
  speakerName: string;
  speakerType: 'User' | 'Participant' | 'Unknown';
  text: string;
  confidence: number;
}

export type RecordingMode = 'TEAMS_SYSTEM' | 'MIC_ONLY';

export interface AudioRecorderState {
  isRecording: boolean;
  isPaused: boolean;
  durationSeconds: number;
  userVolume: number;        // 0.0 to 1.0 (Mic / You)
  participantVolume: number; // 0.0 to 1.0 (System / Tab)
  hasMicPermission: boolean;
  hasSystemAudio: boolean;
  activeSpeaker: string;
  recordingMode: RecordingMode;
}

export function useAudioRecorder(onTranscriptSegment?: (segment: TranscriptSegment) => void) {
  const [isRecording, setIsRecording] = useState(false);
  const [isPaused, setIsPaused] = useState(false);
  const [durationSeconds, setDurationSeconds] = useState(0);
  const [userVolume, setUserVolume] = useState(0);
  const [participantVolume, setParticipantVolume] = useState(0);
  const [hasMicPermission, setHasMicPermission] = useState(false);
  const [hasSystemAudio, setHasSystemAudio] = useState(false);
  const [activeSpeaker, setActiveSpeaker] = useState<string>('Participant');
  const [recordingMode, setRecordingMode] = useState<RecordingMode>('TEAMS_SYSTEM');

  // Recorded Audio Output & Playback State
  const [audioBlob, setAudioBlob] = useState<Blob | null>(null);
  const [audioBlobUrl, setAudioBlobUrl] = useState<string | null>(null);
  const [isPlayingAudio, setIsPlayingAudio] = useState(false);
  const [playbackCurrentTime, setPlaybackCurrentTime] = useState(0);
  const [playbackDuration, setPlaybackDuration] = useState(0);
  const [playbackRate, setPlaybackRateState] = useState(1.0);

  const timerRef = useRef<number | null>(null);
  const audioContextRef = useRef<AudioContext | null>(null);
  const micStreamRef = useRef<MediaStream | null>(null);
  const systemStreamRef = useRef<MediaStream | null>(null);
  const userAnalyserRef = useRef<AnalyserNode | null>(null);
  const participantAnalyserRef = useRef<AnalyserNode | null>(null);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const recordedChunksRef = useRef<Blob[]>([]);
  const animFrameRef = useRef<number | null>(null);
  const recognitionRef = useRef<any>(null);
  const activeSpeakerRef = useRef<string>('Participant');
  const audioElementRef = useRef<HTMLAudioElement | null>(null);

  // Keep ref synchronized with state
  useEffect(() => {
    activeSpeakerRef.current = activeSpeaker;
  }, [activeSpeaker]);

  // Clean up Audio element on unmount
  useEffect(() => {
    const audioEl = new Audio();
    audioElementRef.current = audioEl;

    const handleTimeUpdate = () => {
      setPlaybackCurrentTime(audioEl.currentTime);
    };
    const handleLoadedMetadata = () => {
      setPlaybackDuration(audioEl.duration || 0);
    };
    const handleEnded = () => {
      setIsPlayingAudio(false);
      setPlaybackCurrentTime(0);
    };
    const handlePlay = () => setIsPlayingAudio(true);
    const handlePause = () => setIsPlayingAudio(false);

    audioEl.addEventListener('timeupdate', handleTimeUpdate);
    audioEl.addEventListener('loadedmetadata', handleLoadedMetadata);
    audioEl.addEventListener('ended', handleEnded);
    audioEl.addEventListener('play', handlePlay);
    audioEl.addEventListener('pause', handlePause);

    return () => {
      audioEl.pause();
      audioEl.removeEventListener('timeupdate', handleTimeUpdate);
      audioEl.removeEventListener('loadedmetadata', handleLoadedMetadata);
      audioEl.removeEventListener('ended', handleEnded);
      audioEl.removeEventListener('play', handlePlay);
      audioEl.removeEventListener('pause', handlePause);
    };
  }, []);

  // Update audio element src when blob url changes
  useEffect(() => {
    if (audioElementRef.current && audioBlobUrl) {
      audioElementRef.current.src = audioBlobUrl;
      audioElementRef.current.load();
    }
  }, [audioBlobUrl]);

  // Audio level polling for visualizer
  const updateAudioLevels = useCallback(() => {
    if (!isRecording || isPaused) {
      setUserVolume(0);
      setParticipantVolume(0);
      return;
    }

    if (userAnalyserRef.current) {
      const dataArray = new Uint8Array(userAnalyserRef.current.frequencyBinCount);
      userAnalyserRef.current.getByteFrequencyData(dataArray);
      let sum = 0;
      for (let i = 0; i < dataArray.length; i++) sum += dataArray[i];
      const avg = sum / dataArray.length;
      setUserVolume(Math.min(1, avg / 128));
    }

    if (participantAnalyserRef.current) {
      const dataArray = new Uint8Array(participantAnalyserRef.current.frequencyBinCount);
      participantAnalyserRef.current.getByteFrequencyData(dataArray);
      let sum = 0;
      for (let i = 0; i < dataArray.length; i++) sum += dataArray[i];
      const avg = sum / dataArray.length;
      setParticipantVolume(Math.min(1, avg / 128));
    }

    animFrameRef.current = requestAnimationFrame(updateAudioLevels);
  }, [isRecording, isPaused]);

  useEffect(() => {
    if (isRecording && !isPaused) {
      animFrameRef.current = requestAnimationFrame(updateAudioLevels);
    } else {
      if (animFrameRef.current) cancelAnimationFrame(animFrameRef.current);
      setUserVolume(0);
      setParticipantVolume(0);
    }
    return () => {
      if (animFrameRef.current) cancelAnimationFrame(animFrameRef.current);
    };
  }, [isRecording, isPaused, updateAudioLevels]);

  // Duration timer
  useEffect(() => {
    if (isRecording && !isPaused) {
      timerRef.current = window.setInterval(() => {
        setDurationSeconds(prev => prev + 1);
      }, 1000);
    } else {
      if (timerRef.current) clearInterval(timerRef.current);
    }
    return () => {
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, [isRecording, isPaused]);

  // Start Dual-Channel or Mic-Only Audio Capture
  const startRecording = async (options?: { 
    mode?: RecordingMode;
    includeSystemAudio?: boolean; 
    defaultSpeaker?: string; 
    userDisplayName?: string 
  }) => {
    try {
      const currentMode = options?.mode || recordingMode;
      setRecordingMode(currentMode);
      recordedChunksRef.current = [];

      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      audioContextRef.current = audioCtx;
      if (audioCtx.state === 'suspended') {
        await audioCtx.resume();
      }

      // 1. Microphone capture (You)
      const micStream = await navigator.mediaDevices.getUserMedia({
        audio: {
          echoCancellation: true,
          noiseSuppression: true,
          autoGainControl: true,
        },
      });
      micStreamRef.current = micStream;
      setHasMicPermission(true);

      const micSource = audioCtx.createMediaStreamSource(micStream);
      const userAnalyser = audioCtx.createAnalyser();
      userAnalyser.fftSize = 256;
      micSource.connect(userAnalyser);
      userAnalyserRef.current = userAnalyser;

      // 2. System / Tab audio capture (Meeting Participants) - Only in TEAMS_SYSTEM mode
      let mixedDestination = audioCtx.createMediaStreamDestination();
      micSource.connect(mixedDestination);

      if (currentMode === 'TEAMS_SYSTEM' && options?.includeSystemAudio !== false) {
        try {
          const sysStream = await navigator.mediaDevices.getDisplayMedia({
            video: true,
            audio: true,
          });
          
          const audioTracks = sysStream.getAudioTracks();
          if (audioTracks.length > 0) {
            systemStreamRef.current = sysStream;
            setHasSystemAudio(true);
            const sysAudioOnlyStream = new MediaStream(audioTracks);
            const sysSource = audioCtx.createMediaStreamSource(sysAudioOnlyStream);
            const partAnalyser = audioCtx.createAnalyser();
            partAnalyser.fftSize = 256;
            sysSource.connect(partAnalyser);
            participantAnalyserRef.current = partAnalyser;
            sysSource.connect(mixedDestination);

            // Handle user stopping screen share from browser banner
            sysStream.getVideoTracks().forEach(track => {
              track.onended = () => {
                setHasSystemAudio(false);
              };
            });
          }
        } catch (sysErr) {
          console.log('[useAudioRecorder] System audio prompt bypassed or dismissed:', sysErr);
          setHasSystemAudio(false);
        }
      } else {
        setHasSystemAudio(false);
      }

      // 3. MediaRecorder for mixed audio
      const mimeType = MediaRecorder.isTypeSupported('audio/webm;codecs=opus')
        ? 'audio/webm;codecs=opus'
        : MediaRecorder.isTypeSupported('audio/mp4')
        ? 'audio/mp4'
        : '';

      const mediaRecorder = new MediaRecorder(mixedDestination.stream, mimeType ? { mimeType } : undefined);
      mediaRecorderRef.current = mediaRecorder;

      mediaRecorder.ondataavailable = (e) => {
        if (e.data && e.data.size > 0) {
          recordedChunksRef.current.push(e.data);
        }
      };

      mediaRecorder.start(1000); // 1-second slices

      // 4. Initialize Web Speech Recognition for live transcription stream
      const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
      if (SpeechRecognition) {
        const recognition = new SpeechRecognition();
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.lang = 'en-US';

        recognition.onresult = (event: any) => {
          for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
              const text = event.results[i][0].transcript.trim();
              if (text && onTranscriptSegment) {
                const nowSec = durationSeconds;
                const hours = Math.floor(nowSec / 3600).toString().padStart(2, '0');
                const mins = Math.floor((nowSec % 3600) / 60).toString().padStart(2, '0');
                const secs = (nowSec % 60).toString().padStart(2, '0');

                const isUserVoice = currentMode === 'MIC_ONLY' || (userVolume > 0.08 && participantVolume < 0.05);
                const speakerName = isUserVoice ? (options?.userDisplayName || 'You') : activeSpeakerRef.current;
                const speakerType = isUserVoice ? 'User' : 'Participant';

                onTranscriptSegment({
                  id: 'seg-' + Date.now() + '-' + Math.random().toString(36).substring(2, 6),
                  timestampFormatted: `${hours}:${mins}:${secs}`,
                  speakerName,
                  speakerType,
                  text,
                  confidence: event.results[i][0].confidence || 0.95
                });
              }
            }
          }
        };

        recognition.onerror = (err: any) => {
          console.warn('[useAudioRecorder] Speech recognition notice:', err.error);
        };

        recognition.onend = () => {
          if (isRecording && !isPaused && recognitionRef.current) {
            try {
              recognition.start();
            } catch (e) {}
          }
        };

        try {
          recognition.start();
          recognitionRef.current = recognition;
        } catch (recErr) {
          console.warn('[useAudioRecorder] Could not start speech recognition:', recErr);
        }
      }

      setIsRecording(true);
      setIsPaused(false);
      setDurationSeconds(0);
      if (options?.defaultSpeaker) {
        setActiveSpeaker(options.defaultSpeaker);
      }
      return true;
    } catch (err) {
      console.error('[useAudioRecorder] Error starting audio recording:', err);
      setIsRecording(false);
      return false;
    }
  };

  const pauseRecording = () => {
    if (!isRecording || isPaused) return;
    setIsPaused(true);
    if (mediaRecorderRef.current && mediaRecorderRef.current.state === 'recording') {
      mediaRecorderRef.current.pause();
    }
    if (recognitionRef.current) {
      try { recognitionRef.current.stop(); } catch (e) {}
    }
  };

  const resumeRecording = () => {
    if (!isRecording || !isPaused) return;
    setIsPaused(false);
    if (mediaRecorderRef.current && mediaRecorderRef.current.state === 'paused') {
      mediaRecorderRef.current.resume();
    }
    if (recognitionRef.current) {
      try { recognitionRef.current.start(); } catch (e) {}
    }
  };

  const stopRecording = () => {
    setIsRecording(false);
    setIsPaused(false);

    if (timerRef.current) clearInterval(timerRef.current);
    if (animFrameRef.current) cancelAnimationFrame(animFrameRef.current);

    if (recognitionRef.current) {
      try {
        recognitionRef.current.stop();
      } catch (e) {}
      recognitionRef.current = null;
    }

    if (mediaRecorderRef.current) {
      try {
        if (mediaRecorderRef.current.state !== 'inactive') {
          mediaRecorderRef.current.stop();
        }
      } catch (e) {}
      mediaRecorderRef.current = null;
    }

    // Assemble final audio blob
    setTimeout(() => {
      if (recordedChunksRef.current.length > 0) {
        const type = recordedChunksRef.current[0].type || 'audio/webm';
        const blob = new Blob(recordedChunksRef.current, { type });
        setAudioBlob(blob);
        const url = URL.createObjectURL(blob);
        setAudioBlobUrl(url);
        setPlaybackDuration(durationSeconds);
      }
    }, 200);

    if (micStreamRef.current) {
      micStreamRef.current.getTracks().forEach(t => t.stop());
      micStreamRef.current = null;
    }

    if (systemStreamRef.current) {
      systemStreamRef.current.getTracks().forEach(t => t.stop());
      systemStreamRef.current = null;
    }

    if (audioContextRef.current) {
      audioContextRef.current.close().catch(() => {});
      audioContextRef.current = null;
    }

    setUserVolume(0);
    setParticipantVolume(0);
    setHasSystemAudio(false);
  };

  // Synchronized Audio Playback Handlers
  const playAudio = () => {
    if (audioElementRef.current && audioBlobUrl) {
      audioElementRef.current.play().catch(e => console.warn('Audio play notice:', e));
    }
  };

  const pauseAudio = () => {
    if (audioElementRef.current) {
      audioElementRef.current.pause();
    }
  };

  const seekTo = (seconds: number) => {
    if (audioElementRef.current && audioBlobUrl) {
      audioElementRef.current.currentTime = Math.max(0, Math.min(seconds, audioElementRef.current.duration || 99999));
      if (!isPlayingAudio) {
        audioElementRef.current.play().catch(() => {});
      }
    }
  };

  const setPlaybackRate = (rate: number) => {
    setPlaybackRateState(rate);
    if (audioElementRef.current) {
      audioElementRef.current.playbackRate = rate;
    }
  };

  const downloadAudio = (filename: string = 'meeting-recording.webm') => {
    if (!audioBlob) return;
    const a = document.createElement('a');
    a.href = audioBlobUrl || URL.createObjectURL(audioBlob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return {
    isRecording,
    isPaused,
    durationSeconds,
    userVolume,
    participantVolume,
    hasMicPermission,
    hasSystemAudio,
    activeSpeaker,
    setActiveSpeaker,
    recordingMode,
    setRecordingMode,
    startRecording,
    pauseRecording,
    resumeRecording,
    stopRecording,
    // Playback state and actions
    audioBlob,
    audioBlobUrl,
    isPlayingAudio,
    playbackCurrentTime,
    playbackDuration,
    playbackRate,
    playAudio,
    pauseAudio,
    seekTo,
    setPlaybackRate,
    downloadAudio
  };
}
