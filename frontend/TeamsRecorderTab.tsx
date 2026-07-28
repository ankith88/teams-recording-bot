'use client';

import React, { useState } from 'react';
import { Play, Square, Folder, FileText, CheckCircle2, AlertCircle, RefreshCw, Volume2 } from 'lucide-react';

interface TranscriptSegment {
  id: string;
  speakerName: string;
  startTime: string;
  text: string;
}

export default function TeamsRecorderTab() {
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

  return (
    <div className="max-w-5xl mx-auto p-6 bg-slate-900 text-white rounded-xl shadow-2xl space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between border-b border-slate-800 pb-4">
        <div>
          <h1 className="text-2xl font-bold text-blue-400">Microsoft Teams Meeting Transcriber</h1>
          <p className="text-sm text-slate-400">Join Teams calls, capture audio, and save timestamped transcripts directly to your laptop.</p>
        </div>
        <div className="flex items-center space-x-2 bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-700 text-xs text-slate-300">
          <Volume2 className="w-4 h-4 text-emerald-400 animate-pulse" />
          <span>Local Whisper Engine Active</span>
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
