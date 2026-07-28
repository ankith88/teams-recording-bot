import React from 'react';
import ReactDOM from 'react-dom/client';
import TeamsRecorderTab from '../TeamsRecorderTab';
import '../index.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <div className="min-h-screen bg-slate-950 py-10 px-4">
      <TeamsRecorderTab />
    </div>
  </React.StrictMode>
);
