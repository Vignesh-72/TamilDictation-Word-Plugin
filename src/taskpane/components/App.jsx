import React, { useState, useRef, useEffect } from 'react';

// Import logo via webpack so the path is resolved correctly in all environments
import logoUrl from '../../../assets/logo.jpg';

/* global Office, Word */

// ─────────────────────────────────────────────────────────────────────────────
// Helper: escape regex special characters to prevent ReDoS (Issue #14)
// ─────────────────────────────────────────────────────────────────────────────
const escapeRegex = (str) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// ─────────────────────────────────────────────────────────────────────────────
// Helper: save preferences (Office document settings + localStorage fallback)
// ─────────────────────────────────────────────────────────────────────────────
const savePrefs = (prefs) => {
  const json = JSON.stringify(prefs);
  // Primary: Office document settings (roam with the document)
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.set('linguist_prefs', json);
      Office.context.document.settings.saveAsync();
    }
  } catch (_) { /* ignore */ }
  // Fallback: localStorage
  try { localStorage.setItem('linguist_prefs', json); } catch (_) { /* ignore */ }
};

// ─────────────────────────────────────────────────────────────────────────────
// Helper: load preferences (Office document settings → localStorage → defaults)
// ─────────────────────────────────────────────────────────────────────────────
const loadPrefs = (callback) => {
  const applyRaw = (json) => {
    try {
      const parsed = JSON.parse(json);
      callback(parsed);
    } catch (_) { /* corrupted — use defaults */ }
  };

  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.settings.refreshAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const saved = Office.context.document.settings.get('linguist_prefs');
          if (saved) { applyRaw(saved); return; }
        }
        // Fallback to localStorage
        try {
          const local = localStorage.getItem('linguist_prefs');
          if (local) applyRaw(local);
        } catch (_) { /* ignore */ }
      });
    } else {
      const local = localStorage.getItem('linguist_prefs');
      if (local) applyRaw(local);
    }
  } catch (_) {
    try {
      const local = localStorage.getItem('linguist_prefs');
      if (local) applyRaw(local);
    } catch (_) { /* ignore */ }
  }
};

const App = () => {
  const [dialog, setDialog]       = useState(null);
  const [activeTab, setActiveTab] = useState('main');
  const [uiLang, setUiLang]       = useState('en');
  const [errorMsg, setErrorMsg]   = useState(null);
  const [statusMsg, setStatusMsg] = useState(null);
  const isWriting = useRef(false);

  const [smartPunc, setSmartPunc]   = useState(true);
  const [transMode, setTransMode]   = useState('ta');
  // Default to 'Latha' — a Tamil font shipped with every Windows install (Issue #16)
  const [fontFamily, setFontFamily] = useState('Latha');

  const [commands, setCommands] = useState({
    'முற்றுப்புள்ளி': '.',
    'கேள்விக்குறி': '?',
    'புதிய வரி': '\n',
  });
  const [newCmd, setNewCmd] = useState({ voice: '', result: '' });

  // ── Refs for closures inside Office callbacks ──
  const uiLangRef   = useRef(uiLang);
  const transModeRef = useRef(transMode);
  const fontFamilyRef = useRef(fontFamily);
  const smartPuncRef  = useRef(smartPunc);
  const commandsRef   = useRef(commands);

  useEffect(() => { uiLangRef.current    = uiLang;    }, [uiLang]);
  useEffect(() => { transModeRef.current = transMode; }, [transMode]);
  useEffect(() => { fontFamilyRef.current = fontFamily; }, [fontFamily]);
  useEffect(() => { smartPuncRef.current  = smartPunc;  }, [smartPunc]);
  useEffect(() => { commandsRef.current   = commands;   }, [commands]);

  // ── LOCALIZATION ──
  const locales = {
    en: {
      title:      'Tamil Dictation',
      control:    'Control',
      settings:   'Settings',
      startBtn:   'START DICTATION',
      stopBtn:    'STOP SESSION',
      smartPunc:  'Smart Punctuation',
      smartDesc:  'Automatically convert spoken commands to symbols.',
      transEngine:'Dictation Mode',
      ta_ta:      'Tamil (Direct Dictation)',
      en_en:      'English (Direct Dictation)',
      ta_en:      'Tamil ➔ English (needs internet)',
      en_ta:      'English ➔ Tamil (needs internet)',
      fontStyle:  'Font Style',
      voiceMaps:  'Voice Shortcuts',
      spokenWord: 'When I say...',
      outputText: 'Type this...',
      addCmd:     'Add Shortcut',
      noInternet: 'Translation requires internet connection. Original text inserted.',
      writeErr:   'Could not write to document. Is it read-only?',
      dialogErr:  'Could not open dictation dialog. Please try again.',
    },
    ta: {
      title:      'தமிழ் டிக்டேஷன்',
      control:    'கட்டுப்பாடு',
      settings:   'அமைப்புகள்',
      startBtn:   'டிக்டேஷனைத் தொடங்கு',
      stopBtn:    'அமர்வை நிறுத்து',
      smartPunc:  'ஸ்மார்ட் நிறுத்தற்குறிகள்',
      smartDesc:  'பேசும் வார்த்தைகளை குறியீடுகளாக தானாகவே மாற்றும்.',
      transEngine:'டிக்டேஷன் முறை',
      ta_ta:      'தமிழ் (நேரடி தட்டச்சு)',
      en_en:      'ஆங்கிலம் (நேரடி தட்டச்சு)',
      ta_en:      'தமிழ் ➔ ஆங்கிலம் (இணையம் தேவை)',
      en_ta:      'ஆங்கிலம் ➔ தமிழ் (இணையம் தேவை)',
      fontStyle:  'எழுத்துரு (Font)',
      voiceMaps:  'குரல் சுருக்குவழிகள்',
      spokenWord: 'நான் சொல்லும் வார்த்தை...',
      outputText: 'வர வேண்டிய குறி...',
      addCmd:     'சுருக்குவழியைச் சேர்',
      noInternet: 'மொழிபெயர்ப்புக்கு இணையம் தேவை. அசல் உரை சேர்க்கப்பட்டது.',
      writeErr:   'ஆவணத்தில் எழுத முடியவில்லை. படிக்க மட்டுமே அனுமதியா?',
      dialogErr:  'டிக்டேஷன் சாளரம் திறக்கவில்லை. மீண்டும் முயற்சிக்கவும்.',
    },
  };

  const t = locales[uiLang];

  // ── Load saved preferences on mount ──
  useEffect(() => {
    loadPrefs((parsed) => {
      if (parsed.commands)    setCommands(parsed.commands);
      if (parsed.smartPunc !== undefined) setSmartPunc(parsed.smartPunc);
      if (parsed.uiLang)      setUiLang(parsed.uiLang);
      if (parsed.transMode)   setTransMode(parsed.transMode);
      if (parsed.fontFamily)  setFontFamily(parsed.fontFamily);
    });
  }, []);

  // ── Save preferences on any change ──
  useEffect(() => {
    savePrefs({ commands, smartPunc, uiLang, transMode, fontFamily });
  }, [commands, smartPunc, uiLang, transMode, fontFamily]);

  // ── Flash helpers ──
  const showError = (msg) => {
    setErrorMsg(msg);
    setTimeout(() => setErrorMsg(null), 5000);
  };
  const showStatus = (msg) => {
    setStatusMsg(msg);
    setTimeout(() => setStatusMsg(null), 4000);
  };

  // ────────────────────────────────────────────────────────────────────────────
  // Translation (Issues #3, #15)
  // Requires internet; gracefully falls back to original text when offline.
  // The ta/en "direct" modes skip this entirely — fully offline.
  // ────────────────────────────────────────────────────────────────────────────
  const translate = async (text, mode) => {
    if (mode === 'ta' || mode === 'en' || !text) return text;

    try {
      const source = mode === 'ta-en' ? 'ta' : 'en';
      const target = mode === 'ta-en' ? 'en' : 'ta';
      const url    = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${source}&tl=${target}&dt=t&q=${encodeURIComponent(text)}`;

      // 5-second timeout so we don't block forever offline
      const controller = new AbortController();
      const timeout    = setTimeout(() => controller.abort(), 5000);
      const response   = await fetch(url, { signal: controller.signal });
      clearTimeout(timeout);

      const data = await response.json();
      if (data && data[0]) {
        return data[0].map((item) => item[0]).join('');
      }
      return text;
    } catch (_) {
      // Network unavailable or timed out — insert original text and warn user
      showStatus(locales[uiLangRef.current].noInternet);
      return text;
    }
  };

  const deleteCommand = (key) => {
    const updated = { ...commands };
    delete updated[key];
    setCommands(updated);
  };

  // ────────────────────────────────────────────────────────────────────────────
  // Word Document Injection (Issues #5, #6, #15, #17, #24)
  // ────────────────────────────────────────────────────────────────────────────
  const handleInjection = async (raw, isFinal) => {
    if (isWriting.current && !isFinal) return;

    let processed = raw;

    if (isFinal) {
      // Apply smart punctuation substitutions with regex-escaped keys (Issue #14)
      if (smartPuncRef.current) {
        Object.keys(commandsRef.current).forEach((k) => {
          processed = processed.replace(
            new RegExp(escapeRegex(k), 'g'),
            commandsRef.current[k]
          );
        });
      }
      // Translate (may be a no-op for offline modes)
      processed = await translate(processed, transModeRef.current);
    }

    // FIX #5: Set the guard flag BEFORE entering the async Word.run context
    // to prevent a second concurrent call from slipping through.
    isWriting.current = true;

    try {
      await Word.run(async (context) => {
        const currentFont = fontFamilyRef.current;
        const tag = 'TamilAnchor';

        // Locate or create our dedicated content control anchor
        const ccs = context.document.contentControls.getByTag(tag);
        ccs.load('items');
        await context.sync();

        let cc;
        if (ccs.items.length === 0) {
          // FIX #6: Load and lock the current selection immediately at the
          // start of the Word.run context (before any async gaps), so it
          // reflects the cursor position at the time the message was received.
          const selection = context.document.getSelection();
          cc = selection.insertContentControl();
        } else {
          cc = ccs.items[0];
        }

        cc.tag        = tag;
        cc.appearance = 'Hidden';

        // ── Newline command (Issue #15) ──
        // Handle before translation so we're working with the post-processed text.
        if (isFinal && processed.includes('\n')) {
          cc.delete(true);
          await context.sync(); // Sync the delete before inserting the break
          context.document.getSelection().insertBreak(Word.BreakType.paragraph);
        } else {
          cc.insertText(processed, 'Replace');

          if (isFinal) {
            // FIX #17: Only set paragraph formatting on FINAL results.
            const range = cc.getRange();
            range.font.name = currentFont;
            range.paragraphs.getFirst().alignment    = 'Left';
            range.paragraphs.getFirst().readingOrder = 'LeftToRight';

            // Collapse the content control into static text, append a space,
            // then move the cursor to after the space.
            const finalRange = cc.getRange();
            cc.delete(true);
            const space = finalRange.insertText(' ', 'End');
            space.font.name = currentFont;
            space.select('End');
          } else {
            // Interim: only update text and font — no paragraph reformatting.
            cc.getRange().font.name = currentFont;
          }
        }

        await context.sync();
      });
    } catch (e) {
      // Surface the error to the user instead of silently swallowing it (#24)
      showError(locales[uiLangRef.current].writeErr);
    } finally {
      isWriting.current = false;
    }
  };

  // ────────────────────────────────────────────────────────────────────────────
  // Session management
  // ────────────────────────────────────────────────────────────────────────────
  const stopSession = () => {
    if (dialog) { try { dialog.close(); } catch (_) {} }
    setDialog(null);
  };

  const startSession = () => {
    let listenLang = 'ta-IN';
    if (transMode === 'en' || transMode === 'en-ta') {
      listenLang = 'en-IN';
    }

    // FIX #2: Use window.location.origin (resolves correctly in both dev and
    // prod because the taskpane is always served from the plugin's own domain).
    const url = `${window.location.origin}/assets/dictation.html?lang=${listenLang}&font=${encodeURIComponent(fontFamily)}`;

    setErrorMsg(null);

    Office.context.ui.displayDialogAsync(url, { height: 45, width: 35 }, (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        showError(locales[uiLangRef.current].dialogErr);
        return;
      }

      setDialog(res.value);

      res.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        try {
          const d = JSON.parse(arg.message);
          handleInjection(d.text, d.isFinal);
        } catch (_) { /* ignore malformed messages */ }
      });

      res.value.addEventHandler(Office.EventType.DialogEventReceived, stopSession);
    });
  };

  // ────────────────────────────────────────────────────────────────────────────
  // Render
  // ────────────────────────────────────────────────────────────────────────────
  return (
    <div style={styles.container}>
      <div style={styles.card}>

        {/* ── Tab navigation + language switcher ── */}
        <div style={styles.navContainer}>
          <div style={styles.nav}>
            <button onClick={() => setActiveTab('main')} style={activeTab === 'main' ? styles.tabActive : styles.tab}>{t.control}</button>
            <button onClick={() => setActiveTab('set')}  style={activeTab === 'set'  ? styles.tabActive : styles.tab}>{t.settings}</button>
          </div>
          <div style={styles.langSwitcher}>
            <button onClick={() => setUiLang('en')} style={uiLang === 'en' ? styles.langBtnActive : styles.langBtn}>EN</button>
            <button onClick={() => setUiLang('ta')} style={uiLang === 'ta' ? styles.langBtnActive : styles.langBtn}>தமிழ்</button>
          </div>
        </div>

        {/* ── Inline notification banners (#24) ── */}
        {errorMsg && (
          <div style={styles.errorBanner}>
            <span style={{ fontSize: '14px' }}>⚠️</span>
            <span>{errorMsg}</span>
          </div>
        )}
        {statusMsg && (
          <div style={styles.statusBanner}>
            <span style={{ fontSize: '14px' }}>ℹ️</span>
            <span>{statusMsg}</span>
          </div>
        )}

        {/* ── Main (Control) Tab ── */}
        {activeTab === 'main' ? (
          <div style={styles.scrollArea}>
            <div style={styles.titleRow}>
              {/* FIX #11: Use webpack-imported asset URL instead of root-relative path */}
              <img src={logoUrl} alt="Tamil Dictation Logo" style={styles.logo} />
              <h2 style={styles.title}>{t.title}</h2>
            </div>

            <button
              onClick={dialog ? stopSession : startSession}
              style={{ ...styles.mainBtn, background: dialog ? '#d13438' : '#0078d4' }}
            >
              {dialog ? t.stopBtn : t.startBtn}
            </button>

            <div style={styles.toggleBox}>
              <div style={styles.row}>
                <span style={{ fontWeight: '600', color: '#323130' }}>{t.smartPunc}</span>
                <input
                  type="checkbox"
                  checked={smartPunc}
                  onChange={() => setSmartPunc(!smartPunc)}
                  style={{ cursor: 'pointer', flexShrink: 0 }}
                />
              </div>
              <p style={styles.descText}>{t.smartDesc}</p>
            </div>
          </div>

        ) : (

          /* ── Settings Tab ── */
          <div style={styles.scrollArea}>
            <h4 style={styles.label}>{t.transEngine}</h4>
            <select value={transMode} onChange={(e) => setTransMode(e.target.value)} style={styles.input}>
              <option value="ta">{t.ta_ta}</option>
              <option value="en">{t.en_en}</option>
              <option value="ta-en">{t.ta_en}</option>
              <option value="en-ta">{t.en_ta}</option>
            </select>

            <h4 style={styles.label}>{t.fontStyle}</h4>
            <select value={fontFamily} onChange={(e) => setFontFamily(e.target.value)} style={styles.input}>
              {/* System-available Tamil fonts first; TAU-Marutham is a custom install */}
              <option value="Latha">Latha (System Default)</option>
              <option value="Vijaya">Vijaya</option>
              <option value="Kartika">Kartika</option>
              <option value="TAU-Marutham">TAU-Marutham (custom install required)</option>
            </select>

            <h4 style={styles.label}>{t.voiceMaps}</h4>
            <div style={styles.inputWrap}>
              <input
                placeholder={t.spokenWord}
                value={newCmd.voice}
                onChange={(e) => setNewCmd({ ...newCmd, voice: e.target.value })}
                style={styles.flexInput}
              />
              <input
                placeholder={t.outputText}
                value={newCmd.result}
                onChange={(e) => setNewCmd({ ...newCmd, result: e.target.value })}
                style={styles.flexInput}
              />
            </div>

            <button
              onClick={() => {
                if (newCmd.voice && newCmd.result) {
                  setCommands({ ...commands, [newCmd.voice]: newCmd.result });
                  setNewCmd({ voice: '', result: '' });
                }
              }}
              style={styles.addBtn}
            >
              {t.addCmd}
            </button>

            <div style={styles.list}>
              {Object.keys(commands).map((k) => (
                <div key={k} style={styles.listItem}>
                  <div style={{ display: 'flex', gap: '8px', flex: 1, alignItems: 'center', minWidth: 0 }}>
                    <span style={{ color: '#0078d4', fontWeight: '600', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{k}</span>
                    <b style={{ color: '#edebe9', flexShrink: 0 }}>➔</b>
                    <span style={{ color: '#323130', wordBreak: 'break-all' }}>
                      {commands[k] === '\n' ? '↵' : commands[k]}
                    </span>
                  </div>
                  <button onClick={() => deleteCommand(k)} style={styles.delBtn}>✕</button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// Styles
// ─────────────────────────────────────────────────────────────────────────────
const styles = {
  container:   { padding: '12px', backgroundColor: '#f3f2f1', height: '100vh', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', fontFamily: '"Segoe UI", system-ui, sans-serif' },
  card:        { flex: 1, backgroundColor: '#fff', padding: '16px', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', display: 'flex', flexDirection: 'column', border: '1px solid #edebe9', boxSizing: 'border-box', overflow: 'hidden' },
  navContainer:{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px', borderBottom: '1px solid #edebe9', paddingBottom: '12px', marginBottom: '16px' },
  nav:         { display: 'flex', gap: '4px', flexWrap: 'wrap' },
  tab:         { padding: '8px 12px', background: 'none', border: 'none', color: '#605e5c', cursor: 'pointer', fontWeight: '600', fontSize: '13px', borderRadius: '4px' },
  tabActive:   { padding: '8px 12px', background: '#f3f2f1', border: 'none', color: '#0078d4', fontWeight: 'bold', fontSize: '13px', borderRadius: '4px' },
  langSwitcher:{ display: 'flex', backgroundColor: '#f3f2f1', borderRadius: '4px', padding: '2px', alignItems: 'center', flexShrink: 0 },
  langBtn:     { border: 'none', background: 'transparent', padding: '4px 8px', fontSize: '12px', cursor: 'pointer', color: '#605e5c', fontWeight: '600', borderRadius: '2px', transition: 'all 0.2s' },
  langBtnActive:{ border: 'none', background: '#ffffff', padding: '4px 8px', fontSize: '12px', cursor: 'pointer', color: '#0078d4', fontWeight: 'bold', borderRadius: '2px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', transition: 'all 0.2s' },
  // Notification banners
  errorBanner: { display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#fde7e9', border: '1px solid #f1707b', borderRadius: '4px', padding: '8px 12px', marginBottom: '10px', fontSize: '12px', color: '#a4262c', flexShrink: 0 },
  statusBanner:{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#dff6dd', border: '1px solid #92c353', borderRadius: '4px', padding: '8px 12px', marginBottom: '10px', fontSize: '12px', color: '#107c10', flexShrink: 0 },
  scrollArea:  { flex: 1, overflowY: 'auto', paddingRight: '4px', animation: 'fadeIn 0.3s ease-in', display: 'flex', flexDirection: 'column' },
  titleRow:    { display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '20px' },
  logo:        { width: '32px', height: '32px', borderRadius: '6px', objectFit: 'cover', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' },
  title:       { fontSize: '22px', color: '#2b579a', margin: '0', fontWeight: '700', flexShrink: 0 },
  mainBtn:     { width: '100%', padding: '14px', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', letterSpacing: '0.5px', transition: 'background 0.2s', boxShadow: '0 2px 4px rgba(0,0,0,0.1)', flexShrink: 0 },
  toggleBox:   { marginTop: '20px', borderTop: '1px solid #edebe9', paddingTop: '16px', flexShrink: 0 },
  row:         { display: 'flex', justifyContent: 'space-between', marginBottom: '4px', fontSize: '13px', alignItems: 'center', gap: '10px' },
  descText:    { fontSize: '11px', color: '#888', marginTop: '0', lineHeight: '1.4' },
  label:       { fontSize: '11px', color: '#a19f9d', textTransform: 'uppercase', marginBottom: '8px', marginTop: '8px', letterSpacing: '0.5px', fontWeight: 'bold', flexShrink: 0 },
  input:       { width: '100%', padding: '10px', marginBottom: '12px', boxSizing: 'border-box', border: '1px solid #d2d0ce', borderRadius: '4px', fontSize: '13px', color: '#323130', flexShrink: 0 },
  inputWrap:   { display: 'flex', flexWrap: 'wrap', gap: '8px', marginBottom: '12px', flexShrink: 0 },
  flexInput:   { flex: '1 1 100px', padding: '10px', boxSizing: 'border-box', border: '1px solid #d2d0ce', borderRadius: '4px', fontSize: '13px', color: '#323130', minWidth: '100px' },
  addBtn:      { width: '100%', padding: '12px', backgroundColor: '#f3f2f1', color: '#0078d4', border: '1px solid #0078d4', cursor: 'pointer', borderRadius: '4px', fontWeight: '600', fontSize: '13px', flexShrink: 0 },
  list:        { marginTop: '16px', borderTop: '1px solid #edebe9', paddingTop: '8px', flexShrink: 0 },
  listItem:    { display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', padding: '10px 0', borderBottom: '1px solid #f3f2f1', gap: '8px' },
  delBtn:      { border: 'none', background: '#fde7e9', color: '#a4262c', cursor: 'pointer', padding: '6px 10px', borderRadius: '4px', fontWeight: 'bold', fontSize: '12px', flexShrink: 0 },
};

// Inject global animations once (not on every render)
const _css = document.createElement('style');
_css.innerHTML = `
  @keyframes fadeIn { from { opacity: 0; transform: translateY(2px); } to { opacity: 1; transform: translateY(0); } }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: #c8c6c4; border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: #a19f9d; }
`;
if (!document.getElementById('tamil-dictation-styles')) {
  _css.id = 'tamil-dictation-styles';
  document.head.appendChild(_css);
}

export default App;