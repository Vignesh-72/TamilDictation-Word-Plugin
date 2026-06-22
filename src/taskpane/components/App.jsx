import React, { useState, useRef, useEffect } from 'react';

const App = () => {
  const [dialog, setDialog] = useState(null);
  const [activeTab, setActiveTab] = useState('main');
  const [uiLang, setUiLang] = useState('en'); 
  const isWriting = useRef(false);

  const [smartPunc, setSmartPunc] = useState(true);
  const [transMode, setTransMode] = useState('ta'); // Default to Tamil Direct
  const [fontFamily, setFontFamily] = useState('TAU-Marutham'); 
  
  const [commands, setCommands] = useState({
    "முற்றுப்புள்ளி": ".", "கேள்விக்குறி": "?", "புதிய வரி": "\n"
  });
  const [newCmd, setNewCmd] = useState({ voice: '', result: '' });

  // --- LOCALIZATION DICTIONARY ---
  const locales = {
    en: {
      title: "Tamil Dictation",
      control: "Control",
      settings: "Settings",
      startBtn: "START DICTATION",
      stopBtn: "STOP SESSION",
      smartPunc: "Smart Punctuation",
      smartDesc: "Automatically convert spoken commands to symbols.",
      transEngine: "Dictation Mode",
      ta_ta: "Tamil (Direct Dictation)",
      en_en: "English (Direct Dictation)",
      ta_en: "Tamil ➔ English",
      en_ta: "English ➔ Tamil",
      fontStyle: "Font Style",
      voiceMaps: "Voice Shortcuts",
      spokenWord: "When I say...",
      outputText: "Type this...",
      addCmd: "Add Shortcut"
    },
    ta: {
      title: "தமிழ் டிக்டேஷன்",
      control: "கட்டுப்பாடு",
      settings: "அமைப்புகள்",
      startBtn: "டிக்டேஷனைத் தொடங்கு",
      stopBtn: "அமர்வை நிறுத்து",
      smartPunc: "ஸ்மார்ட் நிறுத்தற்குறிகள்",
      smartDesc: "பேசும் வார்த்தைகளை குறியீடுகளாக தானாகவே மாற்றும்.",
      transEngine: "டிக்டேஷன் முறை",
      ta_ta: "தமிழ் (நேரடி தட்டச்சு)",
      en_en: "ஆங்கிலம் (நேரடி தட்டச்சு)",
      ta_en: "தமிழ் ➔ ஆங்கிலம்",
      en_ta: "ஆங்கிலம் ➔ தமிழ்",
      fontStyle: "எழுத்துரு (Font)",
      voiceMaps: "குரல் சுருக்குவழிகள்",
      spokenWord: "நான் சொல்லும் வார்த்தை...",
      outputText: "வர வேண்டிய குறி...",
      addCmd: "சுருக்குவழியைச் சேர்"
    }
  };

  const t = locales[uiLang];

  useEffect(() => {
    const saved = localStorage.getItem('linguist_prefs');
    if (saved) {
      const parsed = JSON.parse(saved);
      setCommands(parsed.commands || {});
      setSmartPunc(parsed.smartPunc ?? true);
      if (parsed.uiLang) setUiLang(parsed.uiLang);
      if (parsed.transMode) setTransMode(parsed.transMode);
      if (parsed.fontFamily) setFontFamily(parsed.fontFamily);
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('linguist_prefs', JSON.stringify({ commands, smartPunc, uiLang, transMode, fontFamily }));
  }, [commands, smartPunc, uiLang, transMode, fontFamily]);

  // Zero-Setup Translation Engine
  const translate = async (text, mode) => {
    if (mode === 'ta' || mode === 'en' || !text) return text;
    
    try {
      const source = mode === 'ta-en' ? 'ta' : 'en';
      const target = mode === 'ta-en' ? 'en' : 'ta';
      const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${source}&tl=${target}&dt=t&q=${encodeURIComponent(text)}`;
      
      const response = await fetch(url);
      const data = await response.json();
      
      if (data && data[0]) {
        return data[0].map(item => item[0]).join('');
      }
      return text;
    } catch (e) { 
      return text; 
    }
  };

  const deleteCommand = (key) => {
    const updated = { ...commands };
    delete updated[key];
    setCommands(updated);
  };

  // Word Document Injection
  const handleInjection = async (raw, isFinal) => {
    if (isWriting.current && !isFinal) return;
    let processed = raw;

    if (isFinal) {
      if (smartPunc) {
        Object.keys(commands).forEach(k => {
          processed = processed.replace(new RegExp(k, 'g'), commands[k]);
        });
      }
      processed = await translate(processed, transMode);
    }

    try {
      await Word.run(async (context) => {
        isWriting.current = true;
        const tag = 'TamilAnchor';
        let ccs = context.document.contentControls.getByTag(tag);
        ccs.load('items'); 
        await context.sync();

        let cc = ccs.items.length === 0 ? context.document.getSelection().insertContentControl() : ccs.items[0];
        cc.tag = tag; cc.appearance = 'Hidden';

        if (isFinal && processed.includes('\n')) {
          cc.delete(true);
          context.document.getSelection().insertBreak(Word.BreakType.paragraph);
          isWriting.current = false;
        } else {
          cc.insertText(processed, 'Replace');
          const range = cc.getRange();
          range.paragraphs.getFirst().readingOrder = "LeftToRight";
          range.paragraphs.getFirst().alignment = "Left";
          
          range.font.name = fontFamily;
          
          if (isFinal) {
            const finalRange = cc.getRange();
            cc.delete(true);
            const space = finalRange.insertText(' ', 'End');
            space.font.name = fontFamily; 
            space.select('End');
            isWriting.current = false;
          }
        }
        await context.sync();
      });
    } catch (e) { isWriting.current = false; }
  };

  const stopSession = () => {
    if (dialog) { try { dialog.close(); } catch (e) {} }
    setDialog(null);
  };

  const startSession = () => {
    let listenLang = 'ta-IN';
    if (transMode === 'en' || transMode === 'en-ta') {
      listenLang = 'en-IN';
    }

    const url = window.location.origin + `/assets/dictation.html?lang=${listenLang}&font=${encodeURIComponent(fontFamily)}`;
    
    Office.context.ui.displayDialogAsync(url, { height: 45, width: 35 }, (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) return;
      setDialog(res.value); 
      
      res.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const d = JSON.parse(arg.message); 
        handleInjection(d.text, d.isFinal);
      });
      res.value.addEventHandler(Office.EventType.DialogEventReceived, stopSession);
    });
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        
        <div style={styles.navContainer}>
          <div style={styles.nav}>
            <button onClick={() => setActiveTab('main')} style={activeTab === 'main' ? styles.tabActive : styles.tab}>{t.control}</button>
            <button onClick={() => setActiveTab('set')} style={activeTab === 'set' ? styles.tabActive : styles.tab}>{t.settings}</button>
          </div>
          
          <div style={styles.langSwitcher}>
            <button onClick={() => setUiLang('en')} style={uiLang === 'en' ? styles.langBtnActive : styles.langBtn}>EN</button>
            <button onClick={() => setUiLang('ta')} style={uiLang === 'ta' ? styles.langBtnActive : styles.langBtn}>தமிழ்</button>
          </div>
        </div>

        {activeTab === 'main' ? (
          <div style={styles.scrollArea}>
            
            {/* BRANDING HEADER WITH LOGO */}
            <div style={styles.titleRow}>
              <img src="/assets/logo.jpg" alt="Tamil Dictation Logo" style={styles.logo} />
              <h2 style={styles.title}>{t.title}</h2>
            </div>
            
            <button 
              onClick={dialog ? stopSession : startSession} 
              style={{...styles.mainBtn, background: dialog ? '#d13438' : '#0078d4'}}
            >
              {dialog ? t.stopBtn : t.startBtn}
            </button>
            
            <div style={styles.toggleBox}>
              <div style={styles.row}>
                <span style={{ fontWeight: '600', color: '#323130' }}>{t.smartPunc}</span>
                <input type="checkbox" checked={smartPunc} onChange={() => setSmartPunc(!smartPunc)} style={{ cursor: 'pointer', flexShrink: 0 }}/>
              </div>
              <p style={styles.descText}>{t.smartDesc}</p>
            </div>
          </div>
        ) : (
          <div style={styles.scrollArea}>
            <h4 style={styles.label}>{t.transEngine}</h4>
            <select value={transMode} onChange={e => setTransMode(e.target.value)} style={styles.input}>
              <option value="ta">{t.ta_ta}</option>
              <option value="en">{t.en_en}</option>
              <option value="ta-en">{t.ta_en}</option>
              <option value="en-ta">{t.en_ta}</option>
            </select>

            <h4 style={styles.label}>{t.fontStyle}</h4>
            <select value={fontFamily} onChange={e => setFontFamily(e.target.value)} style={styles.input}>
              <option value="TAU-Marutham">TAU-Marutham</option>
              <option value="Latha">Latha</option>
            </select>
            
            <h4 style={styles.label}>{t.voiceMaps}</h4>
            <div style={styles.inputWrap}>
                <input placeholder={t.spokenWord} value={newCmd.voice} onChange={e => setNewCmd({...newCmd, voice: e.target.value})} style={styles.flexInput}/>
                <input placeholder={t.outputText} value={newCmd.result} onChange={e => setNewCmd({...newCmd, result: e.target.value})} style={styles.flexInput}/>
            </div>
            
            <button onClick={() => { if(newCmd.voice && newCmd.result) { setCommands({...commands, [newCmd.voice]: newCmd.result}); setNewCmd({voice:'',result:''}); }}} style={styles.addBtn}>{t.addCmd}</button>
            
            <div style={styles.list}>
              {Object.keys(commands).map(k => (
                <div key={k} style={styles.listItem}>
                  <div style={{display: 'flex', gap: '8px', flex: 1, alignItems: 'center', minWidth: 0}}>
                    <span style={{color: '#0078d4', fontWeight: '600', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'}}>{k}</span>
                    <b style={{color: '#edebe9', flexShrink: 0}}>➔</b>
                    <span style={{ color: '#323130', wordBreak: 'break-all' }}>{commands[k] === '\n' ? '↵' : commands[k]}</span>
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

const styles = {
  container: { padding: '12px', backgroundColor: '#f3f2f1', height: '100vh', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', fontFamily: '"Segoe UI", system-ui, sans-serif' },
  card: { flex: 1, backgroundColor: '#fff', padding: '16px', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', display: 'flex', flexDirection: 'column', border: '1px solid #edebe9', boxSizing: 'border-box', overflow: 'hidden' },
  navContainer: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px', borderBottom: '1px solid #edebe9', paddingBottom: '12px', marginBottom: '16px' },
  nav: { display: 'flex', gap: '4px', flexWrap: 'wrap' },
  tab: { padding: '8px 12px', background: 'none', border: 'none', color: '#605e5c', cursor: 'pointer', fontWeight: '600', fontSize: '13px', borderRadius: '4px' },
  tabActive: { padding: '8px 12px', background: '#f3f2f1', border: 'none', color: '#0078d4', fontWeight: 'bold', fontSize: '13px', borderRadius: '4px' },
  langSwitcher: { display: 'flex', backgroundColor: '#f3f2f1', borderRadius: '4px', padding: '2px', alignItems: 'center', flexShrink: 0 },
  langBtn: { border: 'none', background: 'transparent', padding: '4px 8px', fontSize: '12px', cursor: 'pointer', color: '#605e5c', fontWeight: '600', borderRadius: '2px', transition: 'all 0.2s' },
  langBtnActive: { border: 'none', background: '#ffffff', padding: '4px 8px', fontSize: '12px', cursor: 'pointer', color: '#0078d4', fontWeight: 'bold', borderRadius: '2px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', transition: 'all 0.2s' },
  scrollArea: { flex: 1, overflowY: 'auto', paddingRight: '4px', animation: 'fadeIn 0.3s ease-in', display: 'flex', flexDirection: 'column' },
  
  // LOGO & TITLE ALIGNMENT
  titleRow: { display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '20px' },
  logo: { width: '32px', height: '32px', borderRadius: '6px', objectFit: 'cover', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' },
  title: { fontSize: '22px', color: '#2b579a', margin: '0', fontWeight: '700', flexShrink: 0 },
  
  mainBtn: { width: '100%', padding: '14px', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', letterSpacing: '0.5px', transition: 'background 0.2s', boxShadow: '0 2px 4px rgba(0,0,0,0.1)', flexShrink: 0 },
  toggleBox: { marginTop: '20px', borderTop: '1px solid #edebe9', paddingTop: '16px', flexShrink: 0 },
  row: { display: 'flex', justifyContent: 'space-between', marginBottom: '4px', fontSize: '13px', alignItems: 'center', gap: '10px' },
  descText: { fontSize: '11px', color: '#888', marginTop: '0', lineHeight: '1.4' },
  label: { fontSize: '11px', color: '#a19f9d', textTransform: 'uppercase', marginBottom: '8px', marginTop: '8px', letterSpacing: '0.5px', fontWeight: 'bold', flexShrink: 0 },
  input: { width: '100%', padding: '10px', marginBottom: '12px', boxSizing: 'border-box', border: '1px solid #d2d0ce', borderRadius: '4px', fontSize: '13px', color: '#323130', flexShrink: 0 },
  inputWrap: { display: 'flex', flexWrap: 'wrap', gap: '8px', marginBottom: '12px', flexShrink: 0 },
  flexInput: { flex: '1 1 100px', padding: '10px', boxSizing: 'border-box', border: '1px solid #d2d0ce', borderRadius: '4px', fontSize: '13px', color: '#323130', minWidth: '100px' },
  addBtn: { width: '100%', padding: '12px', backgroundColor: '#f3f2f1', color: '#0078d4', border: '1px solid #0078d4', cursor: 'pointer', borderRadius: '4px', fontWeight: '600', fontSize: '13px', flexShrink: 0 },
  list: { marginTop: '16px', borderTop: '1px solid #edebe9', paddingTop: '8px', flexShrink: 0 },
  listItem: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', padding: '10px 0', borderBottom: '1px solid #f3f2f1', gap: '8px' },
  delBtn: { border: 'none', background: '#fde7e9', color: '#a4262c', cursor: 'pointer', padding: '6px 10px', borderRadius: '4px', fontWeight: 'bold', fontSize: '12px', flexShrink: 0 }
};

const css = document.createElement('style');
css.innerHTML = `
  @keyframes fadeIn { from { opacity: 0; transform: translateY(2px); } to { opacity: 1; transform: translateY(0); } }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: #c8c6c4; border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: #a19f9d; }
`;
document.head.appendChild(css);

export default App;