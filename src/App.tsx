import React, { useState, useEffect } from 'react';
import { 
  FileCode, 
  Folder, 
  Save, 
  Play, 
  RefreshCw, 
  Search, 
  ChevronRight, 
  FileText, 
  AlertCircle,
  Loader2,
  Code2,
  Terminal,
  Menu,
  ChevronLeft,
  ChevronDown,
  Info,
  GitBranch,
  CheckCircle2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  listFiles, 
  getFileContent, 
  saveFile, 
  getCurrentFolderPath, 
  hasSpContext 
} from './services/sharepointService';

export default function App() {
  const [files, setFiles] = useState<any[]>([]);
  const [selectedFile, setSelectedFile] = useState<any>(null);
  const [code, setCode] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [currentFolder, setCurrentFolder] = useState('');

  useEffect(() => {
    if (!hasSpContext()) {
      setError('SharePoint Context não detectado. Usando ambiente de demonstração.');
      setFiles([
        { Name: 'index.aspx', ServerRelativeUrl: '/site/index.aspx', Length: '4096' },
        { Name: 'sp-connector.ts', ServerRelativeUrl: '/site/sp-connector.ts', Length: '2048' },
        { Name: 'DHL-styles.css', ServerRelativeUrl: '/site/DHL-styles.css', Length: '1024' },
      ]);
      setCurrentFolder('/sites/DHL-Supply-Chain');
    } else {
      loadFiles();
    }
  }, []);

  const loadFiles = async () => {
    setLoading(true);
    try {
      const folder = getCurrentFolderPath();
      setCurrentFolder(folder);
      const result = await listFiles(folder);
      if (result.status) {
        setFiles(result.data);
      } else {
        setError(result.message);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleSelectFile = async (file: any) => {
    setSelectedFile(file);
    setLoading(true);
    setError(null);
    try {
      const result = await getFileContent(file.ServerRelativeUrl);
      if (result.status) {
        setCode(result.data);
      } else {
        setError(result.message);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    if (!selectedFile) return;
    setSaving(true);
    try {
      const result = await saveFile(selectedFile.ServerRelativeUrl, code);
      if (!result.status) {
        setError(result.message);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setSaving(false);
    }
  };

  const handleRun = () => {
    const indexPath = `${currentFolder}/index.aspx`;
    window.open(indexPath, '_blank');
  };

  return (
    <div className="flex h-screen w-full overflow-hidden bg-dhl-dark text-slate-300 font-sans select-none">
      {/* Header - Professional Polish Theme */}
      <header className="absolute top-0 left-0 right-0 h-12 bg-dhl-yellow flex items-center px-4 justify-between z-50 shadow-lg border-b border-black/10">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <div className="w-9 h-9 bg-dhl-red flex items-center justify-center font-black italic text-xl text-white transform -skew-x-12 shadow-md">
              DHL
            </div>
            <span className="font-bold text-black uppercase tracking-tighter text-sm ml-1">
              SharePoint Dev Studio
            </span>
          </div>
          
          <div className="h-6 w-px bg-black/10 mx-1" />
          
          <div className="hidden md:flex items-center gap-2 px-3 py-1 bg-black/5 rounded text-[11px] font-medium text-black/70">
            <Folder className="w-3 h-3" />
            <span className="opacity-60">Site Path:</span>
            <span className="font-bold text-black">{currentFolder || 'Root'}</span>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {selectedFile && (
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleSave}
              disabled={saving}
              className="flex items-center gap-2 bg-white/90 text-dhl-red px-4 py-1.5 rounded font-bold text-[11px] shadow-sm hover:bg-white transition-all disabled:opacity-50 border border-black/5"
            >
              {saving ? <Loader2 className="w-3 h-3 animate-spin" /> : <Save className="w-3 h-3" />}
              SALVAR
            </motion.button>
          )}
          <motion.button
            whileHover={{ scale: 1.02, brightness: 1.1 }}
            whileTap={{ scale: 0.98 }}
            onClick={handleRun}
            className="flex items-center gap-2 bg-dhl-red text-white px-5 py-1.5 rounded font-bold text-[11px] shadow-md transition-all uppercase tracking-wider"
          >
            <Play className="w-3 h-3 fill-current" />
            Executar Index.aspx
          </motion.button>
        </div>
      </header>

      <div className="flex h-full w-full pt-12">
        {/* Sidebar - Dark theme */}
        <motion.aside 
          initial={false}
          animate={{ width: isSidebarOpen ? 260 : 0 }}
          className="bg-dhl-sidebar flex flex-col overflow-hidden border-r border-white/5 shadow-2xl z-20"
        >
          <div className="px-4 py-3 flex items-center justify-between shrink-0 bg-black/10">
            <div className="text-[10px] font-extrabold text-white/40 uppercase tracking-[0.2em] flex items-center gap-2">
              <span className="w-1.5 h-1.5 bg-dhl-yellow rounded-full" />
              File Explorer
            </div>
            <button 
              onClick={() => setIsSidebarOpen(false)}
              className="p-1 hover:bg-white/5 rounded text-white/40 hover:text-white transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
          </div>

          <div className="flex-1 overflow-y-auto pt-2">
            {loading && !selectedFile ? (
              <div className="flex flex-col items-center justify-center py-12 gap-3 opacity-30">
                <Loader2 className="w-6 h-6 animate-spin" />
                <span className="text-[10px] font-bold uppercase tracking-widest">Loading...</span>
              </div>
            ) : (
              <div className="px-2 space-y-0.5">
                {files.map((file, idx) => (
                  <button
                    key={idx}
                    onClick={() => handleSelectFile(file)}
                    className={`w-full text-left px-3 py-1.5 rounded flex items-center gap-3 group transition-all relative ${
                      selectedFile?.ServerRelativeUrl === file.ServerRelativeUrl 
                        ? 'bg-white/10 text-white border-l-2 border-dhl-yellow' 
                        : 'hover:bg-white/5 text-slate-400 font-medium'
                    }`}
                  >
                    <FileCode className={`w-4 h-4 shrink-0 ${
                      selectedFile?.ServerRelativeUrl === file.ServerRelativeUrl 
                        ? 'text-blue-400' 
                        : 'text-slate-500 group-hover:text-slate-300'
                    }`} />
                    <span className="text-sm truncate select-none">{file.Name}</span>
                  </button>
                ))}
              </div>
            )}

            <div className="mt-8 px-4 py-2">
              <div className="text-[10px] font-extrabold text-white/40 uppercase tracking-[0.2em] mb-3">SP Resources</div>
              <div className="space-y-2">
                <div className="bg-black/20 rounded-lg p-3 border border-white/5">
                  <div className="text-[9px] text-white/50 mb-1 font-mono uppercase">Request Digest</div>
                  <div className="text-[10px] text-blue-300 break-all font-mono leading-tight bg-black/40 p-1.5 rounded border border-white/5">
                    {hasSpContext() ? '0x' + Math.random().toString(16).slice(2, 20).toUpperCase() : 'ENV_DEV_MOCK'}
                  </div>
                </div>
                <button className="w-full text-left px-3 py-2 rounded border border-white/5 hover:bg-white/5 text-[11px] font-bold text-white/60 transition-colors flex items-center justify-between">
                  List Inspector
                  <ChevronRight className="w-3 h-3" />
                </button>
              </div>
            </div>
          </div>
        </motion.aside>

        {/* Floating Sidebar Toggle */}
        {!isSidebarOpen && (
          <motion.button
            initial={{ x: -20, opacity: 0 }}
            animate={{ x: 0, opacity: 1 }}
            onClick={() => setIsSidebarOpen(true)}
            className="fixed left-0 top-1/2 -translate-y-1/2 w-6 h-12 bg-dhl-sidebar flex items-center justify-center rounded-r border border-white/10 z-40 text-white/50 hover:text-white"
          >
            <ChevronRight className="w-4 h-4" />
          </motion.button>
        )}

        {/* Editor Area */}
        <section className="flex-1 flex flex-col bg-dhl-dark min-w-0">
          {/* Breadcrumbs / Editor Bar */}
          <div className="h-9 bg-black/30 border-b border-white/5 px-4 flex items-center justify-between shrink-0">
            <div className="flex items-center gap-2 text-[11px] text-white/40 font-medium">
              <span className="hover:text-white/80 cursor-pointer">src</span>
              <ChevronRight className="w-3 h-3 opacity-30" />
              <span className="hover:text-white/80 cursor-pointer">core</span>
              <ChevronRight className="w-3 h-3 opacity-30" />
              <span className="text-white/90">{selectedFile?.Name || 'welcome'}</span>
            </div>
            
            {hasSpContext() && (
              <div className="flex items-center gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.5)]" />
                <span className="text-[10px] uppercase font-black text-white/40 tracking-wider">Connected</span>
              </div>
            )}
          </div>

          <div className="flex-1 overflow-hidden relative group">
            {error && (
              <div className="absolute top-0 left-0 right-0 z-30 p-3 bg-red-500/10 border-b border-red-500/20 backdrop-blur-md flex items-center gap-3 text-red-400 text-[11px] font-bold tracking-tight">
                <AlertCircle className="w-4 h-4 shrink-0" />
                {error}
                <button onClick={() => setError(null)} className="ml-auto underline hover:text-white">CLOSE</button>
              </div>
            )}

            {selectedFile ? (
              <div className="flex h-full">
                {/* Line numbers simulation */}
                <div className="w-12 bg-black/10 border-r border-white/5 py-6 flex flex-col items-center text-[11px] text-white/20 font-mono select-none overflow-hidden leading-[1.625rem]">
                  {Array.from({ length: Math.min(code.split('\n').length + 5, 100) }).map((_, i) => (
                    <div key={i}>{i + 1}</div>
                  ))}
                </div>
                <textarea
                  value={code}
                  onChange={(e) => setCode(e.target.value)}
                  className="flex-1 h-full p-6 outline-none resize-none bg-transparent text-slate-200 leading-relaxed font-mono text-[13px] selection:bg-dhl-yellow/30"
                  spellCheck={false}
                  placeholder="// No data found"
                />
              </div>
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center p-12 text-center h-full">
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="max-w-md"
                >
                  <div className="w-24 h-24 bg-dhl-yellow/5 border-2 border-white/5 rounded-3xl flex items-center justify-center mb-8 mx-auto rotate-3 group-hover:rotate-0 transition-transform">
                    <Code2 className="w-10 h-10 text-dhl-yellow opacity-40" />
                  </div>
                  <h2 className="text-3xl font-black text-white italic tracking-tighter mb-3 uppercase">DH<span className="text-dhl-yellow">L</span> Supply Logic</h2>
                  <p className="text-slate-500 text-sm mb-10 leading-loose">
                    Ambiente de desenvolvimento <span className="text-white/80 border-b border-dhl-yellow/50">proativo</span> para customização de processos SharePoint DHL.
                  </p>
                  
                  <div className="grid grid-cols-2 gap-4">
                    <button 
                      onClick={() => setIsSidebarOpen(true)}
                      className="p-4 bg-white/5 border border-white/5 rounded-xl hover:bg-white/10 hover:border-dhl-yellow/30 transition-all text-left group"
                    >
                      <div className="text-[10px] font-black text-white/40 uppercase mb-2">Workspace</div>
                      <div className="text-xs font-bold flex items-center gap-2">Abrir Explorer <ChevronRight className="w-3 h-3" /></div>
                    </button>
                    <button 
                      onClick={handleRun}
                      className="p-4 bg-white/5 border border-white/5 rounded-xl hover:bg-white/10 hover:border-dhl-yellow/30 transition-all text-left"
                    >
                      <div className="text-[10px] font-black text-white/40 uppercase mb-2">Production</div>
                      <div className="text-xs font-bold flex items-center gap-2">Visualizar Site <ChevronRight className="w-3 h-3" /></div>
                    </button>
                  </div>
                </motion.div>
              </div>
            )}
          </div>
        </section>
      </div>

      {/* Footer - Status Bar */}
      <footer className="h-6 shrink-0 bg-dhl-sidebar border-t border-white/5 z-50 flex items-center justify-between px-3 text-[10px] font-medium text-white/40">
        <div className="flex items-center gap-4 h-full">
          <div className="flex items-center gap-1.5 h-full px-2 hover:bg-white/5 cursor-pointer">
            <CheckCircle2 className="w-3 h-3 text-green-500" />
            <span>Ready</span>
          </div>
          <div className="flex items-center gap-2 h-full px-2 hover:bg-white/5 cursor-pointer">
            <GitBranch className="w-3 h-3" />
            <span>master*</span>
          </div>
          <div className="flex items-center gap-2 h-full px-2 hover:bg-white/5 cursor-pointer">
            <Info className="w-3 h-3" />
            <span>0 Errors</span>
          </div>
        </div>

        <div className="flex items-center gap-4 h-full">
          <div className="font-mono">{selectedFile ? `Ln ${code.split('\n').length}, Col 1` : 'Idle'}</div>
          <div>UTF-8</div>
          <div className="flex items-center gap-1.5 text-dhl-yellow">
            <div className="w-1.5 h-1.5 rounded-full bg-dhl-yellow animate-pulse" />
            <span>TypeScript 5.8.2</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
