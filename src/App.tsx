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
  Plus,
  ArrowLeft,
  FolderOpen,
  FilePlus,
  X,
  CheckCircle2,
  Zap
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  listItems, 
  getFileContent, 
  saveFile, 
  createFile,
  getCurrentFolderPath, 
  parseSpUrl,
  spSiteUrl,
  hasSpContext 
} from './services/sharepointService';
import { transform } from 'sucrase';

export default function App() {
  const [files, setFiles] = useState<any[]>([]);
  const [folders, setFolders] = useState<any[]>([]);
  const [selectedFile, setSelectedFile] = useState<any>(null);
  const [code, setCode] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  
  // Workspace Dynamic path state
  const [currentFolder, setCurrentFolder] = useState('');
  const [baseFolder, setBaseFolder] = useState('');
  const [activeSiteUrl, setActiveSiteUrl] = useState('');
  const [ideServerUrl, setIdeServerUrl] = useState('');
  
  // New file state
  const [isNewFileModalOpen, setIsNewFileModalOpen] = useState(false);
  const [newFileName, setNewFileName] = useState('');

  // Workspace Settings
  const [isWorkspaceModalOpen, setIsWorkspaceModalOpen] = useState(false);
  const [tempWorkspaceUrl, setTempWorkspaceUrl] = useState('');
  const [tempIdeUrl, setTempIdeUrl] = useState('');

  useEffect(() => {
    // Carrega workspace salvo do localstorage se existir
    const savedSite = localStorage.getItem('dhl_sp_site');
    const savedFolder = localStorage.getItem('dhl_sp_folder');
    const savedIde = localStorage.getItem('dhl_ide_url');

    if (savedIde) setIdeServerUrl(savedIde);
    else if (!window.location.hostname.includes('sharepoint.com')) {
      // Se não for SharePoint, assume que o servidor é a origem atual
      setIdeServerUrl(window.location.origin);
    }

    if (savedSite && savedFolder) {
      setActiveSiteUrl(savedSite);
      setBaseFolder(savedFolder);
      loadPath(savedFolder, savedSite);
    } else if (!hasSpContext()) {
      setError('Ambiente de demonstração. Configure um workspace real no SharePoint.');
      setFiles([
        { Name: 'index.aspx', ServerRelativeUrl: '/site/index.aspx', Length: '4096' },
        { Name: 'sp-connector.ts', ServerRelativeUrl: '/site/sp-connector.ts', Length: '2048' },
      ]);
      setFolders([
        { Name: 'components', ServerRelativeUrl: '/site/components' },
        { Name: 'assets', ServerRelativeUrl: '/site/assets' }
      ]);
      const mockPath = '/sites/DHL-Supply-Chain';
      setCurrentFolder(mockPath);
      setBaseFolder(mockPath);
      setActiveSiteUrl('https://tenant.sharepoint.com/sites/DHL');
    } else {
      const folder = getCurrentFolderPath();
      const site = spSiteUrl();
      setBaseFolder(folder);
      setActiveSiteUrl(site);
      loadPath(folder, site);
    }
  }, []);

  const loadPath = async (path: string, site?: string) => {
    const targetSite = site || activeSiteUrl;
    
    // Se for ambiente de demo e usarmos a URL de mock, não tentamos buscar via API real
    if (!hasSpContext() && (!targetSite || targetSite.includes('tenant.sharepoint.com'))) {
      setCurrentFolder(path);
      return; 
    }

    setLoading(true);
    setCurrentFolder(path);
    try {
      const result = await listItems(path, targetSite);
      if (result.status) {
        setFiles(result.data.files);
        setFolders(result.data.folders);
        setError(null);
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
    if (!hasSpContext() && activeSiteUrl.includes('tenant.sharepoint.com')) {
      setCode('// Conteúdo de demonstração\nconsole.log("DHL Supply Chain - " + "' + file.Name + '");');
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const result = await getFileContent(file.ServerRelativeUrl, activeSiteUrl);
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
    if (!hasSpContext() && activeSiteUrl.includes('tenant.sharepoint.com')) {
      setError('Atenção: Modo de Demonstração. As alterações não serão persistidas no SharePoint.');
      return;
    }
    setSaving(true);
    try {
      const result = await saveFile(selectedFile.ServerRelativeUrl, code, activeSiteUrl);
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
    const indexPath = `${activeSiteUrl}${currentFolder}/index.aspx`;
    window.open(indexPath, '_blank');
  };

  const handleSaveFolder = (folder: any) => {
    loadPath(folder.ServerRelativeUrl, activeSiteUrl);
  };

  const navigateUp = () => {
    if (currentFolder === baseFolder) return;
    const parentPath = currentFolder.substring(0, currentFolder.lastIndexOf('/'));
    loadPath(parentPath, activeSiteUrl);
  };

  const handleCreateFile = async () => {
    if (!newFileName) return;
    if (!hasSpContext() && activeSiteUrl.includes('tenant.sharepoint.com')) {
      setNewFileName('');
      setIsNewFileModalOpen(false);
      setFiles(prev => [...prev, { Name: newFileName, ServerRelativeUrl: currentFolder + '/' + newFileName, Length: '0' }]);
      return;
    }
    setSaving(true);
    try {
      const result = await createFile(currentFolder, newFileName, '<!-- DHL New File -->', activeSiteUrl);
      if (result.status) {
        setNewFileName('');
        setIsNewFileModalOpen(false);
        loadPath(currentFolder, activeSiteUrl);
      } else {
        setError(result.message);
      }
    } catch (err: any) {
      setError(err.message);
    } finally {
      setSaving(false);
    }
  };

  const applyWorkspace = () => {
    const { siteUrl, folderPath } = parseSpUrl(tempWorkspaceUrl);
    if (!siteUrl || !folderPath) {
      setError('URL do SharePoint Inválida. Use o formato: https://dpdhl.sharepoint.com/sites/NomeDoSite/Pasta');
      return;
    }
    
    // Valida e limpa URL do IDE
    let cleanIdeUrl = tempIdeUrl.trim().replace(/\/$/, "");
    if (cleanIdeUrl && !cleanIdeUrl.startsWith('http')) {
      setError('URL do Servidor IDE Inválida. Deve começar com http:// ou https://');
      return;
    }

    if (cleanIdeUrl) {
      setIdeServerUrl(cleanIdeUrl);
      localStorage.setItem('dhl_ide_url', cleanIdeUrl);
    } else {
      // Se limpar o campo, remove do localStorage
      setIdeServerUrl('');
      localStorage.removeItem('dhl_ide_url');
    }

    setActiveSiteUrl(siteUrl);
    setBaseFolder(folderPath);
    setCurrentFolder(folderPath);
    
    localStorage.setItem('dhl_sp_site', siteUrl);
    localStorage.setItem('dhl_sp_folder', folderPath);
    
    loadPath(folderPath, siteUrl);
    setIsWorkspaceModalOpen(false);
    setError(null);
  };

  // Build via Browser (Sem Backend)
  const handleBrowserBuild = async () => {
    if (!selectedFile) return;
    setExecutingCommand(true);
    setIsTerminalOpen(true);
    setTerminalLogs(prev => [...prev, `[BROWSER-BUILD] Iniciando compilação de ${selectedFile.Name}...`]);
    
    try {
      // Transpilação TSX/TS para JS usando Sucrase (100% Client-side)
      const compiled = transform(code, {
        transforms: ['typescript', 'jsx', 'imports'],
        production: true,
      });

      const jsFileName = selectedFile.Name.replace(/\.(tsx|ts|jsx)$/, '.js');
      const jsFileUrl = `${currentFolder}/${jsFileName}`;
      
      setTerminalLogs(prev => [...prev, `[BROWSER-BUILD] Código transformado com sucesso. Salavando ${jsFileName}...`]);
      
      const saveResult = await saveFile(jsFileUrl, compiled.code, activeSiteUrl);
      
      if (saveResult.status) {
        setTerminalLogs(prev => [...prev, `[BROWSER-BUILD] SUCESSO: Arquivo ${jsFileName} gerado e salvo no SharePoint.`]);
        loadPath(currentFolder, activeSiteUrl); // Atualiza lista
      } else {
        throw new Error(saveResult.message);
      }
    } catch (err: any) {
      setTerminalLogs(prev => [...prev, `[BROWSER-BUILD] ERRO: ${err.message}`]);
    } finally {
      setExecutingCommand(false);
    }
  };

  const [isTerminalOpen, setIsTerminalOpen] = useState(false);
  const [terminalLogs, setTerminalLogs] = useState<string[]>([]);
  const [executingCommand, setExecutingCommand] = useState(false);

  const runCommand = async (command: string) => {
    if (executingCommand) return;
    
    if (!ideServerUrl) {
      setIsWorkspaceModalOpen(true);
      setError('Por favor, configure o "URL do Servidor IDE" no modal de configurações para rodar comandos como NPM.');
      return;
    }
    
    setExecutingCommand(true);
    setIsTerminalOpen(true);
    setTerminalLogs([]); 
    
    try {
      const response = await fetch(`${ideServerUrl}/api/terminal/run`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ command }),
        mode: 'cors'
      }).catch(err => {
        if (err.message.includes('Failed to fetch')) {
          throw new Error(`Erro de Conexão: Não foi possível alcançar o servidor em "${ideServerUrl}". Verifique se o URL está correto e se o Servidor IDE no AI Studio está ativo e com CORS liberado.`);
        }
        throw err;
      });

      const responseText = await response.text();
      let data;
      try {
        data = JSON.parse(responseText);
      } catch (e) {
        throw new Error(`O servidor IDE em "${ideServerUrl}" não retornou um JSON válido.\n\nSe você está no SharePoint, verifique se o URL do Servidor IDE nas configurações está correto e se o servidor está rodando.`);
      }
      
      const taskId = data.taskId;
      if (!taskId) throw new Error(data.error || 'Não foi possível iniciar a tarefa');

      // Polling function
      const poll = async () => {
        try {
          const res = await fetch(`${ideServerUrl}/api/terminal/logs/${taskId}`, { mode: 'cors' });
          const logText = await res.text();
          let logData;
          try {
            logData = JSON.parse(logText);
          } catch (e) {
            setTerminalLogs(prev => [...prev, `ERR_POLL: Resposta inválida`]);
            setExecutingCommand(false);
            return;
          }
          
          setTerminalLogs(logData.logs);
          
          if (logData.status === 'running') {
            setTimeout(poll, 2000);
          } else {
            setExecutingCommand(false);
          }
        } catch (err) {
          console.error('Polling error:', err);
          setExecutingCommand(false);
        }
      };

      poll();
    } catch (err: any) {
      setTerminalLogs(prev => [...prev, `FAILED: ${err.message}`]);
      setExecutingCommand(false);
    }
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
          
          <button 
            onClick={() => {
              setTempWorkspaceUrl(`${activeSiteUrl}${currentFolder}`);
              setTempIdeUrl(ideServerUrl);
              setIsWorkspaceModalOpen(true);
            }}
            className="hidden md:flex items-center gap-2 px-3 py-1 bg-black/5 rounded text-[11px] font-medium text-black/70 hover:bg-black/10 transition-all cursor-pointer group"
          >
            <Folder className="w-3 h-3 translate-y-[1px] group-hover:text-dhl-red transition-colors" />
            <span className="opacity-60">Workspace:</span>
            <span className="font-bold text-black truncate max-w-[250px]">{currentFolder || 'Configurar...'}</span>
          </button>
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
            <div className="flex items-center gap-1">
              <button 
                onClick={() => setIsNewFileModalOpen(true)}
                title="Novo Arquivo"
                className="p-1.5 hover:bg-white/5 rounded text-white/40 hover:text-dhl-yellow transition-colors"
                id="btn-new-file"
              >
                <Plus className="w-4 h-4" />
              </button>
              <button 
                onClick={() => setIsSidebarOpen(false)}
                className="p-1.5 hover:bg-white/5 rounded text-white/40 hover:text-white transition-colors"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Navigation Bar */}
          <div className="px-2 py-1 flex items-center gap-1 bg-black/5 border-b border-white/5 shrink-0">
            <button 
              onClick={navigateUp}
              disabled={currentFolder === baseFolder}
              className={`p-1.5 rounded transition-colors ${currentFolder === baseFolder ? 'opacity-20 pointer-events-none' : 'hover:bg-white/5 text-white/60'}`}
            >
              <ArrowLeft className="w-3.5 h-3.5" />
            </button>
            <div className="text-[9px] font-mono text-white/30 truncate">
              {currentFolder.split('/').pop() || 'Root'}
            </div>
            <button 
              onClick={() => loadPath(currentFolder, activeSiteUrl)}
              className="ml-auto p-1.5 hover:bg-white/5 rounded text-white/40 hover:text-white transition-colors"
            >
              <RefreshCw className={`w-3 h-3 ${loading ? 'animate-spin' : ''}`} />
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
                {/* Folders */}
                {folders.map((folder, idx) => (
                  <button
                    key={`folder-${idx}`}
                    onClick={() => handleSaveFolder(folder)}
                    className="w-full text-left px-3 py-1.5 rounded flex items-center gap-3 group transition-all text-slate-400 font-medium hover:bg-white/5"
                  >
                    <Folder className="w-4 h-4 shrink-0 text-dhl-yellow/60 group-hover:text-dhl-yellow" />
                    <span className="text-sm truncate select-none">{folder.Name}</span>
                  </button>
                ))}

                {/* Files */}
                {files.map((file, idx) => (
                  <button
                    key={`file-${idx}`}
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

                {files.length === 0 && folders.length === 0 && !loading && (
                  <div className="text-center py-8 opacity-20 text-[10px] uppercase font-black">
                    Pasta Vazia
                  </div>
                )}
              </div>
            )}

            <div className="mt-8 px-4 py-2">
              <div className="text-[10px] font-extrabold text-white/40 uppercase tracking-[0.2em] mb-3">SP Resources</div>
              <div className="space-y-2">
                <div className="bg-black/20 rounded-lg p-3 border border-white/5">
                  <div className="text-[9px] text-white/50 mb-1 font-mono uppercase">Site URL</div>
                  <div className="text-[9px] text-blue-300 break-all font-mono leading-tight bg-black/40 p-1.5 rounded border border-white/5">
                    {activeSiteUrl}
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
            <div className="flex items-center gap-2 text-[11px] text-white/40 font-medium overflow-hidden whitespace-nowrap">
              <span className="hover:text-white/80 cursor-pointer shrink-0" onClick={() => loadPath(baseFolder, activeSiteUrl)}>root</span>
              {currentFolder.replace(baseFolder, '').split('/').filter(Boolean).map((part, i) => (
                <React.Fragment key={i}>
                  <ChevronRight className="w-3 h-3 opacity-30 shrink-0" />
                  <span className="hover:text-white/80 cursor-pointer shrink-0">{part}</span>
                </React.Fragment>
              ))}
              {selectedFile && (
                <>
                  <ChevronRight className="w-3 h-3 opacity-30 shrink-0" />
                  <span className="text-white/90 shrink-0 truncate">{selectedFile.Name}</span>
                </>
              )}
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
                  {Array.from({ length: 100 }).map((_, i) => (
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
                      onClick={() => setIsWorkspaceModalOpen(true)}
                      className="p-4 bg-white/5 border border-white/5 rounded-xl hover:bg-white/10 hover:border-dhl-yellow/30 transition-all text-left group"
                    >
                      <div className="text-[10px] font-black text-white/40 uppercase mb-2">Workspace</div>
                      <div className="text-xs font-bold flex items-center gap-2">Trocar Pasta <ChevronRight className="w-3 h-3" /></div>
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

            {/* Terminal Panel */}
            <AnimatePresence>
              {isTerminalOpen && (
                <motion.div 
                  initial={{ height: 0 }}
                  animate={{ height: 200 }}
                  exit={{ height: 0 }}
                  className="absolute bottom-0 left-0 right-0 bg-dhl-sidebar border-t border-white/10 z-40 flex flex-col shadow-[0_-10px_30px_rgba(0,0,0,0.5)]"
                >
                  <div className="h-8 bg-black/40 px-4 flex items-center justify-between shrink-0">
                    <div className="flex items-center gap-4">
                      <div className="text-[10px] font-black text-white/40 uppercase tracking-[0.2em] flex items-center gap-2">
                        <Terminal className="w-3 h-3 text-dhl-yellow" />
                        Terminal Output
                      </div>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={handleBrowserBuild}
                          disabled={executingCommand}
                          className="text-[9px] font-bold text-dhl-yellow hover:text-white transition-colors uppercase flex items-center gap-1"
                        >
                          <Zap className="w-3 h-3" />
                          BUILD DIRETO (Sem Servidor)
                        </button>
                        <div className="w-px h-3 bg-white/10 mx-1" />
                        <button 
                          onClick={() => runCommand('npm install')}
                          disabled={executingCommand}
                          className="text-[9px] font-bold text-blue-400 hover:text-white transition-colors uppercase"
                        >
                          NPM INSTALL
                        </button>
                        <button 
                          onClick={() => runCommand('npm run build')}
                          disabled={executingCommand}
                          className="text-[9px] font-bold text-green-400 hover:text-white transition-colors uppercase"
                        >
                          NPM BUILD
                        </button>
                      </div>
                    </div>
                    <button onClick={() => setIsTerminalOpen(false)} className="text-white/20 hover:text-white">
                      <X className="w-3 h-3" />
                    </button>
                  </div>
                  <div className="flex-1 p-4 overflow-y-auto font-mono text-[11px] leading-relaxed select-text space-y-1 scrollbar-thin scrollbar-thumb-white/10">
                    {terminalLogs.length === 0 ? (
                      <span className="opacity-20 italic">Aguardando comandos...</span>
                    ) : (
                      terminalLogs.map((log, i) => (
                        <div key={i} className={log.startsWith('>') ? 'text-dhl-yellow font-bold mt-2' : log.includes('ERROR') ? 'text-red-400' : 'text-slate-300'}>
                          {log}
                        </div>
                      ))
                    )}
                    {executingCommand && (
                      <div className="flex items-center gap-2 text-white/50 animate-pulse mt-2">
                        <Loader2 className="w-3 h-3 animate-spin" />
                        <span>Processando...</span>
                      </div>
                    )}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* New File Modal */}
          <AnimatePresence>
            {isNewFileModalOpen && (
              <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
                <motion.div 
                  initial={{ scale: 0.9, opacity: 0 }}
                  animate={{ scale: 1, opacity: 1 }}
                  exit={{ scale: 0.9, opacity: 0 }}
                  className="bg-dhl-sidebar border border-white/10 w-full max-w-sm rounded-2xl shadow-2xl overflow-hidden"
                >
                  <div className="px-6 py-4 border-b border-white/5 flex items-center justify-between">
                    <h3 className="text-sm font-bold text-white flex items-center gap-2 italic uppercase italic">
                      <FilePlus className="w-4 h-4 text-dhl-yellow" />
                      Novo Arquivo
                    </h3>
                    <button onClick={() => setIsNewFileModalOpen(false)} className="text-white/30 hover:text-white transition-colors">
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="p-6 space-y-4">
                    <div>
                      <label className="block text-[10px] font-black text-white/40 uppercase mb-2">Nome do Arquivo</label>
                      <input 
                        type="text" 
                        value={newFileName}
                        onChange={(e) => setNewFileName(e.target.value)}
                        autoFocus
                        placeholder="exemplo.aspx"
                        className="w-full bg-black/40 border border-white/10 rounded-lg px-4 py-3 text-white text-sm outline-none focus:border-dhl-yellow/50 transition-all"
                        onKeyDown={(e) => e.key === 'Enter' && handleCreateFile()}
                      />
                    </div>
                    <div className="text-[10px] text-white/30 italic">
                      O arquivo será criado em: <span className="text-white/50">{currentFolder}</span>
                    </div>
                  </div>
                  <div className="px-6 py-4 bg-black/10 border-t border-white/5 flex gap-3">
                    <button 
                      onClick={() => setIsNewFileModalOpen(false)}
                      className="flex-1 py-2 rounded-lg text-xs font-bold text-white/60 hover:bg-white/5 transition-all uppercase"
                    >
                      Cancelar
                    </button>
                    <button 
                      onClick={handleCreateFile}
                      disabled={!newFileName || saving}
                      className="flex-1 py-2 bg-dhl-red rounded-lg text-xs font-bold text-white hover:brightness-110 transition-all shadow-lg uppercase disabled:opacity-50"
                    >
                      {saving ? 'Criando...' : 'Criar Arquivo'}
                    </button>
                  </div>
                </motion.div>
              </div>
            )}
          </AnimatePresence>
          {/* Workspace Modal */}
          <AnimatePresence>
            {isWorkspaceModalOpen && (
              <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
                <motion.div 
                  initial={{ scale: 0.9, opacity: 0 }}
                  animate={{ scale: 1, opacity: 1 }}
                  exit={{ scale: 0.9, opacity: 0 }}
                  className="bg-dhl-sidebar border border-white/10 w-full max-w-xl rounded-2xl shadow-2xl overflow-hidden"
                >
                  <div className="px-6 py-4 border-b border-white/5 flex items-center justify-between bg-dhl-red">
                    <h3 className="text-sm font-bold text-white flex items-center gap-2 uppercase tracking-wide">
                      <FolderOpen className="w-4 h-4" />
                      Configurar Workspace Dinâmico
                    </h3>
                    <button onClick={() => setIsWorkspaceModalOpen(false)} className="text-white/60 hover:text-white transition-colors">
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="p-8 space-y-6">
                    <div>
                      <label className="block text-[10px] font-black text-white/40 uppercase mb-2 tracking-widest">URL da Pasta SharePoint</label>
                      <input 
                        type="text" 
                        value={tempWorkspaceUrl}
                        onChange={(e) => setTempWorkspaceUrl(e.target.value)}
                        placeholder="https://dpdhl.sharepoint.com/sites/SeuSite/Pasta"
                        className="w-full bg-black/40 border border-white/10 rounded-lg px-4 py-3 text-white text-sm outline-none focus:border-dhl-yellow/50 transition-all font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-black text-white/40 uppercase mb-2 tracking-widest">URL do Servidor IDE (Backend)</label>
                      <input 
                        type="text" 
                        value={tempIdeUrl}
                        onChange={(e) => setTempIdeUrl(e.target.value)}
                        placeholder="https://ais-dev-xxxx.run.app"
                        className="w-full bg-black/40 border border-white/10 rounded-lg px-4 py-3 text-white text-sm outline-none focus:border-dhl-yellow/50 transition-all font-mono"
                      />
                      <div className="flex gap-2 mt-2">
                         <p className="text-[9px] text-white/30 italic">Necessário para rodar comandos do Terminal (npm install/build).</p>
                         {!window.location.hostname.includes('sharepoint.com') && (
                           <button 
                             onClick={() => setTempIdeUrl(window.location.origin)}
                             className="text-[9px] text-dhl-yellow font-bold hover:underline ml-auto"
                           >
                             Usar URL atual
                           </button>
                         )}
                      </div>
                    </div>

                    <div className="bg-dhl-yellow/5 border border-dhl-yellow/10 p-4 rounded-xl space-y-2">
                       <p className="text-[11px] text-dhl-yellow font-bold uppercase mb-1">Dica de Uso:</p>
                       <p className="text-[11px] text-white/60 leading-relaxed">
                         1. No SharePoint, abra a pasta desejada e copie o link da barra de endereço.<br/>
                         2. No AI Studio, copie o URL do navegador (ex: https://ais-dev...) e cole em "Servidor IDE".
                       </p>
                    </div>
                  </div>
                  <div className="px-6 py-4 bg-black/10 border-t border-white/5 flex gap-3">
                    <button 
                      onClick={() => setIsWorkspaceModalOpen(false)}
                      className="flex-1 py-3 rounded-lg text-xs font-bold text-white/40 hover:text-white transition-all uppercase"
                    >
                      Cancelar
                    </button>
                    <button 
                      onClick={applyWorkspace}
                      className="flex-1 py-3 bg-dhl-yellow rounded-lg text-xs font-bold text-black hover:brightness-110 transition-all shadow-lg uppercase"
                    >
                      Conectar Workspace
                    </button>
                  </div>
                </motion.div>
              </div>
            )}
          </AnimatePresence>
        </section>
      </div>

      {/* Footer - Status Bar */}
      <footer className="h-6 shrink-0 bg-dhl-sidebar border-t border-white/5 z-50 flex items-center justify-between px-3 text-[10px] font-medium text-white/40">
        <div className="flex items-center gap-4 h-full">
          <div className="flex items-center gap-1.5 h-full px-2 hover:bg-white/5 cursor-pointer">
            <CheckCircle2 className="w-3 h-3 text-green-500" />
            <span>Ready</span>
          </div>
          <button 
            onClick={() => setIsTerminalOpen(!isTerminalOpen)}
            className="flex items-center gap-2 h-full px-2 hover:bg-white/5 cursor-pointer transition-colors"
          >
            <Terminal className="w-3 h-3 text-dhl-yellow" />
            <span>Terminal</span>
          </button>
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
