import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { spawn } from "child_process";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // In-memory logs for terminal tasks
  const tasks: Record<string, { logs: string[], status: 'running' | 'completed' | 'failed' }> = {};

  // API para o Terminal
  app.post("/api/terminal/run", (req, res) => {
    const { command } = req.body;
    
    const allowedCommands = ["npm install", "npm run build", "ls -la", "pwd", "npm list"];
    const baseCommand = command.split(' ')[0] + ' ' + (command.split(' ')[1] || '');
    
    // Verificação relaxada para permitir argumentos extras se necessário
    if (!allowedCommands.some(c => command.startsWith(c))) {
      return res.status(403).json({ error: "Comando não permitido" });
    }

    const taskId = Math.random().toString(36).substring(7);
    tasks[taskId] = { logs: [`Iniciando: ${command}`], status: 'running' };

    console.log(`Executando Tarefa ${taskId}: ${command}`);
    
    const [cmd, ...args] = command.split(' ');
    const child = spawn(cmd, args, { shell: true });

    child.stdout.on('data', (data) => {
      tasks[taskId].logs.push(data.toString());
    });

    child.stderr.on('data', (data) => {
      tasks[taskId].logs.push(`STDERR: ${data.toString()}`);
    });

    child.on('close', (code) => {
      tasks[taskId].status = code === 0 ? 'completed' : 'failed';
      tasks[taskId].logs.push(`Processo finalizado com código ${code}`);
    });

    res.json({ taskId });
  });

  app.get("/api/terminal/logs/:taskId", (req, res) => {
    const { taskId } = req.params;
    const task = tasks[taskId];
    if (!task) return res.status(404).json({ error: "Tarefa não encontrada" });
    res.json(task);
  });

  // Vite middleware para desenvolvimento
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
