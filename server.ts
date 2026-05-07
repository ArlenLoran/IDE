import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { exec } from "child_process";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API para o Terminal
  app.post("/api/terminal/run", (req, res) => {
    const { command } = req.body;
    
    // Lista de comandos permitidos por segurança
    const allowedCommands = ["npm install", "npm run build", "ls -la", "pwd"];
    if (!allowedCommands.includes(command)) {
      return res.status(403).json({ error: "Comando não permitido" });
    }

    console.log(`Executando: ${command}`);
    
    exec(command, (error, stdout, stderr) => {
      res.json({
        stdout: stdout || "",
        stderr: stderr || "",
        error: error ? error.message : null
      });
    });
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
