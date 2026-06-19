import express from "express";
import path from "path";
import cors from "cors";
import { createServer as createViteServer } from "vite";
import { db } from "./src/db/index.ts";
import { materials, packages } from "./src/db/schema.ts";
import { eq } from "drizzle-orm";
import { requireAuth, AuthRequest } from "./src/middleware/auth.ts";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json({ limit: "50mb" }));

  // Public API to read materials
  app.get("/api/materials", async (req, res) => {
    try {
      const data = await db.select().from(materials);
      res.json(data);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to fetch materials" });
    }
  });

  // Admin API to save materials
  app.post("/api/materials", requireAuth, async (req: AuthRequest, res) => {
    try {
      const { id, code, name, content, excelData, images, notes } = req.body;
      const data = await db.insert(materials).values({
        id,
        code,
        name,
        content,
        notes,
        excelData: excelData || [],
        images: images || [],
      }).onConflictDoUpdate({
        target: materials.id,
        set: {
          code,
          name,
          content,
          notes,
          excelData: excelData || [],
          images: images || [],
        }
      }).returning();
      res.json(data[0]);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to save material" });
    }
  });

  app.delete("/api/materials/:id", requireAuth, async (req: AuthRequest, res) => {
    try {
      await db.delete(materials).where(eq(materials.id, req.params.id));
      res.json({ success: true });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to delete material" });
    }
  });

  // Public API to read packages
  app.get("/api/packages", async (req, res) => {
    try {
      const data = await db.select().from(packages);
      res.json(data);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to fetch packages" });
    }
  });

  // Admin API to save packages
  app.post("/api/packages", requireAuth, async (req: AuthRequest, res) => {
    try {
      const { id, name, materialIds, hiddenTags } = req.body;
      const data = await db.insert(packages).values({
        id,
        name,
        materialIds: materialIds || [],
        hiddenTags: hiddenTags || {}
      }).onConflictDoUpdate({
        target: packages.id,
        set: {
          name,
          materialIds: materialIds || [],
          hiddenTags: hiddenTags || {}
        }
      }).returning();
      res.json(data[0]);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to save package" });
    }
  });

  app.delete("/api/packages/:id", requireAuth, async (req: AuthRequest, res) => {
    try {
      await db.delete(packages).where(eq(packages.id, req.params.id));
      res.json({ success: true });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to delete package" });
    }
  });


  // Vite middleware for development
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
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
