import express from "express";
import path from "path";
import cors from "cors";
import { createServer as createViteServer } from "vite";
import { db } from "./src/db/index";
import { materials, materialAuditLogs, packages } from "./src/db/schema";
import { eq, desc } from "drizzle-orm";
import { requireAuth, AuthRequest } from "./src/middleware/auth";
import crypto from "crypto";

async function initDb() {
  try {
    await db.execute(`
      CREATE TABLE IF NOT EXISTS materials (
        id TEXT PRIMARY KEY,
        code TEXT NOT NULL,
        name TEXT NOT NULL,
        content TEXT NOT NULL,
        notes TEXT,
        excel_data JSONB NOT NULL DEFAULT '{"data": [], "merges": [], "tags": {}}'::jsonb,
        images JSONB NOT NULL DEFAULT '[]'::jsonb,
        doc_requirements JSONB,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMP NOT NULL DEFAULT NOW(),
        updated_by TEXT
      );
    `);
    
    // Add missing columns if table already existed from older versions
    await db.execute(`
      ALTER TABLE materials
      ADD COLUMN IF NOT EXISTS excel_data JSONB NOT NULL DEFAULT '{"data": [], "merges": [], "tags": {}}'::jsonb,
      ADD COLUMN IF NOT EXISTS images JSONB NOT NULL DEFAULT '[]'::jsonb,
      ADD COLUMN IF NOT EXISTS doc_requirements JSONB;
    `);
    await db.execute(`
      CREATE TABLE IF NOT EXISTS material_audit_logs (
        id TEXT PRIMARY KEY,
        material_id TEXT NOT NULL,
        action TEXT NOT NULL,
        previous_data JSONB,
        new_data JSONB,
        updated_by TEXT,
        updated_at TIMESTAMP NOT NULL DEFAULT NOW()
      );
    `);
    await db.execute(`
      CREATE TABLE IF NOT EXISTS packages (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        material_ids JSONB NOT NULL,
        hidden_tags JSONB,
        hidden_tables JSONB,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      );
    `);
    try {
      await db.execute(`ALTER TABLE packages ADD COLUMN IF NOT EXISTS hidden_tables JSONB;`);
    } catch (e: any) {
      console.warn("Could not alter packages table to add hidden_tables:", e.message);
    }
    console.log("Database initialized successfully");
  } catch (err: any) {
    console.warn("Database initialization skipped (or failed):", err.message);
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'Accept']
  }));
  app.use(express.json({ limit: "50mb" }));

  // Initialize DB tables
  await initDb();

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
      const { id, code, name, content, excelData, images, notes, docRequirements } = req.body;
      const userEmail = req.user?.email || 'Unknown User';
      
      // Check if material exists to log context correctly
      let existingMaterial = null;
      try {
         const rows = await db.select().from(materials).where(eq(materials.id, id));
         if (rows.length > 0) {
            existingMaterial = rows[0];
         }
      } catch (err) {
         console.warn("Could not fetch existing material", err);
      }

      const data = await db.insert(materials).values({
        id,
        code: code || '',
        name: name || '',
        content: content || '',
        notes: notes || '',
        excelData: excelData || { data: [], merges: [], tags: {} },
        images: images || [],
        docRequirements: docRequirements || {},
        updatedBy: userEmail,
      }).onConflictDoUpdate({
        target: materials.id,
        set: {
          code: code || '',
          name: name || '',
          content: content || '',
          notes: notes || '',
          excelData: excelData || { data: [], merges: [], tags: {} },
          images: images || [],
          docRequirements: docRequirements || {},
          updatedBy: userEmail,
          updatedAt: new Date()
        }
      }).returning();
      
      // Log the audit event
      try {
        const simplifyData = (d: any) => {
          if (!d) return null;
          const { images, excelData, excel_data, ...rest } = d;
          return { ...rest, _note: "Images and formatting data omitted for brevity" };
        };

        await db.insert(materialAuditLogs).values({
           id: crypto.randomUUID(),
           materialId: id,
           action: existingMaterial ? 'update' : 'create',
           previousData: simplifyData(existingMaterial),
           newData: simplifyData(data[0]),
           updatedBy: userEmail
        });
      } catch(auditErr) {
        console.error("Failed to write audit log:", auditErr);
      }

      res.json(data[0]);
    } catch (error: any) {
      console.error(error);
      res.status(500).json({ error: "Failed to save material", details: error.message });
    }
  });

  app.delete("/api/materials/:id", requireAuth, async (req: AuthRequest, res) => {
    try {
      const materialId = req.params.id;
      const userEmail = req.user?.email || 'Unknown User';

      // Keep previous data for the log
      let existingMaterial = null;
      try {
         const rows = await db.select().from(materials).where(eq(materials.id, materialId));
         if (rows.length > 0) {
            existingMaterial = rows[0];
         }
      } catch (err) {}

      await db.delete(materials).where(eq(materials.id, materialId));
      
      if (existingMaterial) {
          try {
             const simplifyData = (d: any) => {
               if (!d) return null;
               const { images, excelData, excel_data, ...rest } = d;
               return { ...rest, _note: "Images and formatting data omitted for brevity" };
             };

             await db.insert(materialAuditLogs).values({
                id: crypto.randomUUID(),
                materialId: materialId,
                action: 'delete',
                previousData: simplifyData(existingMaterial),
                newData: null,
                updatedBy: userEmail
             });
          } catch(auditErr) {
             console.error("Failed to write audit log:", auditErr);
          }
      }

      res.json({ success: true });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to delete material" });
    }
  });

  // Audit Logs API
  app.get("/api/materials/:id/audit-logs", async (req, res) => {
     try {
        const materialId = req.params.id;
        const logs = await db.select().from(materialAuditLogs)
           .where(eq(materialAuditLogs.materialId, materialId))
           .orderBy(desc(materialAuditLogs.updatedAt));
        res.json(logs);
     } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Failed to fetch audit logs" });
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
      const { id, name, materialIds, hiddenTags, hiddenTables } = req.body;
      const data = await db.insert(packages).values({
        id,
        name,
        materialIds: materialIds || [],
        hiddenTags: hiddenTags || {},
        hiddenTables: hiddenTables || {}
      }).onConflictDoUpdate({
        target: packages.id,
        set: {
          name,
          materialIds: materialIds || [],
          hiddenTags: hiddenTags || {},
          hiddenTables: hiddenTables || {}
        }
      }).returning();
      res.json(data[0]);
    } catch (error: any) {
      console.error("Failed to save package:", error);
      // Fallback in case "hidden_tables" column does not exist in their custom database yet
      if (error?.message && (error.message.includes("hidden_tables") || error.message.includes("column"))) {
        try {
          console.log("Attempting fallback saving without hiddenTables column...");
          const { id, name, materialIds, hiddenTags } = req.body;
          const data = await db.insert(packages).values({
            id,
            name,
            materialIds: materialIds || [],
            hiddenTags: hiddenTags || {}
          } as any).onConflictDoUpdate({
            target: packages.id,
            set: {
              name,
              materialIds: materialIds || [],
              hiddenTags: hiddenTags || {}
            } as any
          }).returning();
          return res.json(data[0]);
        } catch (fbErr: any) {
          console.error("Fallback save failed too:", fbErr);
          return res.status(500).json({ error: fbErr.message || "Failed to save package (fallback)" });
        }
      }
      res.status(500).json({ error: error?.message || "Failed to save package" });
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
