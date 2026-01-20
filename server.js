const express = require("express");
const multer = require("multer");
const cors = require("cors");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");

const readExcel = require("./readExcel");

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer({ dest: "uploads/" });
let uploadedFilePath = null;

/* ================= UPLOAD ================= */
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    // 🔒 ล้างไฟล์เก่าถ้ามี (กัน export ผิดไฟล์)
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      try {
        fs.unlinkSync(uploadedFilePath);
      } catch (e) {
        console.warn("⚠️ Cannot delete old upload:", e.message);
      }
    }

    // ✅ rename ให้ตรงนามสกุลจริง
    const ext = path.extname(req.file.originalname) || ".xlsx";
    const newPath = path.resolve(req.file.path + ext);
    fs.renameSync(req.file.path, newPath);

    uploadedFilePath = newPath;

    const excelData = readExcel(uploadedFilePath);

    res.json({
      status: "success",
      data: excelData, // ✔ excelRowOffset มาจาก readExcel ใหม่
    });
  } catch (err) {
    console.error("❌ Upload error:", err);
    res.status(500).json({
      error: "Failed to read Excel",
      message: err.message,
    });
  }
});

/* ================= EXPORT ================= */
app.post("/export", (req, res) => {
  try {
    if (!uploadedFilePath || !fs.existsSync(uploadedFilePath)) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const edits = req.body.edits || [];

    if (!Array.isArray(edits) || edits.length === 0) {
      return res.status(400).json({ error: "No edits provided" });
    }

    const outputDir = path.resolve("output");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const outputPath = path.join(outputDir, "output.xlsx");

    /* ✅ 1. copy ไฟล์ต้นฉบับ */
    fs.copyFileSync(uploadedFilePath, outputPath);

    /* ✅ 2. เรียก Python แก้เฉพาะ cell */
    const pythonScript = path.resolve(__dirname, "python", "excel.py");

    const py = spawn("python", [
      pythonScript,
      outputPath,
      JSON.stringify(edits),
    ]);

    py.stdout.on("data", (data) => {
      console.log("🐍 Python:", data.toString());
    });

    py.stderr.on("data", (data) => {
      console.error("🐍 Python error:", data.toString());
    });

    py.on("close", (code) => {
      console.log("🐍 Python exit code:", code);

      if (code !== 0) {
        return res.status(500).json({
          error: "Excel export failed",
          code,
        });
      }

      /* ✅ 3. ส่งไฟล์กลับ */
      res.download(outputPath, "Hospital_Asset.xlsx");
    });
  } catch (err) {
    console.error("❌ Export error:", err);
    res.status(500).json({
      error: "Export failed",
      message: err.message,
    });
  }
});

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
  console.log(`🚀 Backend running on port ${PORT}`);
});


