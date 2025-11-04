import express from "express";
import formidable from "formidable";
import fs from "fs";
import { parse } from "csv-parse/sync";
import XLSX from "xlsx";
import nodemailer from "nodemailer";
import path from "path";
import { fileURLToPath } from "url";
import pLimit from "p-limit";

const app = express();
const PORT = process.env.PORT || 3000;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Serve the HTML frontend
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// POST route for sending emails
app.post("/send_emails", (req, res) => {
  const form = formidable({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(500).json({ status: "error", message: err.message });

    try {
      // --- Extract and validate email body ---
      let emailBody = fields.email_body;
      if (Array.isArray(emailBody)) emailBody = emailBody[0];
      emailBody = String(emailBody || "").trim();

      if (!emailBody) {
        return res.status(400).json({ status: "error", message: "Email body is required." });
      }

      // --- Validate uploaded file ---
      let uploadedFile = files.file;
      if (Array.isArray(uploadedFile)) uploadedFile = uploadedFile[0];
      if (!uploadedFile) {
        return res.status(400).json({ status: "error", message: "No file uploaded." });
      }

      const filePath = uploadedFile.filepath;
      const filename = uploadedFile.originalFilename;
      if (!filePath || !filename) {
        return res.status(400).json({ status: "error", message: "Invalid file upload." });
      }

      // --- Parse recipients ---
      let recipients = [];
      if (filename.toLowerCase().endsWith(".csv")) {
        const content = fs.readFileSync(filePath, "utf8");
        const records = parse(content, { skip_empty_lines: true });
        recipients = records.slice(1); // skip header
      } else if (filename.toLowerCase().endsWith(".xlsx")) {
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        recipients = data.slice(1);
      } else {
        return res.status(400).json({ status: "error", message: "Unsupported file type." });
      }

      if (!recipients.length) {
        return res.status(400).json({ status: "error", message: "No recipients found in file." });
      }

      // --- Configure transporter with pooling (fast!) ---
      const transporter = nodemailer.createTransport({
        host: "smtp.hostinger.com",
        port: 587,
        secure: false,
        auth: {
          user: "emma@laskontech.org",
          pass: "=VJSv4Ov$",
        },
        pool: true,
        maxConnections: 5, // 5 parallel SMTP connections
        maxMessages: 100,  // reuse connection for multiple sends
      });

      // --- Limit concurrency to avoid server overload ---
      const limit = pLimit(5); // send 5 emails concurrently

      // --- Send all emails in parallel (with concurrency control) ---
      await Promise.all(
        recipients.map(row => limit(async () => {
          const [name, recipient] = row;
          if (!recipient) return;

          const body = emailBody.replace("{{name}}", name || "");

          try {
            await transporter.sendMail({
              from: "emma@laskontech.org",
              to: recipient,
              subject: "Message from LaskonTech",
              html: body,
            });
            console.log(`✅ Sent to ${recipient}`);
          } catch (sendError) {
            console.error(`❌ Failed to ${recipient}:`, sendError.message);
          }
        }))
      );

      res.json({ status: "success", message: "Emails sent successfully!" });

    } catch (error) {
      console.error("❌ Error:", error);
      res.status(500).json({ status: "error", message: error.message });
    }
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
