const express = require("express");
const cors = require("cors");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(express.json());
app.use(cors());

// CONFIGURA ESTO:
const OUTLOOK_USER = "cristian.fortuny.acosta@escolamontserrat.cat";
const OUTLOOK_PASS = "R#291679631561onn";
const EXCEL_FILE = path.join(__dirname, "respuestas.xlsx");

// Config SMTP Outlook
const transporter = nodemailer.createTransport({
    host: "smtp.office365.com",
    port: 587,
    secure: false,
    auth: {
        user: OUTLOOK_USER,
        pass: OUTLOOK_PASS
    }
});

// Añadir fila a Excel
function appendToExcel(row) {
    let workbook;
    let worksheet;
    if (fs.existsSync(EXCEL_FILE)) {
        workbook = XLSX.readFile(EXCEL_FILE);
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        data.push(row);
        const newWs = XLSX.utils.aoa_to_sheet(data);
        workbook.Sheets[sheetName] = newWs;
    } else {
        workbook = XLSX.utils.book_new();
        const data = [
            ["timestamp", "email", "fase", "actividad", "respuesta"],
            row
        ];
        worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Respuestas");
    }
    XLSX.writeFile(workbook, EXCEL_FILE);
}

// Endpoint para recibir respuestas
app.post("/submit", async (req, res) => {
    const { email, fase, actividad, respuesta } = req.body;

    if (!email || !fase || !actividad || !respuesta) {
        return res.status(400).json({ ok: false, error: "Faltan campos" });
    }

    try {
        const timestamp = new Date().toISOString();

        // 1. Guardar en Excel
        appendToExcel([timestamp, email, fase, actividad, respuesta]);

        // 2. Enviar correo a Outlook
        await transporter.sendMail({
            from: OUTLOOK_USER,
            to: OUTLOOK_USER,
            subject: `Nueva respuesta – ${fase} / ${actividad}`,
            text: `
Alumno: ${email}
Fase: ${fase}
Actividad: ${actividad}

Respuesta:
${respuesta}
            `
        });

        res.json({ ok: true, message: "Respuesta enviada correctamente" });
    } catch (err) {
        console.error(err);
        res.status(500).json({ ok: false, error: "Error al procesar la respuesta" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Servidor escuchando en el puerto " + PORT));
