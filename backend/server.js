const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
app.use(express.json());
app.use(cors());

let submissions = []; // Temp in-memory data (will reset on Vercel function reload)

// Save data
app.post("/save-excel", (req, res) => {
    const { businessName, ownerName, email, phone, whatsapp, businessType } = req.body;
    submissions.push([businessName, ownerName, email, phone, whatsapp, businessType]);
    res.json({ message: "Thanks for submitting your details.\nTry our billing software with a 7-day free trial!" });
});

// Download as Excel (on the fly)
app.get("/download-excel", (req, res) => {
    const workbook = XLSX.utils.book_new();
    const worksheetData = [
        ["Business Name", "Business Owner Name", "Email", "Phone No", "Whatsapp No", "Type of Business"],
        ...submissions
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Enquiry Data");

    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Disposition", 'attachment; filename="Billing_Software_Enquiry.xlsx"');
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
});

module.exports = app;
