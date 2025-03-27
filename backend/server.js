const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
app.use(express.json());
app.use(cors());

const filePath = path.join("/tmp", "Billing_Software_Enquiry.xlsx");

// Save Data to Excel
app.post("/save-excel", (req, res) => {
    const { businessName, ownerName, email, phone, whatsapp, businessType } = req.body;

    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
    } else {
        workbook = XLSX.utils.book_new();
    }

    let worksheet = workbook.Sheets["Enquiry Data"];
    if (!worksheet) {
        worksheet = XLSX.utils.aoa_to_sheet([
            ["Business Name", "Business Owner Name", "Email", "Phone No", "Whatsapp No", "Type of Business"]
        ]);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Enquiry Data");
    }

    let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    data.push([businessName, ownerName, email, phone, whatsapp, businessType]);

    const updatedSheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets["Enquiry Data"] = updatedSheet;

    XLSX.writeFile(workbook, filePath);

    res.json({ message: "Thanks for submitting your details.\nTry our billing software with a 7-day free trial!" });
});

// Download Excel File
app.get("/download-excel", (req, res) => {
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ message: "File not found!" });
    }

    res.download(filePath, "Billing_Software_Enquiry.xlsx");
});

// Export Express app as a Serverless Function
module.exports = app;
