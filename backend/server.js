const express = require("express");
const fs = require("fs");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
const port = 3000;
const filePath = "Billing_Software_Enquiry_FROM_LANDINGPAGE.xlsx";

app.use(express.json());
app.use(cors());

app.post("/save-excel", (req, res) => {
    const { businessName, ownerName, email, phone, whatsapp, businessType } = req.body;

    let workbook;
    if (fs.existsSync(filePath)) {
        // Load existing workbook
        workbook = XLSX.readFile(filePath);
    } else {
        // Create a new workbook if file doesn't exist
        workbook = XLSX.utils.book_new();
    }

    let worksheet = workbook.Sheets["Enquiry Data"];

    if (!worksheet) {
        // Create a new sheet if it doesn't exist
        worksheet = XLSX.utils.aoa_to_sheet([
            ["Business Name", "Business Owner Name", "Email", "Phone No", "Whatsapp No", "Type of Business"]
        ]);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Enquiry Data");
    }

    // Convert sheet data to JSON and append new data
    let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    data.push([businessName, ownerName, email, phone, whatsapp, businessType]);

    // Update the worksheet
    const updatedSheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets["Enquiry Data"] = updatedSheet;

    // Save the updated workbook
    XLSX.writeFile(workbook, filePath);

    res.json({ message: "Thanks for submitting your details.\nTry our billing software with a 7-day free trial!" });
});

app.get("/download-excel", (req, res) => {
    if (fs.existsSync(filePath)) {
        res.download(filePath); // Sends the file for download
    } else {
        res.status(404).json({ message: "File not found!" });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
