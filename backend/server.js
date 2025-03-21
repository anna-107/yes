const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Excel File Path
const excelFilePath = "responses.xlsx";

// Function to ensure Excel file exists
async function createExcelFileIfNotExists() {
  try {
    if (!fs.existsSync(excelFilePath)) {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Responses");

      // Add headers
      worksheet.addRow([
        "Name",
        "Roll Number",
        "Email",
        "Phone Number",
        "Role",
        "Question 1",
        "Answer 1",
        "Question 2",
        "Answer 2",
        "Submitted At"
      ]);

      // Save the file
      await workbook.xlsx.writeFile(excelFilePath);
      console.log("Excel file created successfully.");
    }
  } catch (error) {
    console.error("Error creating Excel file:", error);
  }
}

// Run this function at the start
createExcelFileIfNotExists();

// Function to check for duplicate roll numbers in the Excel file
async function isDuplicateRollNumber(roll) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.getWorksheet("Responses");

    for (let i = 2; i <= worksheet.rowCount; i++) {
      const existingRoll = worksheet.getRow(i).getCell(2).value; // Roll number is in the 2nd column
      if (existingRoll === roll) {
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error("Error reading Excel file:", error);
    throw new Error("Failed to read Excel file. The file may be corrupted.");
  }
}

// API Endpoint to Save Responses
app.post("/submit", async (req, res) => {
  const { name, roll, email, phone, role, question1, answer1, question2, answer2 } = req.body;

  if (!name || !roll || !email || !phone || !role || !question1 || !answer1 || !question2 || !answer2) {
    return res.status(400).json({ status: "error", message: "All fields are required." });
  }

  try {
    if (await isDuplicateRollNumber(roll)) {
      return res.status(400).json({ status: "error", message: "This roll number has already been used." });
    }

    const timestamp = new Date().toISOString();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.getWorksheet("Responses");

    worksheet.addRow([name, roll, email, phone, role, question1, answer1, question2, answer2, timestamp]);

    await workbook.xlsx.writeFile(excelFilePath);

    res.json({ status: "success", message: "Response saved successfully in Excel!" });
  } catch (error) {
    console.error("Error saving to Excel:", error);

    // If the error is related to file corruption, recreate the file
    if (error.message.includes("Failed to read Excel file")) {
      console.log("Recreating the Excel file...");
      await createExcelFileIfNotExists();
    }

    res.status(500).json({ status: "error", message: "Failed to save response in Excel." });
  }
});

// Start Server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
