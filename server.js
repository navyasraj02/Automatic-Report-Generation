const express = require("express");
const fs = require("fs");
const csv = require("csv-parser");
const handlebars = require("handlebars");
const puppeteer = require("puppeteer");
const path = require("path");

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static("public"));

// Ensure generated_reports directory exists
if (!fs.existsSync("./generated_reports")) {
  fs.mkdirSync("./generated_reports");
}

// Helper function to convert image to base64
function imageToBase64(imagePath) {
  try {
    const imageBuffer = fs.readFileSync(imagePath);
    const base64Image = imageBuffer.toString("base64");
    const ext = path.extname(imagePath).toLowerCase();

    // Map file extensions to MIME types
    const mimeTypes = {
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".jfif": "image/jpeg",
      ".png": "image/png",
      ".gif": "image/gif",
    };

    const mimeType = mimeTypes[ext] || "image/jpeg";
    return `data:${mimeType};base64,${base64Image}`;
  } catch (error) {
    console.error("Error loading logo:", error);
    return "";
  }
}

// Helper function to read CSV data
function readCSV(filePath) {
  return new Promise((resolve, reject) => {
    const results = [];
    fs.createReadStream(filePath)
      .pipe(csv())
      .on("data", (data) => results.push(data))
      .on("end", () => resolve(results))
      .on("error", (error) => reject(error));
  });
}

// Get available experiences
app.get("/api/experiences", async (req, res) => {
  try {
    const goalData = await readCSV("./data/goal_setting.csv");
    const experiences = [
      ...new Set(goalData.map((row) => row.experience)),
    ].filter(Boolean);
    res.json({ experiences });
  } catch (error) {
    res.status(500).json({ error: "Failed to load experiences" });
  }
});

// Generate report endpoint
app.post("/api/generate-report", async (req, res) => {
  try {
    const { reportType, experience } = req.body;

    // Read all CSV files
    const goalData = await readCSV("./data/goal_setting.csv");
    const entryData = await readCSV("./data/entry_forms.csv");
    const exitData = await readCSV("./data/exit_forms.csv");

    // Filter data by experience
    const filteredGoals = goalData.filter(
      (row) => row.experience === experience
    );
    const filteredEntry = entryData.filter(
      (row) => row.experience === experience
    );
    const filteredExit = exitData.filter(
      (row) => row.experience === experience
    );

    // Calculate participation statistics
    const totalRegistered = filteredGoals.length; // Assuming all who set goals are registered
    const completedStudents = filteredEntry.length; // Those who completed entry forms
    const completionPercentage =
      totalRegistered > 0
        ? Math.round((completedStudents / totalRegistered) * 100)
        : 0;

    // Convert logo to base64
    const logoPath = path.join(__dirname, "assets", "logo.jfif"); // or logo.png, logo.jpg
    const logoBase64 = imageToBase64(logoPath);

    // Prepare template data
    const templateData = {
      experience,
      generatedDate: new Date().toLocaleDateString(),
      goals: filteredGoals,
      entryForms: filteredEntry,
      exitForms: filteredExit,
      totalParticipants: filteredGoals.length,
      totalRegistered,
      completedStudents,
      completionPercentage,
      logoBase64,
    };

    // Select template based on report type
    const templatePath =
      reportType === "profile"
        ? "./templates/profile_report.hbs"
        : "./templates/growth_report.hbs";

    // Read and compile template
    const templateSource = fs.readFileSync(templatePath, "utf8");
    const template = handlebars.compile(templateSource);
    const html = template(templateData);

    // Generate PDF using Puppeteer
    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const fileName = `${reportType}_${experience}_${Date.now()}.pdf`;
    const filePath = path.join(__dirname, "generated_reports", fileName);

    await page.pdf({
      path: filePath,
      format: "A4",
      landscape: true,
      printBackground: true,
      margin: {
        top: "0mm",
        right: "0mm",
        bottom: "0mm",
        left: "0mm",
      },
    });

    await browser.close();

    // Send file for download
    res.download(filePath, fileName, (err) => {
      if (err) {
        console.error("Download error:", err);
      }
    });
  } catch (error) {
    console.error("Report generation error:", error);
    res
      .status(500)
      .json({ error: "Failed to generate report", details: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
