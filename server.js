const express = require("express");
const fs = require("fs");
const csv = require("csv-parser");
const handlebars = require("handlebars");
const puppeteer = require("puppeteer");
const path = require("path");
const { MongoClient } = require("mongodb");

// Load .env in local dev (optional)
try {
  // eslint-disable-next-line global-require
  require("dotenv").config();
} catch {
  // dotenv is optional; ignore if not installed
}

const app = express();
const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;

app.use(express.json());
app.use(express.static("public"));

// Ensure generated_reports directory exists
if (!fs.existsSync("./generated_reports")) {
  fs.mkdirSync("./generated_reports");
}

function getMongoDbName() {
  return process.env.MONGODB_DB_NAME || undefined;
}

let mongoClient;
let mongoClientPromise;

async function getMongoDb() {
  const uri = process.env.MONGODB_URI;
  if (!uri) {
    throw new Error(
      "Missing MONGODB_URI. Set MONGODB_URI (and optionally MONGODB_DB_NAME).",
    );
  }

  if (!mongoClientPromise) {
    mongoClient = new MongoClient(uri, {
      // Sensible defaults; topology engine is automatic in modern driver
      maxPoolSize: 10,
    });
    mongoClientPromise = mongoClient.connect();
  }

  const client = await mongoClientPromise;
  return client.db(getMongoDbName());
}

// Helper function to convert image to base64
function imageToBase64(imagePath) {
  try {
    const imageBuffer = fs.readFileSync(imagePath);
    const base64Image = imageBuffer.toString("base64");
    const ext = path.extname(imagePath).toLowerCase();
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

// Helper function to normalize email for comparison
function normalizeEmail(email) {
  if (!email) return "";
  return email.toLowerCase().trim();
}

// Get available sessions (from sessionData)
app.get("/api/sessions", async (req, res) => {
  try {
    if (!process.env.MONGODB_URI) {
      return res
        .status(500)
        .json({ error: "MongoDB is not configured for sessions." });
    }

    const db = await getMongoDb();
    const sessionCollection = db.collection("sessionData");

    // Prefer active sessions; if none, fall back to all
    let sessions = await sessionCollection
      .find({ sessionStatus: true })
      .project({ _id: 1, sessionName: 1 })
      .toArray();

    if (!sessions.length) {
      sessions = await sessionCollection
        .find({})
        .project({ _id: 1, sessionName: 1 })
        .toArray();
    }

    const result = sessions
      .map((doc) => ({
        id: String(doc._id),
        name: doc.sessionName || String(doc._id),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return res.json({ sessions: result });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to load sessions",
      details: error.message,
    });
  }
});

// Get available experiences for a specific session
app.get("/api/experiences", async (req, res) => {
  try {
    if (!process.env.MONGODB_URI) {
      return res
        .status(500)
        .json({ error: "MongoDB is not configured for experiences." });
    }

    const { sessionId } = req.query;

    if (!sessionId) {
      return res.status(400).json({ error: "sessionId query parameter is required" });
    }

    const db = await getMongoDb();
    const expInstanceCollection = db.collection("expInstanceData");

    // Active experiences for this session
    const instances = await expInstanceCollection
      .find({ sessionID: sessionId, expInstanceStatus: true })
      .project({ _id: 1, experience: 1 })
      .toArray();

    const experiences = instances
      .map((doc) => {
        const expObj = doc.experience || {};
        const id = expObj.id ? String(expObj.id) : String(doc._id);
        const name =
          expObj.name === undefined || expObj.name === null
            ? ""
            : String(expObj.name).trim();
        const category =
          expObj.category === undefined || expObj.category === null
            ? ""
            : String(expObj.category).trim();
        const label = `${name} ${category}`.replace(/\s+/g, " ").trim();
        if (!label) return null;
        return { id, label };
      })
      .filter(Boolean);

    // De‑duplicate by id and sort by label
    const uniqueById = new Map();
    experiences.forEach((exp) => {
      if (!uniqueById.has(exp.id)) {
        uniqueById.set(exp.id, exp);
      }
    });

    const sorted = Array.from(uniqueById.values()).sort((a, b) =>
      a.label.localeCompare(b.label),
    );

    return res.json({ experiences });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to load experiences",
      details: error.message,
    });
  }
});

// Helper function to safely access nested object properties
function getNestedValue(obj, path) {
  if (!obj || !path) return undefined;
  const parts = path.split(".");
  let current = obj;
  for (const part of parts) {
    if (current === null || current === undefined) return undefined;
    current = current[part];
  }
  return current;
}

// Helper function to calculate distribution percentages
// Supports both flat fields (e.g., "field") and nested paths (e.g., "studentInformation.enrolledUHInfo.majors")
function calculateDistribution(data, field) {
  const counts = {};
  let total = 0;

  data.forEach((row) => {
    const raw = getNestedValue(row, field);
    const value =
      raw === undefined || raw === null ? "" : String(raw).trim();
    if (value !== "") {
      counts[value] = (counts[value] || 0) + 1;
      total++;
    }
  });

  return Object.entries(counts)
    .map(([key, count]) => ({
      label: key,
      count: count,
      percentage: total > 0 ? Math.round((count / total) * 100) : 0,
    }))
    .sort((a, b) => b.percentage - a.percentage);
}

// Helper function to calculate distribution for array fields (e.g., majors array)
function calculateDistributionFromArray(data, field) {
  const counts = {};
  let total = 0;

  data.forEach((row) => {
    const raw = getNestedValue(row, field);
    if (Array.isArray(raw)) {
      raw.forEach((item) => {
        const value =
          item === undefined || item === null ? "" : String(item).trim();
        if (value !== "") {
          counts[value] = (counts[value] || 0) + 1;
          total++;
        }
      });
    } else if (raw !== undefined && raw !== null) {
      const value = String(raw).trim();
      if (value !== "") {
        counts[value] = (counts[value] || 0) + 1;
        total++;
      }
    }
  });

  return Object.entries(counts)
    .map(([key, count]) => ({
      label: key,
      count: count,
      percentage: total > 0 ? Math.round((count / total) * 100) : 0,
    }))
    .sort((a, b) => b.percentage - a.percentage);
}

// Helper function to calculate boolean distribution
function calculateBooleanDistribution(data, field, trueLabel = "Yes", falseLabel = "No") {
  const counts = { [trueLabel]: 0, [falseLabel]: 0 };
  let total = 0;

  data.forEach((row) => {
    const raw = getNestedValue(row, field);
    if (raw !== undefined && raw !== null) {
      const value = Boolean(raw);
      counts[value ? trueLabel : falseLabel]++;
      total++;
    }
  });

  return Object.entries(counts)
    .map(([key, count]) => ({
      label: key,
      count: count,
      percentage: total > 0 ? Math.round((count / total) * 100) : 0,
    }))
    .sort((a, b) => b.percentage - a.percentage);
}

// Helper function to calculate growth data
function calculateGrowthData(goalData, exitData, expectedField, achievedField) {
  const categories = ["None", "Little", "Moderate", "A lot"];
  const expectedCounts = { None: 0, Little: 0, Moderate: 0, "A lot": 0 };
  const achievedCounts = { None: 0, Little: 0, Moderate: 0, "A lot": 0 };

  goalData.forEach((row) => {
    const value = row[expectedField];
    if (value && expectedCounts.hasOwnProperty(value)) {
      expectedCounts[value]++;
    }
  });

  exitData.forEach((row) => {
    const value = row[achievedField];
    if (value && achievedCounts.hasOwnProperty(value)) {
      achievedCounts[value]++;
    }
  });

  const totalExpected = goalData.length || 1;
  const totalAchieved = exitData.length || 1;

  return categories.map((cat) => ({
    category: cat,
    expected: Math.round((expectedCounts[cat] / totalExpected) * 100),
    achieved: Math.round((achievedCounts[cat] / totalAchieved) * 100),
  }));
}

// Generate report endpoint
app.post("/api/generate-report", async (req, res) => {
  try {
    const {
      reportType,
      sessionId,
      sessionLabel,
      experienceId,
      experienceLabel,
      instructorName,
    } = req.body;

    const sessionValue =
      sessionId === undefined || sessionId === null
        ? ""
        : String(sessionId).trim();
    const sessionDisplay =
      sessionLabel === undefined || sessionLabel === null
        ? sessionValue
        : String(sessionLabel).trim() || sessionValue;

    const experienceValue =
      experienceId === undefined || experienceId === null
        ? ""
        : String(experienceId).trim();
    const experienceDisplay =
      experienceLabel === undefined || experienceLabel === null
        ? experienceValue
        : String(experienceLabel).trim() || experienceValue;

    if (!sessionValue) {
      return res.status(400).json({ error: "session is required" });
    }

    if (!experienceValue) {
      return res.status(400).json({ error: "experience is required" });
    }

    // Convert logo to base64
    const logoPath = path.join(__dirname, "assets", "logo.jfif");
    const logoBase64 = imageToBase64(logoPath);

    let templateData;

    if (reportType === "profile") {
      // ============= PROFILE REPORT DATA =============

      const db = await getMongoDb();
      const goalCollection = db.collection("goalSettingFormData");
      const entryCollection = db.collection("studentEntryFormData");

      const goalPipeline = [
        { $match: { completed: true } },
        {
          $lookup: {
            from: "expRegistrationData",
            localField: "expRegistrationID",
            foreignField: "_id",
            as: "reg",
          },
        },
        { $unwind: "$reg" },
        {
          $lookup: {
            from: "expInstanceData",
            localField: "reg.expInstanceID",
            foreignField: "_id",
            as: "inst",
          },
        },
        { $unwind: "$inst" },
        {
          $lookup: {
            from: "sessionData",
            localField: "inst.sessionID",
            foreignField: "_id",
            as: "sess",
          },
        },
        { $unwind: "$sess" },
        {
          $match: {
            "sess._id": sessionValue,
            "sess.sessionName": sessionDisplay,
            "inst.experience.id": experienceValue,
            "inst.expInstanceStatus": true,
          },
        },
        { $count: "count" },
      ];

      const registeredPipeline = [
        { $match: { completed: true } },
        {
          $lookup: {
            from: "expRegistrationData",
            localField: "organizationID",
            foreignField: "_id",
            as: "reg",
          },
        },
        { $unwind: "$reg" },
        {
          $lookup: {
            from: "expInstanceData",
            localField: "reg.expInstanceID",
            foreignField: "_id",
            as: "inst",
          },
        },
        { $unwind: "$inst" },
        {
          $lookup: {
            from: "sessionData",
            localField: "inst.sessionID",
            foreignField: "_id",
            as: "sess",
          },
        },
        { $unwind: "$sess" },
        {
          $match: {
            "sess._id": sessionValue,
            "sess.sessionName": sessionDisplay,
            "inst.experience.id": experienceValue,
            "inst.expInstanceStatus": true,
          },
        },
        { $count: "count" },
      ];

      const [goalAgg, registeredAgg, filteredGoals, filteredEntry] =
        await Promise.all([
          goalCollection.aggregate(goalPipeline).toArray(),
          entryCollection.aggregate(registeredPipeline).toArray(),
          goalCollection.find({ expRegistrationID: experienceValue }).toArray(),
          entryCollection.find({ organizationID: experienceValue }).toArray(),
        ]);

      const completedStudents =
        (goalAgg && goalAgg.length > 0 && goalAgg[0].count) || 0;
      const totalRegistered =
        (registeredAgg &&
          registeredAgg.length > 0 &&
          registeredAgg[0].count) ||
        0;

      // Calculate completion percentage
      const completionPercentage =
        totalRegistered > 0
          ? Math.round((completedStudents / totalRegistered) * 100)
          : 0;

      // Calculate Demographics
      // Note: Some fields (Classification, Gender, Ethnicity, First Generation, International)
      // are not present in the provided schema, so they return empty distributions
      const demographics = {
        classification: [], // Not in schema
        gender: [], // Not in schema
        ethnicity: [], // Not in schema
        minors: (() => {
          // Combine honorsMinors and otherMinors arrays
          const allMinors = [];
          filteredEntry.forEach((entry) => {
            const honorsMinors = getNestedValue(
              entry,
              "studentInformation.enrolledUHInfo.honorsMinors",
            );
            const otherMinors = getNestedValue(
              entry,
              "studentInformation.enrolledUHInfo.otherMinors",
            );
            if (Array.isArray(honorsMinors)) {
              allMinors.push(...honorsMinors);
            }
            if (Array.isArray(otherMinors)) {
              allMinors.push(...otherMinors);
            }
            if (!Array.isArray(honorsMinors) && honorsMinors) {
              allMinors.push(honorsMinors);
            }
            if (!Array.isArray(otherMinors) && otherMinors) {
              allMinors.push(otherMinors);
            }
          });
          const counts = {};
          let total = 0;
          allMinors.forEach((minor) => {
            const value =
              minor === undefined || minor === null
                ? ""
                : String(minor).trim();
            if (value !== "") {
              counts[value] = (counts[value] || 0) + 1;
              total++;
            }
          });
          return Object.entries(counts)
            .map(([key, count]) => ({
              label: key,
              count: count,
              percentage: total > 0 ? Math.round((count / total) * 100) : 0,
            }))
            .sort((a, b) => b.percentage - a.percentage);
        })(),
        firstGeneration: [], // Not in schema
        international: [], // Not in schema
      };

      // Calculate General Profile
      const generalProfile = {
        graduationYear: calculateDistribution(
          filteredEntry,
          "studentInformation.enrolledUHInfo.expectedGraduationYear",
        ),
        majors: calculateDistributionFromArray(
          filteredEntry,
          "studentInformation.enrolledUHInfo.majors",
        ),
        minors: (() => {
          // Combine honorsMinors and otherMinors arrays
          const allMinors = [];
          filteredEntry.forEach((entry) => {
            const honorsMinors = getNestedValue(
              entry,
              "studentInformation.enrolledUHInfo.honorsMinors",
            );
            const otherMinors = getNestedValue(
              entry,
              "studentInformation.enrolledUHInfo.otherMinors",
            );
            if (Array.isArray(honorsMinors)) {
              allMinors.push(...honorsMinors);
            }
            if (Array.isArray(otherMinors)) {
              allMinors.push(...otherMinors);
            }
            if (!Array.isArray(honorsMinors) && honorsMinors) {
              allMinors.push(honorsMinors);
            }
            if (!Array.isArray(otherMinors) && otherMinors) {
              allMinors.push(otherMinors);
            }
          });
          const counts = {};
          let total = 0;
          allMinors.forEach((minor) => {
            const value =
              minor === undefined || minor === null
                ? ""
                : String(minor).trim();
            if (value !== "") {
              counts[value] = (counts[value] || 0) + 1;
              total++;
            }
          });
          return Object.entries(counts)
            .map(([key, count]) => ({
              label: key,
              count: count,
              percentage: total > 0 ? Math.round((count / total) * 100) : 0,
            }))
            .sort((a, b) => b.percentage - a.percentage);
        })(),
        firstGeneration: [], // Not in schema
        international: [], // Not in schema
        housing: calculateBooleanDistribution(
          filteredEntry,
          "studentInformation.enrolledUHInfo.livingOnCampus",
          "On Campus",
          "Off Campus",
        ),
        honorsCollegeAffiliation: calculateBooleanDistribution(
          filteredEntry,
          "studentInformation.enrolledUHInfo.honorsCollegeAffiliatedStatus",
          "Yes",
          "No",
        ),
        communityService: calculateBooleanDistribution(
          filteredEntry,
          "studentInformation.communityServiceInfo.serviceStatus",
          "Yes",
          "No",
        ),
        graduateSchoolInterest: (() => {
          // Extract from programGradProType object (has id, label, checked)
          const categories = { Masters: 0, PhD: 0, "MD/DO": 0, Other: 0 };
          let total = 0;
          filteredEntry.forEach((entry) => {
            const gradType = getNestedValue(
              entry,
              "studentInformation.graduateProfessionalSchool.programGradProType",
            );
            if (gradType && typeof gradType === "object") {
              const label =
                gradType.label ||
                gradType.id ||
                (gradType.checked ? "Yes" : null);
              if (label) {
                const normalized = String(label).trim().toLowerCase();
                if (normalized.includes("master")) {
                  categories.Masters++;
                } else if (normalized.includes("phd") || normalized.includes("ph.d")) {
                  categories.PhD++;
                } else if (
                  normalized.includes("md") ||
                  normalized.includes("do") ||
                  normalized.includes("medical")
                ) {
                  categories["MD/DO"]++;
                } else {
                  categories.Other++;
                }
                total++;
              }
            }
          });
          return Object.entries(categories)
            .map(([key, count]) => ({
              label: key,
              count: count,
              percentage: total > 0 ? Math.round((count / total) * 100) : 0,
            }))
            .filter((item) => item.count > 0 || total === 0)
            .sort((a, b) => b.percentage - a.percentage);
        })(),
      };

      // Academic Profile - with dummy data
      const academicProfile = {
        gpa: [
          { label: "3.5-4.0", percentage: 45 },
          { label: "3.0-3.49", percentage: 35 },
          { label: "2.5-2.99", percentage: 15 },
          { label: "Below 2.5", percentage: 5 },
        ],
        creditHours: [
          { label: "12-15", percentage: 60 },
          { label: "16-18", percentage: 30 },
          { label: "9-11", percentage: 10 },
        ],
        research: [
          { label: "Yes", percentage: 40 },
          { label: "No", percentage: 60 },
        ],
        internship: [
          { label: "Yes", percentage: 35 },
          { label: "No", percentage: 65 },
        ],
      };

      // Leadership Profile - with dummy data
      const leadershipProfile = {
        organizations: [
          { label: "1-2", percentage: 45 },
          { label: "3-4", percentage: 30 },
          { label: "5+", percentage: 15 },
          { label: "None", percentage: 10 },
        ],
        positions: [
          { label: "Officer", percentage: 25 },
          { label: "Member", percentage: 50 },
          { label: "President", percentage: 15 },
          { label: "None", percentage: 10 },
        ],
        volunteerHours: [
          { label: "10-20", percentage: 40 },
          { label: "20-40", percentage: 30 },
          { label: "40+", percentage: 20 },
          { label: "Less than 10", percentage: 10 },
        ],
        serviceProjects: [
          { label: "1-2", percentage: 50 },
          { label: "3-4", percentage: 30 },
          { label: "5+", percentage: 20 },
        ],
      };

      // Research Profile - with dummy data
      const researchProfile = {
        projects: [
          { label: "1", percentage: 40 },
          { label: "2", percentage: 30 },
          { label: "3+", percentage: 20 },
          { label: "None", percentage: 10 },
        ],
        publications: [
          { label: "None", percentage: 70 },
          { label: "1", percentage: 20 },
          { label: "2+", percentage: 10 },
        ],
        presentations: [
          { label: "None", percentage: 60 },
          { label: "1-2", percentage: 30 },
          { label: "3+", percentage: 10 },
        ],
        posters: [
          { label: "None", percentage: 65 },
          { label: "1-2", percentage: 25 },
          { label: "3+", percentage: 10 },
        ],
      };

      templateData = {
        experience: experienceDisplay,
        session: sessionDisplay,
        instructorName,
        generatedDate: new Date().toLocaleDateString(),
        logoBase64,
        totalRegistered,
        completedStudents,
        completionPercentage,
        demographics,
        generalProfile,
        academicProfile,
        leadershipProfile,
        researchProfile,
      };
    } else {
      // ============= GROWTH REPORT DATA =============

      // For now, growth report remains CSV-backed (existing behavior).
      if (
        !fs.existsSync("./data/goal_setting.csv") ||
        !fs.existsSync("./data/entry_forms.csv") ||
        !fs.existsSync("./data/exit_forms.csv")
      ) {
        return res.status(500).json({
          error:
            "Growth report data files are missing. Configure MongoDB for growth reports or restore the CSVs.",
        });
      }

      const goalData = await readCSV("./data/goal_setting.csv");
      const entryData = await readCSV("./data/entry_forms.csv");
      const exitData = await readCSV("./data/exit_forms.csv");

      const filteredGoals = goalData.filter((row) => row.Experience === experienceValue);
      const filteredEntry = entryData.filter((row) => row.Experience === experienceValue);
      const filteredExit = exitData.filter((row) => row.Experience === experienceValue);

      const totalRegistered = filteredEntry.length;
      const goalSettingCompleted = filteredGoals.length;
      const exitFormCompleted = filteredExit.length;
      const goalSettingPercentage =
        totalRegistered > 0
          ? Math.round((goalSettingCompleted / totalRegistered) * 100)
          : 0;
      const exitFormPercentage =
        totalRegistered > 0
          ? Math.round((exitFormCompleted / totalRegistered) * 100)
          : 0;

      let totalGoals = 0;
      filteredGoals.forEach((row) => {
        if (row["Goal 1"] && row["Goal 1"].trim()) totalGoals++;
        if (row["Goal 2"] && row["Goal 2"].trim()) totalGoals++;
        if (row["Goal 3"] && row["Goal 3"].trim()) totalGoals++;
      });

      const studentGrowth = {
        teamwork: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Teamwork Expected",
          "Teamwork Achieved",
        ),
        professionalResponsibility: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Professional Responsibility Expected",
          "Professional Responsibility Achieved",
        ),
        effectiveCommunication: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Effective Communication Expected",
          "Effective Communication Achieved",
        ),
        problemSolving: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Problem Solving Expected",
          "Problem Solving Achieved",
        ),
        culturalHumility: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Cultural Humility Expected",
          "Cultural Humility Achieved",
        ),
        ethicalDecisionMaking: calculateGrowthData(
          filteredGoals,
          filteredExit,
          "Ethical Decision Making Expected",
          "Ethical Decision Making Achieved",
        ),
      };

      const progressCategories = ["None", "Little", "Some", "Lots"];
      const progressCounts = { None: 0, Little: 0, Some: 0, Lots: 0 };
      filteredExit.forEach((row) => {
        ["Goal 1 Progress", "Goal 2 Progress", "Goal 3 Progress"].forEach(
          (field) => {
            const value = row[field];
            if (value && progressCounts.hasOwnProperty(value)) {
              progressCounts[value]++;
            }
          },
        );
      });

      const totalProgressResponses =
        Object.values(progressCounts).reduce((a, b) => a + b, 0) || 1;
      const progressTowardsGoals = progressCategories.map((cat) => ({
        category: cat,
        percentage: Math.round(
          (progressCounts[cat] / totalProgressResponses) * 100,
        ),
      }));

      const connectionCategories = ["None", "Partly", "Largely"];
      const connectionCounts = { None: 0, Partly: 0, Largely: 0 };
      filteredExit.forEach((row) => {
        ["Goal 1 Connection", "Goal 2 Connection", "Goal 3 Connection"].forEach(
          (field) => {
            const value = row[field];
            if (value && connectionCounts.hasOwnProperty(value)) {
              connectionCounts[value]++;
            }
          },
        );
      });

      const totalConnectionResponses =
        Object.values(connectionCounts).reduce((a, b) => a + b, 0) || 1;
      const experienceConnection = connectionCategories.map((cat) => ({
        category: cat,
        percentage: Math.round(
          (connectionCounts[cat] / totalConnectionResponses) * 100,
        ),
      }));

      const npsFields = [
        "Complete Minor Likelihood",
        "Repeat Experience Likelihood",
        "Pursue Career Likelihood",
        "Recommend Friend Likelihood",
      ];

      const netPromoterScores = npsFields.map((field, index) => {
        const likelyCounts = filteredExit.filter(
          (row) => row[field] === "Likely" || row[field] === "Extremely likely",
        ).length;
        const percentage =
          filteredExit.length > 0
            ? Math.round((likelyCounts / filteredExit.length) * 100)
            : 0;
        return { category: String.fromCharCode(65 + index), percentage };
      });

      const netPromoterDescriptions = [
        {
          label: "A",
          text: `${netPromoterScores[0].percentage}% of students reported they are likely to extremely likely to complete the Data & Society minor`,
        },
        {
          label: "B",
          text: `${netPromoterScores[1].percentage}% reported they are likely to extremely likely to enroll in another course/repeat the experience`,
        },
        {
          label: "C",
          text: `${netPromoterScores[2].percentage}% reported they are extremely likely to pursue a career in data science`,
        },
        {
          label: "D",
          text: `${netPromoterScores[3].percentage}% reported they are likely to extremely likely to recommend this course/experience to a friend`,
        },
      ];

      const activityFields = [
        "Activity Class Discussion",
        "Activity HPE DSI Short Course",
        "Activity Class Presentations",
        "Activity Weekly OneNote",
        "Activity Peer Feedback",
        "Activity Extra Credit",
      ];

      const activitiesScores = activityFields.map((field, index) => {
        const yesCounts = filteredExit.filter(
          (row) => row[field] === "Yes",
        ).length;
        const percentage =
          filteredExit.length > 0
            ? Math.round((yesCounts / filteredExit.length) * 100)
            : 0;
        return { category: String.fromCharCode(65 + index), percentage };
      });

      const activitiesDescriptions = [
        {
          label: "A",
          text: `${activitiesScores[0].percentage}% Class discussion`,
        },
        {
          label: "B",
          text: `${activitiesScores[1].percentage}% HPE DSI Short Course`,
        },
        {
          label: "C",
          text: `${activitiesScores[2].percentage}% Class Presentations`,
        },
        {
          label: "D",
          text: `${activitiesScores[3].percentage}% Weekly OneNote check ins`,
        },
        {
          label: "E",
          text: `${activitiesScores[4].percentage}% Anonymous Peer feedback (Teammates)`,
        },
        {
          label: "F",
          text: `${activitiesScores[5].percentage}% Extra Credit Report`,
        },
      ];

      const biggestLessons = filteredExit
        .map((row) => row["Biggest Lesson Learned"])
        .filter((val) => val && val.trim());
      const experienceContributions = filteredExit
        .map((row) => row["Experience Contribution"])
        .filter((val) => val && val.trim());
      const additionalComments = filteredExit
        .map((row) => row["Additional Comments"])
        .filter((val) => val && val.trim());

      templateData = {
        experience: experienceValue,
        generatedDate: new Date().toLocaleDateString(),
        logoBase64,
        totalRegistered,
        goalSettingCompleted,
        goalSettingPercentage,
        exitFormCompleted,
        exitFormPercentage,
        totalGoals,
        studentGrowth,
        progressTowardsGoals,
        experienceConnection,
        netPromoterScores,
        netPromoterDescriptions,
        activitiesScores,
        activitiesDescriptions,
        biggestLessons,
        experienceContributions,
        additionalComments,
      };
    }

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

    const safeSession = (sessionDisplay || sessionValue || "session")
      .replace(/[^\w\-]+/g, "-")
      .toLowerCase();
    const safeExperience = (experienceDisplay || experienceValue || "experience")
      .replace(/[^\w\-]+/g, "-")
      .toLowerCase();
    const fileName = `${reportType}_${safeSession}_${safeExperience}.pdf`;
    const filePath = path.join(__dirname, "generated_reports", fileName);

    await page.pdf({
      path: filePath,
      format: "A4",
      landscape: true,
      printBackground: true,
      margin: { top: "0mm", right: "0mm", bottom: "0mm", left: "0mm" },
    });

    await browser.close();

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

// Register Handlebars helpers
handlebars.registerHelper("gt", function (a, b) {
  return a > b;
});

handlebars.registerHelper("multiply", function (a, b) {
  return a * b;
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
