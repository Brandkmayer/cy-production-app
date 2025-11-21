import React, { useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

// Constants
const BAG_NUMBERS = [1, 3, 5];
const LBS_PER_ACRE_FACTOR = 55.7612;

// ---------- Helper functions ----------

// Normalize the Date column into a string like "2025-09-30"
function normalizeDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    const yyyy = value.getFullYear();
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    const dd = String(value.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }
  if (typeof value === "string") {
    // You can tighten this up if needed
    return value;
  }
  return "";
}

// Split Ancestry: "Turkey Creek Allotment > Turkey Creek Pasture"
// -> ALLOTMENT = "Turkey Creek", PASTURE = "Turkey Creek"
function parseAncestry(ancestry) {
  if (!ancestry) return { allotment: "", pasture: "" };

  const parts = String(ancestry).split(">");
  let allotmentRaw = parts[0] ? parts[0].trim() : "";
  let pastureRaw = parts[1] ? parts[1].trim() : "";

  // Drop trailing "Allotment" / "Pasture" words
  allotmentRaw = allotmentRaw.replace(/\s+Allotment$/i, "").trim();
  pastureRaw = pastureRaw.replace(/\s+Pasture$/i, "").trim();

  return { allotment: allotmentRaw, pasture: pastureRaw };
}

// Extract KA from SiteID: keep everything from first letter on
// e.g. "03-01-01-00112-001-C3" -> "C3"
function extractKA(siteId) {
  if (!siteId) return "";
  const m = String(siteId).match(/[A-Za-z].*$/);
  return m ? m[0] : "";
}

function asNumber(x) {
  const n = Number(x);
  return Number.isFinite(n) ? n : NaN;
}

function round(value, digits = 2) {
  const n = asNumber(value);
  if (!Number.isFinite(n)) return "";
  const factor = Math.pow(10, digits);
  return Math.round(n * factor) / factor;
}

// Parse an Excel file in the browser using XLSX
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// Build the production template rows (3 rows per site / KA)
function makeProductionTemplateRows(cyRowsProcessed) {
  const templateRows = [];
  const seen = new Set();

  cyRowsProcessed.forEach((row) => {
    const DATE = row.DATE;
    const ALLOTMENT = row.ALLOTMENT;
    const PASTURE = row.PASTURE;
    const KA = row.KA;

    if (!DATE || !KA) return;

    const key = `${DATE}||${ALLOTMENT}||${PASTURE}||${KA}`;
    if (seen.has(key)) return;
    seen.add(key);

    BAG_NUMBERS.forEach((bag) => {
      templateRows.push({
        DATE,
        ALLOTMENT,
        PASTURE,
        KA,
        "BAG #": bag,
        "GW (g)": "",
        "Dry WT. (g)": "",
        "(-BAG)": "",
        "NET WT.": "",
      });
    });
  });

  // Optional: sort nicely
  templateRows.sort((a, b) => {
    if (a.ALLOTMENT !== b.ALLOTMENT)
      return a.ALLOTMENT.localeCompare(b.ALLOTMENT);
    if (a.PASTURE !== b.PASTURE)
      return a.PASTURE.localeCompare(b.PASTURE);
    if (a.DATE !== b.DATE) return String(a.DATE).localeCompare(String(b.DATE));
    return String(a.KA).localeCompare(String(b.KA));
  });

  return templateRows;
}

// Compute slopes (NET WT. vs BAG #, intercept = 0) by KA
function computeSlopesByKA(prodRows) {
  const perKA = {};

  prodRows.forEach((row) => {
    const KA = row.KA || row["KA"];
    const x = asNumber(row["BAG #"]);
    const y = asNumber(row["NET WT."]);

    if (!KA || !Number.isFinite(x) || !Number.isFinite(y)) return;

    if (!perKA[KA]) {
      perKA[KA] = { sumXY: 0, sumX2: 0, count: 0 };
    }
    perKA[KA].sumXY += x * y;
    perKA[KA].sumX2 += x * x;
    perKA[KA].count += 1;
  });

  const slopes = {};
  Object.entries(perKA).forEach(([ka, stats]) => {
    if (stats.sumX2 > 0) {
      slopes[ka] = stats.sumXY / stats.sumX2;
    }
  });

  return slopes;
}

// Combine CY nValues and KA slopes into Production (lbs/acre)
function computeProductionFromCY(cyRowsProcessed, prodRows) {
  const slopes = computeSlopesByKA(prodRows);
  const missingKAs = new Set();

  // Aggregate nValue by DATE, ALLOTMENT, PASTURE, KA
  const groups = {};

  cyRowsProcessed.forEach((row) => {
    const DATE = row.DATE;
    const ALLOTMENT = row.ALLOTMENT;
    const PASTURE = row.PASTURE;
    const KA = row.KA;

    if (!KA || !DATE) return;

    const nValue = asNumber(row.nValue || row["nValue"]);
    if (!Number.isFinite(nValue)) return;

    const key = `${DATE}||${ALLOTMENT}||${PASTURE}||${KA}`;
    if (!groups[key]) {
      groups[key] = {
        DATE,
        ALLOTMENT,
        PASTURE,
        KA,
        sumN: 0,
        count: 0,
      };
    }
    groups[key].sumN += nValue;
    groups[key].count += 1;
  });

  const outRows = [];

  Object.values(groups).forEach((g) => {
    if (!g.count) return;
    const avgN = g.sumN / g.count;
    const slope = slopes[g.KA];

    if (!Number.isFinite(slope)) {
      missingKAs.add(g.KA);
      return;
    }

    const productionLbsAcre = slope * avgN * LBS_PER_ACRE_FACTOR;

    outRows.push({
      DATE: g.DATE,
      ALLOTMENT: g.ALLOTMENT,
      PASTURE: g.PASTURE,
      KA: g.KA,
      "avg nValue": round(avgN, 3),
      "slope_g_per_bag": round(slope, 4),
      "Production (lbs/acre)": round(productionLbsAcre, 2),
    });
  });

  // Optional sorting
  outRows.sort((a, b) => {
    if (a.ALLOTMENT !== b.ALLOTMENT)
      return a.ALLOTMENT.localeCompare(b.ALLOTMENT);
    if (a.PASTURE !== b.PASTURE)
      return a.PASTURE.localeCompare(b.PASTURE);
    if (a.DATE !== b.DATE) return String(a.DATE).localeCompare(String(b.DATE));
    return String(a.KA).localeCompare(String(b.KA));
  });

  return { rows: outRows, missingKAs: Array.from(missingKAs), slopes };
}

// Export rows to an XLSX file with a single sheet
function downloadXlsxFromJson(rows, filename, sheetName = "Sheet1") {
  if (!rows || !rows.length) {
    alert("Nothing to export.");
    return;
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// ---------- React component ----------

function App() {
  const [cyRows, setCyRows] = useState([]);      // processed Comparative Yield rows
  const [prodRows, setProdRows] = useState([]);  // production calibration rows
  const [status, setStatus] = useState("");

  // Upload RAW files and extract Comparative Yield
  const handleUploadRaw = async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    setStatus("Reading RAW export file(s)...");
    try {
      const allCY = [...cyRows]; // append to existing

      for (const file of files) {
        const wb = await readExcelFile(file);

        const cySheet = wb.Sheets["Comparative Yield"];
        if (!cySheet) {
          console.warn(`File ${file.name} has no 'Comparative Yield' sheet; skipping.`);
          continue;
        }

        const cyData = XLSX.utils.sheet_to_json(cySheet, { defval: null });

        cyData.forEach((row) => {
          const dateRaw = row.Date || row["Date"];
          const DATE = normalizeDate(dateRaw);
          const { allotment, pasture } = parseAncestry(row.Ancestry || row["Ancestry"]);
          const KA = extractKA(row.SiteID || row["SiteID"]);

          allCY.push({
            ...row,
            DATE,
            ALLOTMENT: allotment,
            PASTURE: pasture,
            KA,
          });
        });
      }

      setCyRows(allCY);
      setStatus(`Loaded ${allCY.length} Comparative Yield rows.`);
    } catch (err) {
      console.error(err);
      setStatus("Error reading RAW files. Check console for details.");
    } finally {
      // Clear file input so you can re-upload same files if needed
      event.target.value = "";
    }
  };

  // Upload filled production Excel files (like 20250930 TurkeyCreekProduction.xlsx)
  const handleUploadProduction = async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    setStatus("Reading production (Bag / NET WT.) file(s)...");
    try {
      const allProd = [...prodRows];

      for (const file of files) {
        const wb = await readExcelFile(file);

        // Use the first sheet in the workbook
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

        rows.forEach((row) => {
          // Normalize KA + DATE column names (in case of subtle differences)
          allProd.push({
            ...row,
            KA: row.KA || row["KA"],
            DATE: normalizeDate(row.DATE || row["DATE"]),
          });
        });
      }

      setProdRows(allProd);
      setStatus(`Loaded ${allProd.length} production rows.`);
    } catch (err) {
      console.error(err);
      setStatus("Error reading production files. Check console for details.");
    } finally {
      event.target.value = "";
    }
  };

  // Export the template with 3 rows per KA/site
  const handleExportTemplate = () => {
    if (!cyRows.length) {
      setStatus("No Comparative Yield data loaded.");
      return;
    }
    const templateRows = makeProductionTemplateRows(cyRows);
    downloadXlsxFromJson(
      templateRows,
      "CY_ProductionTemplate.xlsx",
      "ProductionTemplate"
    );
    setStatus(`Exported template with ${templateRows.length} rows.`);
  };

  // Export the Production (lbs/acre) file
  const handleExportProduction = () => {
    if (!cyRows.length) {
      setStatus("No Comparative Yield data loaded.");
      return;
    }
    if (!prodRows.length) {
      setStatus("No production (Bag / NET WT.) data loaded.");
      return;
    }

    const { rows, missingKAs } = computeProductionFromCY(cyRows, prodRows);

    if (!rows.length) {
      setStatus("No Production rows could be computed. Check that KAs match between CY and production files.");
      return;
    }

    downloadXlsxFromJson(
      rows,
      "CY_Production_lbs_per_acre.xlsx",
      "Production_lbs_acre"
    );

    if (missingKAs.length) {
      setStatus(
        `Exported ${rows.length} rows. Note: no calibration found for KAs: ${missingKAs.join(
          ", "
        )}`
      );
    } else {
      setStatus(`Exported ${rows.length} rows (Production lbs/acre).`);
    }
  };

  // Count distinct KAs in each dataset (for a quick sanity check)
  const cyKAs = new Set(cyRows.map((r) => r.KA).filter(Boolean));
  const prodKAs = new Set(prodRows.map((r) => r.KA).filter(Boolean));

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", padding: "1.5rem", maxWidth: 900, margin: "0 auto" }}>
      <h1>Comparative Yield → Production App</h1>

      <p style={{ marginTop: "0.5rem" }}>
        1) Upload RAW Excel exports (C3VGS-Export_RAW...) to build a combined{" "}
        <strong>Comparative Yield</strong> dataset and export a 3-row-per-site
        production template.
        <br />
        2) Fill that template in the field (Bag #, NET WT.) and upload it here.
        <br />
        3) Export a <strong>Production (lbs/acre)</strong> table per KA.
      </p>

      <hr style={{ margin: "1rem 0" }} />

      <section style={{ marginBottom: "1.5rem" }}>
        <h2>Step 1: Upload RAW export(s)</h2>
        <input
          type="file"
          multiple
          accept=".xlsx"
          onChange={handleUploadRaw}
        />
        <div style={{ marginTop: "0.5rem", fontSize: "0.9rem" }}>
          Loaded CY rows: <strong>{cyRows.length}</strong>{" "}
          | Distinct KAs: <strong>{cyKAs.size}</strong>
        </div>
        <button
          style={{ marginTop: "0.75rem" }}
          onClick={handleExportTemplate}
          disabled={!cyRows.length}
        >
          Export 3-row-per-site Production Template
        </button>
      </section>

      <hr style={{ margin: "1rem 0" }} />

      <section style={{ marginBottom: "1.5rem" }}>
        <h2>Step 2: Upload filled Production file(s)</h2>
        <input
          type="file"
          multiple
          accept=".xlsx"
          onChange={handleUploadProduction}
        />
        <div style={{ marginTop: "0.5rem", fontSize: "0.9rem" }}>
          Loaded production rows: <strong>{prodRows.length}</strong>{" "}
          | Distinct KAs: <strong>{prodKAs.size}</strong>
        </div>

        <button
          style={{ marginTop: "0.75rem" }}
          onClick={handleExportProduction}
          disabled={!cyRows.length || !prodRows.length}
        >
          Export Production (lbs/acre) from CY + Bag NET WT.
        </button>
      </section>

      <hr style={{ margin: "1rem 0" }} />

      <div style={{ fontSize: "0.9rem", color: "#444" }}>
        <strong>Status:</strong> {status || "Idle"}
      </div>
    </div>
  );
}

export default App;
