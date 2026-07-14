import express from "express";
import path from "path";
import dotenv from "dotenv";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";

dotenv.config();

const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

// Helper for fallback fuzzy matching when GenAI fails or is rate-limited
function findFallbackSemanticMatch(target: string, goldenList: string[]): string | null {
  const clean = (s: string) => s.toLowerCase().trim().replace(/[\s-_()/]+/g, " ");
  const cleanTarget = clean(target);
  if (!cleanTarget) return null;

  // Split into words
  const targetWords = cleanTarget.split(" ").filter(w => w.length > 1);

  let bestMatch: string | null = null;
  let highestScore = 0;

  for (const golden of goldenList) {
    const cleanGolden = clean(golden);
    if (!cleanGolden) continue;

    // Direct substring checks
    if (cleanGolden.includes(cleanTarget) || cleanTarget.includes(cleanGolden)) {
      const score = Math.min(cleanTarget.length, cleanGolden.length) / Math.max(cleanTarget.length, cleanGolden.length);
      if (score > highestScore) {
        highestScore = score + 0.5; // Bonus for substring match
        bestMatch = golden;
      }
    }

    // Word overlap check
    const goldenWords = cleanGolden.split(" ").filter(w => w.length > 1);
    if (targetWords.length > 0 && goldenWords.length > 0) {
      const intersection = targetWords.filter(w => goldenWords.includes(w));
      const overlapScore = intersection.length / Math.max(targetWords.length, goldenWords.length);
      if (overlapScore > highestScore && overlapScore >= 0.3) {
        highestScore = overlapScore;
        bestMatch = golden;
      }
    }

    // Character-level bigram Jaccard similarity as fallback
    const getBigrams = (str: string) => {
      const bigrams = new Set<string>();
      for (let i = 0; i < str.length - 1; i++) {
        bigrams.add(str.substring(i, i + 2));
      }
      return bigrams;
    };

    const targetBigrams = getBigrams(cleanTarget);
    const goldenBigrams = getBigrams(cleanGolden);
    if (targetBigrams.size > 0 && goldenBigrams.size > 0) {
      const unionSize = new Set([...targetBigrams, ...goldenBigrams]).size;
      const intersectionSize = [...targetBigrams].filter(b => goldenBigrams.has(b)).length;
      const jaccard = intersectionSize / unionSize;
      if (jaccard > highestScore && jaccard > 0.25) {
        highestScore = jaccard;
        bestMatch = golden;
      }
    }
  }

  // Only return if we have a reasonably confident match
  return highestScore >= 0.25 ? bestMatch : null;
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: "50mb" }));

  // API endpoint for AI Column Mapping
  app.post("/api/match-columns", async (req, res) => {
    try {
      const { goldenColumns, uploadedColumns } = req.body as {
        goldenColumns: string[];
        uploadedColumns: string[];
      };

      if (!goldenColumns || !Array.isArray(goldenColumns) || !uploadedColumns || !Array.isArray(uploadedColumns)) {
        res.status(400).json({ error: "Invalid request payload. Must provide goldenColumns and uploadedColumns arrays." });
        return;
      }

      // 1. Direct exact matching
      const normalize = (s: string) => s.trim().toLowerCase().replace(/[\s-_()/]+/g, "");

      const mappings: Array<{
        target: string;
        golden: string | null;
        matchType: "exact" | "ai" | "unmatched";
      }> = [];

      const unmatchedTargets: string[] = [];

      for (const target of uploadedColumns) {
        const normTarget = normalize(target);
        // Find if there's an exact/near-exact match
        const exactMatch = goldenColumns.find(g => normalize(g) === normTarget);

        if (exactMatch) {
          mappings.push({
            target,
            golden: exactMatch,
            matchType: "exact"
          });
        } else {
          unmatchedTargets.push(target);
        }
      }

      // 2. AI matching for those that don't have an exact match
      const hasValidApiKey = process.env.GEMINI_API_KEY && 
                             process.env.GEMINI_API_KEY.trim() !== "" && 
                             process.env.GEMINI_API_KEY !== "undefined" && 
                             process.env.GEMINI_API_KEY !== "null";

      console.log(`[AI Match] Total uploaded: ${uploadedColumns.length}, exact matched: ${mappings.length}, unmatched remaining: ${unmatchedTargets.length}`);
      console.log(`[AI Match] Valid API Key present: ${hasValidApiKey}`);

      if (unmatchedTargets.length > 0 && hasValidApiKey) {
        try {
          const prompt = `
You are an Excel spreadsheet column-matching assistant.
We have a master list of "golden" column headers (desired template schema):
${JSON.stringify(goldenColumns)}

We have some uploaded columns that do not have direct exact matches. Find the single closest semantic match for each from our golden list:
${JSON.stringify(unmatchedTargets)}

Rules:
1. Align the uploaded target headers to the golden list of headers.
2. Only match if there is a strong semantic similarity (e.g., "Regulatory Approvals" maps to "Regulatory", "Warranty period" maps to "Warranty", "Sync in" maps to "Sync Input", "Stand" maps to "Foot", "Speaker" maps to "Built-in Speakers").
3. If an uploaded column has absolutely no relation or representation in the golden columns list, map it to empty string "".
          `.trim();

          // Race the Gemini API call against a 5-second timeout to prevent hanging requests
          const timeoutPromise = new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error("Gemini API call timed out (5s)")), 5000)
          );

          // Retry logic up to 2 times
          let response;
          let retries = 2;
          let delay = 1000;
          for (let attempt = 1; attempt <= retries; attempt++) {
            try {
              console.log(`[AI Match] Calling Gemini API (attempt ${attempt}/${retries})...`);
              const apiCall = ai.models.generateContent({
                model: "gemini-3.5-flash",
                contents: prompt,
                config: {
                  responseMimeType: "application/json",
                  responseSchema: {
                    type: Type.OBJECT,
                    properties: {
                      mappings: {
                        type: Type.ARRAY,
                        description: "List of matched column pairs",
                        items: {
                          type: Type.OBJECT,
                          properties: {
                            target: { type: Type.STRING, description: "The original unmatched target header" },
                            golden: { type: Type.STRING, description: "The semantically matched golden header from the goldenColumns list, or empty string if no reasonable semantic match exists" }
                          },
                          required: ["target", "golden"]
                        }
                      }
                    },
                    required: ["mappings"]
                  }
                }
              });

              response = await Promise.race([apiCall, timeoutPromise]);
              console.log("[AI Match] Gemini API response received successfully!");
              break; // Success, exit retry loop
            } catch (err: any) {
              console.warn(`[AI Match] Gemini API call failed (attempt ${attempt}/${retries}):`, err.message || err);
              if (attempt === retries) {
                throw err; // Re-throw the error on final failure to trigger fallback
              }
              await new Promise(resolve => setTimeout(resolve, delay));
              delay *= 1.5;
            }
          }

          const resText = response?.text || "{}";
          const aiResult = JSON.parse(resText) as {
            mappings?: Array<{ target: string; golden: string }>;
          };

          if (aiResult.mappings && Array.isArray(aiResult.mappings)) {
            for (const item of aiResult.mappings) {
              const matchedGolden = item.golden && item.golden.trim() !== "" && goldenColumns.includes(item.golden) ? item.golden : null;
              mappings.push({
                target: item.target,
                golden: matchedGolden,
                matchType: matchedGolden ? "ai" : "unmatched"
              });
            }
          }
        } catch (aiError) {
          console.error("[AI Match] Gemini API matching error, falling back to local fuzzy match:", aiError);
        }
      }

      // Ensure every single unmatched target is mapped (fallback to local fuzzy matching if AI failed or didn't return some items)
      const mappedTargets = new Set(mappings.map(m => m.target));
      const remainingTargets = unmatchedTargets.filter(t => !mappedTargets.has(t));

      if (remainingTargets.length > 0) {
        console.log(`[AI Match] Running fallback matching for ${remainingTargets.length} columns.`);
        for (const target of remainingTargets) {
          const fallbackGolden = findFallbackSemanticMatch(target, goldenColumns);
          mappings.push({
            target,
            golden: fallbackGolden,
            matchType: fallbackGolden ? "ai" : "unmatched"
          });
        }
      }

      res.json({ mappings });
    } catch (error) {
      console.error("Match columns api error:", error);
      res.status(500).json({ error: "Internal Server Error during column matching." });
    }
  });

  // Serve static files / Vite middleware
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
