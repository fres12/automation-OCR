require("dotenv").config();
const { GoogleGenerativeAI } = require("@google/generative-ai");

const apiKey = (process.env.GEMINI_API_KEY || "").trim();
const modelName = process.env.GEMINI_MODEL || "gemini-3.6-flash";

console.log("============================================================");
console.log("          TES RESPON CEPAT MODEL DENGAN THINKING MINIMAL    ");
console.log("============================================================");
console.log(`API Key : ${apiKey.substring(0, 8)}...`);
console.log(`Model   : ${modelName}`);

const genAI = new GoogleGenerativeAI(apiKey);
const model = genAI.getGenerativeModel({ model: modelName });

async function testPrompt() {
  process.stdout.write(`\nMengirim prompt uji coba ke ${modelName}... `);
  const start = Date.now();
  try {
    const result = await model.generateContent({
      contents: [
        {
          role: "user",
          parts: [{ text: "Jawab singkat JSON saja: {\"status\": \"ok\", \"pesan\": \"siap OCR\"}" }],
        },
      ],
      generationConfig: {
        responseMimeType: "application/json",
        thinkingConfig: { thinkingLevel: "minimal" },
      },
    });

    const response = await result.response;
    const text = response.text().trim();
    const duration = ((Date.now() - start) / 1000).toFixed(2);

    console.log(`\n✅ BERHASIL (${duration}s)`);
    console.log("Output Model:", text);
    console.log("\nKonfigurasi thinkingConfig sudah tepat & siap dipakai di ocr.js!");
  } catch (err) {
    const duration = ((Date.now() - start) / 1000).toFixed(2);
    console.log(`\n❌ GAGAL (${duration}s)`);
    console.log("Error:", err.message);
  }
}

testPrompt();
