import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import * as XLSX from "xlsx";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Body parser with 50mb limit for large PDF/Excel files
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// Lazy initialization of Gemini client
let aiClient: GoogleGenAI | null = null;
function getAI(): GoogleGenAI {
  if (!aiClient) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error("GEMINI_API_KEY is not configured.");
    }
    aiClient = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiClient;
}

// Primary model: gemini-3.1-flash-lite provides fast response times and high availability
const PRIMARY_MODEL = "gemini-3.1-flash-lite";
const FALLBACK_MODEL = "gemini-3.8-flash";

// Helper: execute Gemini generateContent with automatic retry and model fallback on temporary high demand / 503
async function generateWithRetry(ai: GoogleGenAI, params: any, maxRetries = 3): Promise<any> {
  let lastError: any;
  const modelsToTry = [
    params.model || PRIMARY_MODEL,
    PRIMARY_MODEL,
    FALLBACK_MODEL,
  ];
  // Deduplicate while preserving order
  const modelQueue = Array.from(new Set(modelsToTry));

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const currentModel = modelQueue[Math.min(attempt, modelQueue.length - 1)];
    try {
      return await ai.models.generateContent({
        ...params,
        model: currentModel,
      });
    } catch (err: any) {
      lastError = err;
      const errMsg = String(err?.message || err);
      const isTemporary =
        errMsg.includes("503") ||
        errMsg.includes("UNAVAILABLE") ||
        errMsg.includes("429") ||
        errMsg.includes("high demand") ||
        errMsg.includes("Resource has been exhausted") ||
        errMsg.includes("overloaded");

      if (isTemporary && attempt < maxRetries) {
        const backoffMs = (attempt + 1) * 750 + Math.floor(Math.random() * 350);
        console.warn(`[Gemini API] Temporary spike on ${currentModel} (attempt ${attempt + 1}/${maxRetries}), retrying with model in ${backoffMs}ms...`);
        await new Promise((r) => setTimeout(r, backoffMs));
        continue;
      }
      throw err;
    }
  }
  throw lastError;
}

// Fallback heuristic statement parser in case AI service is temporarily unavailable
interface FallbackSubscription {
  name: string;
  category: string;
  clusterInfo: string;
  price: number;
  icon: string;
  color: string;
  frequency?: string;
  detectedDates?: string[];
}

function parseStatementFallback(rawText: string, filename: string): {
  subscriptions: FallbackSubscription[];
  statementSummary: {
    bankDetected: string;
    period: string;
    totalTransactionsAnalyzed: number;
    note: string;
  };
} {
  let bankDetected = "Банковская выписка";
  const lowerText = rawText.toLowerCase();
  if (lowerText.includes("сбербанк") || lowerText.includes("sber") || lowerText.includes("сбер")) {
    bankDetected = "СберБанк";
  } else if (lowerText.includes("тинькофф") || lowerText.includes("т-банк") || lowerText.includes("t-bank") || lowerText.includes("tcs")) {
    bankDetected = "Т-Банк";
  } else if (lowerText.includes("втб") || lowerText.includes("vtb")) {
    bankDetected = "ВТБ";
  } else if (lowerText.includes("альфа") || lowerText.includes("alfa")) {
    bankDetected = "Альфа-Банк";
  } else if (lowerText.includes("газпромбанк")) {
    bankDetected = "Газпромбанк";
  }

  const catalog: Array<{
    name: string;
    category: string;
    icon: string;
    color: string;
    patterns: RegExp[];
    defaultPrice: number;
  }> = [
    {
      name: "Яндекс Плюс",
      category: "Кино",
      icon: "fa-film",
      color: "#ffcc00",
      patterns: [/yndx[\s*._-]*plus/i, /yandex[\s*._-]*plus/i, /яндекс[\s*._-]*плюс/i, /кинопоиск/i, /kinopoisk/i],
      defaultPrice: 299,
    },
    {
      name: "Netflix",
      category: "Кино",
      icon: "fa-film",
      color: "#E50914",
      patterns: [/netflix/i, /\bnflx\b/i],
      defaultPrice: 999,
    },
    {
      name: "Иви (ivi)",
      category: "Кино",
      icon: "fa-film",
      color: "#ea1d5d",
      patterns: [/ivi\.ru/i, /\bivi\b/i, /\bиви\b/i],
      defaultPrice: 399,
    },
    {
      name: "Okko",
      category: "Кино",
      icon: "fa-tv",
      color: "#6a0dad",
      patterns: [/\bokko\b/i, /\bокко\b/i],
      defaultPrice: 499,
    },
    {
      name: "Premier",
      category: "Кино",
      icon: "fa-film",
      color: "#00a8e8",
      patterns: [/\bpremier\b/i, /премьер/i],
      defaultPrice: 299,
    },
    {
      name: "Start",
      category: "Кино",
      icon: "fa-film",
      color: "#ff4500",
      patterns: [/start\.ru/i, /онлайн-кинотеатр\s*старт/i],
      defaultPrice: 499,
    },
    {
      name: "VK Музыка",
      category: "Музыка",
      icon: "fa-headphones",
      color: "#0077FF",
      patterns: [/vk[\s*._-]*music/i, /vk[\s*._-]*pay/i, /вк[\s*._-]*музык/i, /boom[\s*._-]*music/i],
      defaultPrice: 169,
    },
    {
      name: "Яндекс Музыка",
      category: "Музыка",
      icon: "fa-music",
      color: "#ffcc00",
      patterns: [/yandex[\s*._-]*music/i, /яндекс[\s*._-]*музык/i],
      defaultPrice: 299,
    },
    {
      name: "СберЗвук (Звук)",
      category: "Музыка",
      icon: "fa-headphones",
      color: "#21a038",
      patterns: [/сберзвук/i, /sberzvuk/i, /\bзвук\b.*подписк/i],
      defaultPrice: 199,
    },
    {
      name: "ЛитРес",
      category: "Книги",
      icon: "fa-book",
      color: "#ff6600",
      patterns: [/litres/i, /литрес/i],
      defaultPrice: 399,
    },
    {
      name: "Букмейт",
      category: "Книги",
      icon: "fa-book-open",
      color: "#ff8c42",
      patterns: [/bookmate/i, /букмейт/i],
      defaultPrice: 399,
    },
    {
      name: "Storytel",
      category: "Книги",
      icon: "fa-book-open",
      color: "#ff8c42",
      patterns: [/storytel/i, /сторител/i],
      defaultPrice: 549,
    },
    {
      name: "Apple Services",
      category: "Другое",
      icon: "fa-apple",
      color: "#555555",
      patterns: [/apple\.com\/bill/i, /itunes/i, /icloud/i],
      defaultPrice: 169,
    },
    {
      name: "Telegram Premium",
      category: "Другое",
      icon: "fa-paper-plane",
      color: "#24A1DE",
      patterns: [/telegram[\s*._-]*premium/i, /t\.me\/premium/i, /telegram\s*fzco/i],
      defaultPrice: 299,
    },
    {
      name: "СберПрайм",
      category: "Другое",
      icon: "fa-crown",
      color: "#21a038",
      patterns: [/sberprime/i, /sbprime/i, /сберпрайм/i, /сбер\s*прайм/i],
      defaultPrice: 399,
    },
    {
      name: "Т-Банк Pro",
      category: "Другое",
      icon: "fa-shield-halved",
      color: "#ffdd2d",
      patterns: [/tinkoff[\s*._-]*pro/i, /t-pro/i, /тинькофф[\s*._-]*про/i],
      defaultPrice: 299,
    },
    {
      name: "Облако Mail.ru",
      category: "Другое",
      icon: "fa-cloud",
      color: "#005ff9",
      patterns: [/cloud\.mail\.ru/i, /облако[\s*._-]*mail/i],
      defaultPrice: 99,
    },
    {
      name: "Плата за уведомления",
      category: "Другое",
      icon: "fa-receipt",
      color: "#21a038",
      patterns: [/плата\s*за\s*уведомлен/i, /мобильный\s*банк/i, /смс-банк/i, /плата\s*за\s*обслуживан/i],
      defaultPrice: 99,
    },
    {
      name: "МТС",
      category: "Другое",
      icon: "fa-mobile-screen",
      color: "#e30611",
      patterns: [/mts[\s*._-]*premium/i, /оплата[\s*._-]*мтс/i, /\bmts\b.*(связь|абонент)/i],
      defaultPrice: 450,
    },
    {
      name: "МегаФон",
      category: "Другое",
      icon: "fa-mobile-screen",
      color: "#00b956",
      patterns: [/megafon/i, /мегафон/i],
      defaultPrice: 500,
    },
    {
      name: "Билайн",
      category: "Другое",
      icon: "fa-mobile-screen",
      color: "#ffc800",
      patterns: [/beeline/i, /билайн/i],
      defaultPrice: 480,
    },
    {
      name: "Tele2 / T2",
      category: "Другое",
      icon: "fa-mobile-screen",
      color: "#1f2229",
      patterns: [/tele2/i, /теле2/i, /\bt2\b/i],
      defaultPrice: 450,
    },
  ];

  const lines = rawText.split(/\r?\n/).filter((l) => l.trim().length > 3);
  const detectedMap = new Map<string, {
    catItem: typeof catalog[0];
    dates: string[];
    amounts: number[];
    rawSnippets: string[];
  }>();

  for (const line of lines) {
    const dateMatch = line.match(/\b(\d{2}[./-]\d{2}[./-]\d{2,4}|\d{4}[./-]\d{2}[./-]\d{2})\b/);
    const dateStr = dateMatch ? dateMatch[0] : "";

    const cleanedLine = line.replace(/(\d)\s+(\d)/g, "$1$2");
    const amountMatches = cleanedLine.match(/-?\b\d+([.,]\d{1,2})?\b/g);
    let amount = 0;
    if (amountMatches && amountMatches.length > 0) {
      for (const m of amountMatches) {
        const num = Math.abs(parseFloat(m.replace(",", ".")));
        if (num >= 50 && num <= 50000 && num !== 2024 && num !== 2025 && num !== 2026) {
          amount = num;
          break;
        }
      }
    }

    for (const item of catalog) {
      const matched = item.patterns.some((p) => p.test(line));
      if (matched) {
        let entry = detectedMap.get(item.name);
        if (!entry) {
          entry = {
            catItem: item,
            dates: [],
            amounts: [],
            rawSnippets: [],
          };
          detectedMap.set(item.name, entry);
        }
        if (dateStr && !entry.dates.includes(dateStr)) {
          entry.dates.push(dateStr);
        }
        if (amount > 0) {
          entry.amounts.push(amount);
        }
        if (entry.rawSnippets.length < 3) {
          entry.rawSnippets.push(line.trim().slice(0, 70));
        }
        break;
      }
    }
  }

  const subscriptions: FallbackSubscription[] = [];
  for (const [name, entry] of detectedMap.entries()) {
    let finalPrice = entry.catItem.defaultPrice;
    if (entry.amounts.length > 0) {
      const freq: Record<number, number> = {};
      entry.amounts.forEach((a) => (freq[a] = (freq[a] || 0) + 1));
      let bestA = entry.amounts[0];
      let maxF = 0;
      for (const [a, f] of Object.entries(freq)) {
        if (f > maxF) {
          maxF = f;
          bestA = Number(a);
        }
      }
      finalPrice = Math.round(bestA);
    }

    subscriptions.push({
      name,
      category: entry.catItem.category,
      clusterInfo: entry.rawSnippets.length > 0
        ? `Сгруппировано из операций: ${entry.rawSnippets.slice(0, 2).join(" | ")}`
        : `Регулярная подписка сервиса ${name}`,
      price: finalPrice,
      icon: entry.catItem.icon,
      color: entry.catItem.color,
      frequency: "Ежемесячно",
      detectedDates: entry.dates.slice(0, 5),
    });
  }

  return {
    subscriptions,
    statementSummary: {
      bankDetected,
      period: "По предоставленной выписке",
      totalTransactionsAnalyzed: Math.max(lines.length, subscriptions.length),
      note: "Анализ выполнен резервным интеллектуальным парсером банковских выписок.",
    },
  };
}

// Health check endpoint
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    hasApiKey: Boolean(process.env.GEMINI_API_KEY),
  });
});

// Helper: decode text with fallback to Windows-1251 (common in Russian bank CSV exports)
function decodeBuffer(buffer: Buffer): string {
  try {
    const utf8Decoder = new TextDecoder("utf-8", { fatal: true });
    return utf8Decoder.decode(buffer);
  } catch {
    try {
      const win1251Decoder = new TextDecoder("windows-1251");
      return win1251Decoder.decode(buffer);
    } catch {
      return buffer.toString("utf-8");
    }
  }
}

// Endpoint: analyze uploaded bank statement
app.post("/api/analyze-statement", async (req, res) => {
  try {
    const { filename = "statement", mimeType = "", base64 = "", text = "" } = req.body;

    if (!base64 && !text) {
      return res.status(400).json({ error: "Файл не передан или пуст" });
    }

    const ai = getAI();
    const ext = path.extname(filename).toLowerCase();

    let geminiContents: any[] = [];
    let fileDescription = `Имя файла: ${filename}`;
    let extractedText = "";

    const promptInstruction = `Внимательно проанализируй эту банковскую выписку (по счёту или карте) и найди ВСЕ РЕАЛЬНЫЕ периодические подписки, регулярные сервисы и повторяющиеся списания, которые есть в данных.

КРИТИЧЕСКИ ВАЖНО:
1. Используй ТОЛЬКО РЕАЛЬНЫЕ транзакции из предоставленной выписки. НЕ придумывай вымышленные подписки и платежи, которых нет в файле!
2. Что искать:
   - Сервисы потокового видео и кино: Иви (ivi), Кинопоиск (Kinopoisk / Yandex Plus), Okko, Premier, Start, Netflix, Amediateka, More.tv, Wink, MEGOGO и др.
   - Музыкальные подписки: Яндекс Музыка, VK Музыка, СберЗвук (Звук), Apple Music, Spotify, YouTube Music.
   - Экосистемные подписки: Яндекс Плюс, СберПрайм / СберПрайм+, Tinkoff Pro / Premium, Газпром Бонус, Ozon Premium, МТС Premium.
   - Книги и аудиокниги: Литрес (Litres), Строки, Букмейт (Bookmate), Storytel.
   - Цифровые сервисы и облака: Telegram Premium, Google One / Drive, Apple iCloud (APPLE.COM/BILL), Облако Mail.ru, Яндекс 360, ChatGPT, Claude, хостинги, домены, VPN.
   - Мобильная связь и интернет (периодические тарифы): МТС, Мегафон, Билайн, Т-Мобайл, Tele2, Ростелеком, Дом.ru.
   - Регулярные банковские комиссии за пакеты или смс: "Плата за уведомления", "Мобильный банк", плата за обслуживание карты/пакета.
   - Фитнес-клубы, регулярные секции, регулярные пожертвования, повторяющиеся списания одной и той же суммы с периодичностью (ежемесячно, ежегодно, еженедельно).
3. Кластеризация и группировка:
   - Объединяй транзакции с разным написанием одного и того же сервиса (например, "YNDX PLUS", "Yandex Plus", "YM*Plus" -> "Яндекс Плюс").
   - В поле clusterInfo обязательно укажи: "Сгруппировано из: [реальные описания операций из выписки, даты и суммы]".
4. Для поля category выбери одно из:
   - "Кино"
   - "Музыка"
   - "Книги"
   - "Другое" (связь, экосистемы, облака, фитнес, плата за уведомления и любые другие сервисы)
5. Поле price:
   - Ежемесячная сумма в рублях (целое число или число с плавающей точкой). Если списание годовое, укажи приблизительный месячный эквивалент или раздели на 12.
6. Поле icon: подходящий класс FontAwesome:
   - "fa-film", "fa-tv", "fa-music", "fa-headphones", "fa-book", "fa-book-open", "fa-crown", "fa-mobile-screen", "fa-cloud", "fa-gamepad", "fa-dumbbell", "fa-receipt", "fa-shield-halved", "fa-bolt", "fa-wifi"
7. Поле color: соответствующий брендовый HEX-цвет (Яндекс: "#ffcc00", Сбер: "#21a038", VK: "#0077FF", Netflix: "#E50914", Telegram: "#24A1DE", МТС: "#e30611", Apple: "#555555", Литрес: "#ff6600", Okko: "#6a0dad", Ivi: "#ea1d5d", Иной: "#21a038" или "#f97316").
8. Если в файле вообще нет регулярных списаний или подписок, верни пустой массив subscriptions: [].`;

    // Process different file types
    if (ext === ".xlsx" || ext === ".xls" || mimeType.includes("spreadsheet") || mimeType.includes("excel")) {
      // Excel file: parse sheets to CSV/text
      const fileBuffer = Buffer.from(base64, "base64");
      const workbook = XLSX.read(fileBuffer, { type: "buffer" });
      let sheetContents = "";
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        sheetContents += `\n--- ЛИСТ: ${sheetName} ---\n` + XLSX.utils.sheet_to_csv(sheet);
      }
      extractedText = sheetContents;
      geminiContents = [
        {
          text: `${promptInstruction}\n\n${fileDescription}\n\nДанные выписки из Excel:\n${sheetContents}`,
        },
      ];
    } else if (ext === ".csv" || ext === ".txt" || mimeType.includes("csv") || mimeType.includes("text")) {
      // CSV or plain text
      let textContent = text;
      if (!textContent && base64) {
        const fileBuffer = Buffer.from(base64, "base64");
        textContent = decodeBuffer(fileBuffer);
      }
      extractedText = textContent;
      geminiContents = [
        {
          text: `${promptInstruction}\n\n${fileDescription}\n\nДанные выписки (CSV/текст):\n${textContent}`,
        },
      ];
    } else if (ext === ".pdf" || mimeType === "application/pdf") {
      // PDF file: pass as inlineData directly to Gemini
      geminiContents = [
        {
          inlineData: {
            mimeType: "application/pdf",
            data: base64,
          },
        },
        {
          text: `${promptInstruction}\n\n${fileDescription}`,
        },
      ];
      // Also try to extract text streams from PDF for potential offline fallback
      try {
        const fileBuffer = Buffer.from(base64, "base64");
        const rawLatin = fileBuffer.toString("latin1");
        extractedText = rawLatin;
      } catch (e) {
        // ignore
      }
    } else if (mimeType.startsWith("image/") || [".png", ".jpg", ".jpeg", ".webp"].includes(ext)) {
      // Image of statement
      const imageMime = mimeType || (ext === ".png" ? "image/png" : "image/jpeg");
      geminiContents = [
        {
          inlineData: {
            mimeType: imageMime,
            data: base64,
          },
        },
        {
          text: `${promptInstruction}\n\n${fileDescription}`,
        },
      ];
    } else {
      // Fallback: try as PDF if possible, otherwise text decode
      const fileBuffer = Buffer.from(base64, "base64");
      const decoded = decodeBuffer(fileBuffer);
      if (decoded && !decoded.includes("\u0000")) {
        extractedText = decoded;
        geminiContents = [
          {
            text: `${promptInstruction}\n\n${fileDescription}\n\nДанные выписки:\n${decoded}`,
          },
        ];
      } else {
        geminiContents = [
          {
            inlineData: {
              mimeType: "application/pdf",
              data: base64,
            },
          },
          {
            text: `${promptInstruction}\n\n${fileDescription}`,
          },
        ];
      }
    }

    let result: any = null;
    let geminiError: any = null;

    try {
      const response = await generateWithRetry(ai, {
        model: PRIMARY_MODEL,
        contents: {
          parts: geminiContents,
        },
        config: {
          systemInstruction:
            "Ты — специализированная AI-система банковской аналитики. Ты анализируешь реальные выписки клиентов и извлекаешь подписки и регулярные платежи. Всегда возвращай валидный JSON строго по схеме.",
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              subscriptions: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING, description: "Название сервиса или подписки" },
                    category: { type: Type.STRING, description: "Категория: Кино, Музыка, Книги или Другое" },
                    clusterInfo: { type: Type.STRING, description: "Реальные данные из выписки: из каких списаний сгруппировано, даты, суммы" },
                    price: { type: Type.NUMBER, description: "Ежемесячная стоимость в рублях" },
                    icon: { type: Type.STRING, description: "Класс иконки FontAwesome (например fa-film, fa-music, fa-book, fa-crown, fa-mobile-screen)" },
                    color: { type: Type.STRING, description: "HEX-код цвета бренда" },
                    frequency: { type: Type.STRING, description: "Периодичность, например 'Ежемесячно'" },
                    detectedDates: {
                      type: Type.ARRAY,
                      items: { type: Type.STRING },
                      description: "Даты найденных списаний",
                    },
                  },
                  required: ["name", "category", "clusterInfo", "price", "icon", "color"],
                },
              },
              statementSummary: {
                type: Type.OBJECT,
                properties: {
                  bankDetected: { type: Type.STRING, description: "Определенный банк (например Сбербанк, Т-Банк, ВТБ, Альфа-Банк или Неизвестен)" },
                  period: { type: Type.STRING, description: "Период выписки" },
                  totalTransactionsAnalyzed: { type: Type.INTEGER, description: "Приблизительное число проверенных операций" },
                  note: { type: Type.STRING, description: "Краткий комментарий AI о выписке" },
                },
              },
            },
            required: ["subscriptions"],
          },
        },
      });

      const rawText = response.text || "{}";
      result = JSON.parse(rawText);
    } catch (err: any) {
      console.warn("Gemini API call failed, evaluating intelligent fallback...", err?.message || err);
      geminiError = err;
    }

    // If Gemini succeeded, return parsed data
    if (result && Array.isArray(result.subscriptions)) {
      return res.json({
        success: true,
        subscriptions: result.subscriptions || [],
        statementSummary: result.statementSummary || {},
      });
    }

    // Fallback parser if text content is present
    if (extractedText && extractedText.trim().length > 10) {
      console.log("Applying backup heuristic statement parser...");
      const fallbackResult = parseStatementFallback(extractedText, filename);
      if (fallbackResult.subscriptions.length > 0 || !geminiError) {
        return res.json({
          success: true,
          subscriptions: fallbackResult.subscriptions,
          statementSummary: fallbackResult.statementSummary,
        });
      }
    }

    // If both Gemini and fallback failed, format a clean friendly error message
    const errString = String(geminiError?.message || "");
    let friendlyError = "Не удалось проанализировать выписку. Проверьте формат файла.";
    if (errString.includes("503") || errString.includes("UNAVAILABLE") || errString.includes("high demand")) {
      friendlyError = "AI-модель сейчас испытывает временный пик нагрузки (код 503). Пожалуйста, повторите попытку через пару секунд или загрузите выписку в Excel/CSV формате.";
    } else if (errString.includes("429") || errString.includes("quota") || errString.includes("Resource has been exhausted")) {
      friendlyError = "Превышен лимит запросов к AI-модели. Пожалуйста, подождите минуту и повторите попытку.";
    }

    res.status(503).json({
      success: false,
      error: friendlyError,
    });
  } catch (error: any) {
    console.error("Analysis error:", error);
    res.status(500).json({
      success: false,
      error: error.message || "Ошибка при анализе файла выписки через Gemini",
    });
  }
});

// Endpoint: AI-generated cancellation letter
app.post("/api/generate-cancellation", async (req, res) => {
  try {
    const { serviceName, monthlyPrice, isBulk = false, services = [] } = req.body;
    const ai = getAI();

    let prompt = "";
    if (isBulk && Array.isArray(services) && services.length > 0) {
      const itemsList = services.map((s: any) => `- ${s.name} (${s.price} ₽/мес)`).join("\n");
      prompt = `Составь юридически грамотное, вежливое и строгое заявление об отказе от следующих платных подписок и отзыве согласия на автосписания (рекуррентные платежи) с банковской карты:\n${itemsList}\nТребование: удалить данные банковской карты, прекратить списания и выслать подтверждение.`;
    } else {
      prompt = `Составь официальное, вежливое и юридически корректное обращение в службу поддержки сервиса "${serviceName}" с требованием отменить платную подписку (${monthlyPrice} ₽/мес), удалить привязанную банковскую карту Сбербанка/Т-Банка из профиля, отвязать автосписание (рекуррентные платежи) и выслать письменное подтверждение расторжения.`;
    }

    const response = await generateWithRetry(ai, {
      model: PRIMARY_MODEL,
      contents: prompt,
      config: {
        systemInstruction: "Ты — юрист по защите прав потребителей. Пиши готовый для отправки текст заявления/письма без лишних вступлений.",
      },
    });

    res.json({
      success: true,
      letter: response.text || "",
    });
  } catch (error: any) {
    console.error("Cancellation generation error:", error);
    const { serviceName, monthlyPrice, isBulk = false, services = [] } = req.body;
    let fallbackLetter = "";
    if (isBulk && Array.isArray(services) && services.length > 0) {
      fallbackLetter = `Здравствуйте!\n\nПрошу отменить платные подписки на следующие сервисы и удалить платежные данные банковской карты:\n\n` +
        services.map((s: any) => `- ${s.name} (${s.price} ₽/мес)`).join("\n") +
        `\n\nНастоящим отзываю согласие на проведение регулярных (рекуррентных) автоплатежей. Прошу направить письменное подтверждение расторжения в ответном письме.\n\nДата: ${new Date().toLocaleDateString("ru-RU")}\nС уважением,\n[Ваше ФИО]`;
    } else {
      fallbackLetter = `В службу клиентской поддержки сервиса «${serviceName || "Сервис"}»\nот: [Ваше ФИО]\nEmail: [Ваш email]\n\nЗАЯВЛЕНИЕ ОБ ОТКАЗЕ ОТ ПОДПИСКИ И ОТЗЫВЕ СОГЛАСИЯ НА АВТОСПИСАНИЯ\n\nНастоящим уведомляю об отказе от использования платной подписки на сервис «${serviceName || "сервис"}» (${monthlyPrice || "..."} ₽/мес) и требую расторгнуть пользовательский договор.\n\nНа основании ст. 32 Закона РФ «О защите прав потребителей» и ст. 782 ГК РФ требую:\n1. Прекратить действие подписки;\n2. Отозвать согласие на автосписания (рекуррентные платежи);\n3. Удалить данные привязанной карты из платежного профиля;\n4. Направить письменное подтверждение расторжения ответным сообщением.\n\nДата: ${new Date().toLocaleDateString("ru-RU")}\nПодпись: ____________ / [Ваше ФИО]`;
    }

    res.json({
      success: true,
      letter: fallbackLetter,
    });
  }
});

// Server setup with Vite middleware
async function startServer() {
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
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
