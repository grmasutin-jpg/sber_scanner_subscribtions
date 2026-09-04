import 'dotenv/config';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const app = express();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = Number(process.env.PORT || 3000);
const MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-v4-pro';
const REASONING_EFFORT = process.env.DEEPSEEK_REASONING_EFFORT || 'max';

app.use(express.json({ limit: '16mb' }));
app.use(express.static(__dirname));

const allowedCategories = new Set([
  'Кино и видео', 'Музыка', 'Книги', 'Облако', 'Софт', 'Игры',
  'Связь', 'Фитнес', 'Образование', 'Доставка', 'Финансы',
  'Рестораны и кафе', 'Продукты', 'Товары и покупки', 'Услуги',
  'Транспорт', 'Такси', 'Здоровье', 'ЖКХ', 'Другое'
]);

const roundMoney = value => Math.round((Number(value) || 0) * 100) / 100;
const clamp = (value, min, max) => Math.max(min, Math.min(max, Number(value) || 0));

function sanitizeText(value, max = 300) {
  return String(value || '').replace(/\s+/g, ' ').trim().slice(0, max);
}

function sanitizeCandidate(candidate) {
  return {
    candidate_id: sanitizeText(candidate.candidate_id, 80),
    suggested_name: sanitizeText(candidate.suggested_name, 120),
    merchant_key: sanitizeText(candidate.merchant_key, 180),
    cadence: sanitizeText(candidate.cadence, 40),
    confidence: clamp(candidate.confidence, 0, 1),
    typical_amount: roundMoney(candidate.typical_amount),
    observed_total: roundMoney(candidate.observed_total),
    first_date: sanitizeText(candidate.first_date, 20),
    last_date: sanitizeText(candidate.last_date, 20),
    local_category: sanitizeText(candidate.local_category || 'Другое', 80),
    local_subscription_hint: candidate.local_subscription_hint === true,
    local_veto: candidate.local_veto === true,
    evidence: Array.isArray(candidate.evidence)
      ? candidate.evidence.slice(0, 18).map((e) => ({
          transaction_id: sanitizeText(e.transaction_id, 80),
          date: sanitizeText(e.date, 20),
          amount: roundMoney(e.amount),
          merchant_name: sanitizeText(e.merchant_name, 130),
          description: sanitizeText(e.description, 320)
        }))
      : []
  };
}

function sanitizeTransaction(transaction) {
  return {
    transaction_id: sanitizeText(transaction.transaction_id, 90),
    date: sanitizeText(transaction.date, 20),
    amount: roundMoney(transaction.amount),
    direction: ['income', 'expense'].includes(transaction.direction) ? transaction.direction : 'unknown',
    category: allowedCategories.has(transaction.category) ? transaction.category : 'Другое',
    merchant_name: sanitizeText(transaction.merchant_name, 130),
    description: sanitizeText(transaction.description, 340),
    raw_source: sanitizeText(transaction.raw_source, 750)
  };
}

async function deepSeekJson(system, user, maxTokens = 8000) {
  const response = await fetch('https://api.deepseek.com/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${process.env.DEEPSEEK_API_KEY}`
    },
    body: JSON.stringify({
      model: MODEL,
      thinking: { type: 'enabled' },
      reasoning_effort: REASONING_EFFORT,
      max_tokens: maxTokens,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user }
      ]
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`DeepSeek ${response.status}: ${body.slice(0, 500)}`);
  }

  const payload = await response.json();
  const content = payload?.choices?.[0]?.message?.content || '{}';
  let parsed;
  try { parsed = JSON.parse(content); } catch { parsed = {}; }
  return { parsed, model: payload?.model || MODEL };
}

function rawContainsDate(raw, isoDate) {
  const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const [, y, mo, d] = m;
  const variants = [
    `${y}-${mo}-${d}`, `${y}.${mo}.${d}`, `${y}/${mo}/${d}`,
    `${d}.${mo}.${y}`, `${d}/${mo}/${y}`, `${d}-${mo}-${y}`,
    `${d}.${mo}.${y.slice(-2)}`, `${d}/${mo}/${y.slice(-2)}`, `${d}-${mo}-${y.slice(-2)}`
  ];
  return variants.some(v => String(raw).includes(v));
}

function parseLooseMoney(value) {
  let s = String(value || '').replace(/\u00a0/g, ' ').replace(/[₽$€£]/g, '').replace(/\b(RUB|RUR|руб\.?)/gi, '').trim().replace(/−/g, '-');
  const neg = /^-|\(.*\)/.test(s);
  s = s.replace(/[()]/g, '').replace(/\s/g, '');
  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.includes(',')) s = s.replace(',', '.');
  else if (/^\d{1,3}(?:\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = parseFloat(s.replace(/[^\d.+-]/g, ''));
  return Number.isFinite(n) ? (neg ? -Math.abs(n) : n) : NaN;
}

function rawContainsAmount(raw, target) {
  const text = String(raw || '');
  const tokens = [...text.matchAll(/[+\-−]?\s*\d{1,3}(?:[ \u00a0]\d{3})*(?:[.,]\d{1,2})|[+\-−]?\s*\d{1,9}(?:[.,]\d{1,2})?/g)].map(m => m[0]);
  const wanted = Math.abs(Number(target));
  return tokens.some(token => {
    const n = parseLooseMoney(token);
    return Number.isFinite(n) && Math.abs(Math.abs(n) - wanted) < 0.005;
  });
}

function rawSupportsDirection(raw, direction) {
  const text = String(raw || '').toLowerCase();
  if (direction === 'income') return /зачислен|приход|кредит|поступлен|пополн|зарплат|refund|возврат|cashback|кэшбэк|перевод\s+от|сбп\s+от/.test(text);
  if (direction === 'expense') return /списан|расход|дебет|оплата|покупк|purchase|payment|снятие|перевод\s+(?:на|для|кому)|сбп\s+(?:на|для)/.test(text) || /(?:^|\s)[−-]\s*\d/.test(text);
  return false;
}

function merchantSupported(raw, value) {
  const normalize = s => String(s || '').toLowerCase().replace(/[^a-zа-яё0-9]+/gi, ' ').replace(/\s+/g, ' ').trim();
  const r = normalize(raw), v = normalize(value);
  if (v.length < 2) return false;
  const tokens = v.split(' ').filter(x => x.length >= 3);
  return tokens.length ? tokens.some(token => r.includes(token)) : r.includes(v);
}

function markSafeAuditIssue(issue, source) {
  const field = String(issue?.field || '');
  const confidence = clamp(issue?.confidence, 0, 1);
  const suggested = issue?.suggested_value;
  let safe = false;
  if (confidence >= 0.97) {
    if (field === 'date') safe = rawContainsDate(source.raw_source, suggested);
    else if (field === 'amount') safe = Number(suggested) > 0 && rawContainsAmount(source.raw_source, suggested);
    else if (field === 'direction') safe = ['income', 'expense'].includes(suggested) && rawSupportsDirection(source.raw_source, suggested);
    else if (field === 'merchant_name') safe = merchantSupported(source.raw_source, suggested);
  }
  return {
    transaction_id: source.transaction_id,
    field: ['date', 'amount', 'merchant_name', 'description', 'direction', 'category'].includes(field) ? field : 'description',
    suggested_value: typeof suggested === 'number' ? roundMoney(suggested) : sanitizeText(suggested, 160),
    confidence,
    reason: sanitizeText(issue?.reason, 300),
    safe_to_apply: safe
  };
}

app.get('/api/health', (_req, res) => {
  res.json({
    ok: true,
    deepseekConfigured: Boolean(process.env.DEEPSEEK_API_KEY),
    model: MODEL,
    thinking: true,
    reasoningEffort: REASONING_EFFORT
  });
});

app.post('/api/deepseek/audit', async (req, res) => {
  try {
    const raw = Array.isArray(req.body?.transactions) ? req.body.transactions : [];
    const transactions = raw.slice(0, 260).map(sanitizeTransaction).filter(t => t.transaction_id && t.raw_source);
    if (!transactions.length) return res.json({ mode: 'deterministic', checked_count: 0, coverage: 0, issues: [], message: 'Нет строк для AI-аудита.' });
    if (!process.env.DEEPSEEK_API_KEY) return res.json({ mode: 'deterministic', checked_count: 0, coverage: 0, issues: [], message: 'DeepSeek не настроен.' });

    const system = `Ты — аудитор банковской выписки. Работай предельно строго и с thinking/reasoning. Тебе передают уже распознанные транзакции вместе с raw_source — исходной строкой из файла.\n\nОБЯЗАТЕЛЬНО для КАЖДОЙ транзакции сравни с raw_source четыре поля: date, amount, merchant_name и description. Дополнительно проверь direction, если исходная строка явно содержит признак списания/зачисления.\n\nПРАВИЛА:\n1. Ничего не придумывай. Единственный источник истины — raw_source.\n2. checked_ids должен содержать КАЖДЫЙ transaction_id из входа ровно один раз, даже если строка полностью корректна.\n3. В issues добавляй только реальную проблему, которую можно объяснить содержимым raw_source.\n4. suggested_value для date/amount давай только когда правильное значение явно присутствует в raw_source.\n5. Не исправляй обычное описание стилистически. description — проблема только если распознанный текст потерял важную часть или явно не соответствует raw_source.\n6. merchant_name можно нормализовать только из слов raw_source.\n7. Не считай повторяющиеся одинаковые операции дублями автоматически.\n8. Выведи только JSON.\n\nФормат: {"checked_ids":["..."],"issues":[{"transaction_id":"...","field":"date|amount|merchant_name|description|direction|category","suggested_value":"...","confidence":0.0,"reason":"кратко"}],"batch_summary":"кратко"}.`;

    const { parsed, model } = await deepSeekJson(system, `Проверь все строки пакета ${Number(req.body?.batch_index || 1)} из ${Number(req.body?.total_batches || 1)}. JSON:\n${JSON.stringify({ summary: req.body?.summary || {}, transactions })}`, 9000);
    const validIds = new Set(transactions.map(t => t.transaction_id));
    const checkedIds = [...new Set((Array.isArray(parsed.checked_ids) ? parsed.checked_ids : []).map(String).filter(id => validIds.has(id)))];
    const byId = new Map(transactions.map(t => [t.transaction_id, t]));
    const issues = [];
    for (const issue of Array.isArray(parsed.issues) ? parsed.issues : []) {
      const source = byId.get(String(issue?.transaction_id || ''));
      if (!source) continue;
      issues.push(markSafeAuditIssue(issue, source));
    }
    const coverage = transactions.length ? checkedIds.length / transactions.length : 0;
    res.json({ mode: 'deepseek', model, checked_count: checkedIds.length, coverage, issues: issues.slice(0, 160), batch_summary: sanitizeText(parsed.batch_summary, 500) });
  } catch (error) {
    console.error(error);
    res.status(502).json({ error: 'AI statement audit failed', details: sanitizeText(error.message || error, 500) });
  }
});

app.post('/api/deepseek/validate', async (req, res) => {
  try {
    const raw = Array.isArray(req.body?.candidates) ? req.body.candidates : [];
    const candidates = raw
      .slice(0, 100)
      .map(sanitizeCandidate)
      .filter(c => c.candidate_id && c.evidence.length >= 2 && !c.local_veto);

    if (!candidates.length) {
      return res.json({ mode: 'deterministic', subscriptions: [], message: 'Нет безопасных кандидатов для AI-проверки.' });
    }

    if (!process.env.DEEPSEEK_API_KEY) {
      return res.json({
        mode: 'deterministic',
        subscriptions: candidates
          .filter(c => c.local_subscription_hint)
          .map(c => ({
            candidate_id: c.candidate_id,
            is_subscription: true,
            name: c.suggested_name || c.merchant_key || 'Регулярный платёж',
            category: allowedCategories.has(c.local_category) ? c.local_category : 'Другое',
            confidence_adjustment: -0.04,
            reason: 'DeepSeek не настроен. Подписка показана только потому, что локальный классификатор нашёл признаки регулярного сервиса и не обнаружил признаков обычной покупки.'
          })),
        message: 'DeepSeek не настроен: неизвестные повторяющиеся покупки не считаются подписками автоматически.'
      });
    }

    const system = `Ты — строгий финансовый аудитор подписок. Используй thinking/reasoning максимально тщательно. На входе только кандидаты, уже найденные детерминированным алгоритмом в банковской выписке.\n\nДЛЯ КАЖДОГО кандидата обязательно проверь:\n- все даты evidence и интервалы между ними;\n- каждую сумму evidence, typical_amount и observed_total;\n- merchant_name и suggested_name;\n- каждое description evidence;\n- соответствие cadence фактическим датам;\n- не является ли повторяемость обычными покупками.\n\nКРИТИЧЕСКИЕ ПРАВИЛА:\n1. НИКОГДА не придумывай сервисы, транзакции, суммы или даты.\n2. Возвращай только candidate_id, присутствующие во входном JSON.\n3. Нельзя добавлять новую подписку, если её нет среди кандидатов.\n4. is_subscription=true только когда совокупность dates + amounts + merchant_name + description действительно указывает на подписку, автопродление или регулярный счёт за сервис.\n5. Переводы между людьми, зарплата, возвраты, пополнения, рестораны/кафе, супермаркеты, маркетплейсы, обычные покупки, такси, разовые услуги, налоги и снятие наличных — НЕ подписки, даже при повторении.\n6. Если хотя бы одно ключевое поле противоречит гипотезе подписки — понизь уверенность; при сомнении ставь false.\n7. Название можно нормализовать только если оно очевидно следует из merchant_name/description.\n8. category только из разрешённого списка.\n9. Выведи только JSON.\n\nФормат: {"results":[{"candidate_id":"...","is_subscription":true,"name":"...","category":"Другое","confidence_adjustment":0.0,"reason":"почему, с опорой на даты/суммы/описания"}]}. Разрешённые category: ${[...allowedCategories].join(', ')}.`;

    const { parsed, model } = await deepSeekJson(system, `Перепроверь всех кандидатов. JSON:\n${JSON.stringify({ candidates })}`, 14000);
    const byId = new Map(candidates.map(c => [c.candidate_id, c]));
    const seen = new Set();
    const subscriptions = [];

    for (const item of Array.isArray(parsed.results) ? parsed.results : []) {
      const id = String(item?.candidate_id || '');
      const source = byId.get(id);
      if (!source || seen.has(id) || source.local_veto) continue;
      seen.add(id);
      const category = allowedCategories.has(item.category)
        ? item.category
        : (allowedCategories.has(source.local_category) ? source.local_category : 'Другое');
      const safeName = source.suggested_name || source.merchant_key || 'Регулярный платёж';
      subscriptions.push({
        candidate_id: id,
        is_subscription: item.is_subscription === true,
        name: safeName,
        category,
        confidence_adjustment: clamp(item.confidence_adjustment, -0.3, 0.12),
        reason: sanitizeText(item.reason, 420),
        evidence_count: source.evidence.length,
        source_guard: true
      });
    }

    res.json({ mode: 'deepseek', subscriptions, model });
  } catch (error) {
    console.error(error);
    res.status(502).json({ error: 'AI validation failed', details: sanitizeText(error.message || error, 500) });
  }
});

app.listen(PORT, () => {
  console.log(`Sber Finance Scanner: http://localhost:${PORT}`);
});
