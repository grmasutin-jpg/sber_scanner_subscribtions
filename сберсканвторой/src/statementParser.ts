/**
 * Client-side Bank Statement Parser for XLSX and CSV files
 * Powered by SheetJS (xlsx) — 100% Client-side, works completely without a backend!
 *
 * Усовершенствованный алгоритм поиска подписок в выписке банка (СберБанк, Т-Банк и др.)
 * с учетом:
 * 1. Пробных периодов (1 ₽, 10 ₽, 11 ₽ верификация карты)
 * 2. Изменения цены (инфляция, скидочные промо-периоды, смена тарифа)
 * 3. Смещения дат (разная длина месяцев, задержки ретраев при нехватке средств, выходные)
 * 4. Надежного отделения от переводов родственникам (P2P / СБП / переводы близким)
 */

import * as XLSX from 'xlsx';

// ===== ИНТЕРФЕЙСЫ ТРАНЗАКЦИЙ И ПОДПИСОК =====

export interface SberbankTransaction {
  id?: string;
  date: string;         // 'YYYY-MM-DD' или 'DD.MM.YYYY'
  amount: number;       // сумма списания (положительное число)
  description: string;  // текст операции из выписки банка
  mcc?: string;         // МСС-код операции
}

export interface SubscriptionGroup {
  serviceTitle: string;
  currentPrice: number;        // Актуальная (последняя) цена подписки
  hasTrial: boolean;           // Было ли промо/пробное списание 1-15 ₽
  priceChanged: boolean;       // Менялась ли цена за анализируемый период
  priceHistory: number[];      // История сумм (например [1, 299, 299, 399])
  totalSpent6Months: number;
  chargeCount: number;
  transactions: SberbankTransaction[];
  classification: {
    isSubscription: boolean;
    isRelativeOrP2P: boolean;
    detectedPersonName?: string;
    detectedServiceName?: string;
    reasons: string[];
  };
}

export interface Subscription {
  name: string;
  category: 'Кино' | 'Музыка' | 'Книги' | 'Другое';
  clusterInfo: string;
  price: number;
  icon: string;
  color: string;
  frequency?: string;
  detectedDates?: string[];
  hasTrial?: boolean;
  priceChanged?: boolean;
  priceHistory?: number[];
  totalSpent6Months?: number;
  chargeCount?: number;
  reasons?: string[];
}

export interface StatementAnalysisResult {
  subscriptions: Subscription[];
  relativesTransfers?: SubscriptionGroup[]; // переводы родственникам/P2P, отделенные от подписок
  statementSummary: {
    bankDetected: string;
    period: string;
    totalTransactionsAnalyzed: number;
    filteredRelativesCount?: number;
    note: string;
  };
}

interface CatalogItem {
  name: string;
  category: 'Кино' | 'Музыка' | 'Книги' | 'Другое';
  icon: string;
  color: string;
  patterns: RegExp[];
  defaultPrice: number;
}

// ===== ПРАВИЛА И МАРКЕРЫ ДЛЯ ОТДЕЛЕНИЯ ПЕРЕВОДОВ РОДСТВЕННИКАМ (P2P / СБП) =====

// МСС-коды переводов между физлицами
export const P2P_MCC_CODES = new Set(['6536', '6538', '4829', '6540', '6012']);

// МСС-коды подписок и стримингов
export const SUBSCRIPTION_MCC_CODES = new Set(['4899', '5815', '5816', '5735', '5968', '7997', '7372']);

// Маркеры человека и родства с поддержкой кириллицы
export const PATRONYMIC_REGEX = /(?:^|[\s:;,])([А-ЯЁ][а-яё]*(?:ович|евич|овна|евна|ич|ична|кызы|оглы))(?:[\s:;,.!?()]|$)/iu;
export const SBER_FIO_FULL = /(?:^|[\s:;,])([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ]\.?)(?:[\s:;,.]|$)/iu;
export const SBER_NAME_WITH_INITIAL = /(?:^|[\s:;,])([А-ЯЁ][а-яё]{2,})\s+([А-ЯЁ]\.?)(?:[\s:;,.]|$)/iu;

export const P2P_PHRASES = [
  'перевод клиенту сбербанка',
  'перевод клиенту банка',
  'перевод через сбп',
  'перевод по сбп',
  'перевод с карты на карту',
  'перевод по номеру телефона',
  'перевод частному лицу',
  'p2p sberbank',
  'sbol transfer',
  'перевод клиенту',
  'перевод физлицу',
];

export const FAMILY_KEYWORDS = [
  'маме', 'мама', 'папе', 'папа', 'сестре', 'брату', 'сыну',
  'дочке', 'дочери', 'бабушке', 'дедушке', 'жене', 'мужу', 'на карманные'
];

export const KNOWN_SUBSCRIPTIONS: Record<string, string> = {
  'yandex': 'Яндекс Плюс',
  'яндекс': 'Яндекс Плюс',
  'sberprime': 'СберПрайм',
  'сберпрайм': 'СберПрайм',
  'vk music': 'VK Музыка',
  'вк музыка': 'VK Музыка',
  'telegram': 'Telegram Premium',
  'okko': 'Okko',
  'ivi': 'IVI',
  'иви': 'IVI',
  'kinopoisk': 'Кинопоиск',
  'kion': 'KION',
  'premier': 'Premier',
  'litres': 'Литрес',
  'apple.com/bill': 'Apple Services',
  'fitmost': 'Fitmost',
  'ozon premium': 'Ozon Premium',
  'букмейт': 'Букмейт',
  'bookmate': 'Букмейт',
  'wink': 'Wink',
  'start.ru': 'Start',
  'mybook': 'MyBook',
  'звук': 'СберЗвук',
  'sberzvuk': 'СберЗвук',
  'уведомления об операциях': 'Плата за уведомления',
  'плата за уведомления': 'Плата за уведомления',
  'sms-bank': 'Плата за уведомления',
  'мобильный банк': 'Плата за уведомления',
};

// Каталог для категоризации, иконок и цветов в интерфейсе
export const SUBSCRIPTION_CATALOG: CatalogItem[] = [
  // --- Кино и ТВ ---
  {
    name: 'Яндекс Плюс',
    category: 'Кино',
    icon: 'fa-film',
    color: '#ffcc00',
    patterns: [
      /yndx[\s*._-]*plus/i,
      /yandex[\s*._-]*plus/i,
      /яндекс[\s*._-]*плюс/i,
      /кинопоиск/i,
      /kinopoisk/i,
      /yandex\*plus/i,
    ],
    defaultPrice: 299,
  },
  {
    name: 'Иви (ivi.ru)',
    category: 'Кино',
    icon: 'fa-film',
    color: '#ea1d5d',
    patterns: [/ivi\.ru/i, /\bivi\b/i, /\bиви\b/i, /онлайн-кинотеатр\s*иви/i],
    defaultPrice: 399,
  },
  {
    name: 'Okko',
    category: 'Кино',
    icon: 'fa-tv',
    color: '#6a0dad',
    patterns: [/\bokko\b/i, /\bокко\b/i, /okko\.tv/i],
    defaultPrice: 499,
  },
  {
    name: 'Premier',
    category: 'Кино',
    icon: 'fa-film',
    color: '#00a8e8',
    patterns: [/\bpremier\b/i, /премьер/i, /premier\.one/i],
    defaultPrice: 299,
  },
  {
    name: 'Start',
    category: 'Кино',
    icon: 'fa-film',
    color: '#ff4500',
    patterns: [/start\.ru/i, /\bстарт\b/i, /start[\s*._-]*media/i],
    defaultPrice: 499,
  },
  {
    name: 'KION',
    category: 'Кино',
    icon: 'fa-tv',
    color: '#e30611',
    patterns: [/\bkion\b/i, /\bкион\b/i, /kion\.ru/i],
    defaultPrice: 249,
  },
  {
    name: 'Wink',
    category: 'Кино',
    icon: 'fa-tv',
    color: '#ff5c00',
    patterns: [/\bwink\b/i, /\bвинк\b/i, /wink\.ru/i, /ростелеком.*wink/i],
    defaultPrice: 399,
  },
  {
    name: 'Amediateka',
    category: 'Кино',
    icon: 'fa-film',
    color: '#1a1a1a',
    patterns: [/amediateka/i, /амедиатека/i],
    defaultPrice: 599,
  },
  {
    name: 'Netflix',
    category: 'Кино',
    icon: 'fa-film',
    color: '#E50914',
    patterns: [/netflix/i, /\bnflx\b/i],
    defaultPrice: 999,
  },
  {
    name: 'YouTube Premium',
    category: 'Кино',
    icon: 'fa-brands fa-youtube',
    color: '#ff0000',
    patterns: [/youtube[\s*._-]*premium/i, /google[\s*._-]*youtube/i],
    defaultPrice: 399,
  },

  // --- Музыка ---
  {
    name: 'VK Музыка',
    category: 'Музыка',
    icon: 'fa-headphones',
    color: '#0077FF',
    patterns: [
      /vk[\s*._-]*music/i,
      /vk[\s*._-]*pay/i,
      /вк[\s*._-]*музык/i,
      /boom[\s*._-]*music/i,
      /vkontakte/i,
      /вконтакте.*подписк/i,
    ],
    defaultPrice: 169,
  },
  {
    name: 'Яндекс Музыка',
    category: 'Музыка',
    icon: 'fa-music',
    color: '#ffcc00',
    patterns: [/yandex[\s*._-]*music/i, /яндекс[\s*._-]*музык/i, /music\.yandex/i],
    defaultPrice: 299,
  },
  {
    name: 'СберЗвук (Звук)',
    category: 'Музыка',
    icon: 'fa-headphones',
    color: '#21a038',
    patterns: [/сберзвук/i, /sberzvuk/i, /\bзвук\b.*подписк/i, /zvuk\.com/i],
    defaultPrice: 199,
  },
  {
    name: 'Apple Music',
    category: 'Музыка',
    icon: 'fa-brands fa-apple',
    color: '#fc3c44',
    patterns: [/apple[\s*._-]*music/i, /itunes[\s*._-]*music/i],
    defaultPrice: 169,
  },
  {
    name: 'Spotify',
    category: 'Музыка',
    icon: 'fa-brands fa-spotify',
    color: '#1DB954',
    patterns: [/spotify/i],
    defaultPrice: 299,
  },

  // --- Книги ---
  {
    name: 'ЛитРес',
    category: 'Книги',
    icon: 'fa-book',
    color: '#ff6600',
    patterns: [/litres/i, /литрес/i, /litres\.ru/i],
    defaultPrice: 399,
  },
  {
    name: 'Букмейт (Bookmate)',
    category: 'Книги',
    icon: 'fa-book-open',
    color: '#ff8c42',
    patterns: [/bookmate/i, /букмейт/i],
    defaultPrice: 399,
  },
  {
    name: 'Storytel',
    category: 'Книги',
    icon: 'fa-book-open',
    color: '#ff8c42',
    patterns: [/storytel/i, /сторител/i],
    defaultPrice: 549,
  },
  {
    name: 'Строки (МТС)',
    category: 'Книги',
    icon: 'fa-book',
    color: '#e30611',
    patterns: [/строки.*мтс/i, /мтс.*строки/i, /stroki\.mts/i],
    defaultPrice: 299,
  },
  {
    name: 'MyBook',
    category: 'Книги',
    icon: 'fa-book',
    color: '#2575fc',
    patterns: [/mybook/i, /майбук/i],
    defaultPrice: 399,
  },

  // --- Банковские подписки и комиссии ---
  {
    name: 'СберПрайм / Прайм+',
    category: 'Другое',
    icon: 'fa-crown',
    color: '#21a038',
    patterns: [/sberprime/i, /sbprime/i, /сберпрайм/i, /сбер\s*прайм/i, /прайм\+/i],
    defaultPrice: 399,
  },
  {
    name: 'Т-Банк Pro (Tinkoff Pro)',
    category: 'Другое',
    icon: 'fa-shield-halved',
    color: '#ffdd2d',
    patterns: [/tinkoff[\s*._-]*pro/i, /t-pro/i, /тинькофф[\s*._-]*про/i, /т-банк[\s*._-]*про/i],
    defaultPrice: 299,
  },
  {
    name: 'Плата за уведомления (СМС-банк)',
    category: 'Другое',
    icon: 'fa-receipt',
    color: '#21a038',
    patterns: [
      /плата\s*за\s*уведомлен/i,
      /мобильный\s*банк/i,
      /смс-банк/i,
      /информирование\s*об\s*операциях/i,
      /уведомления\s*по\s*карте/i,
      /sms[\s*._-]*bank/i,
    ],
    defaultPrice: 99,
  },
  {
    name: 'Обслуживание карты/счёта',
    category: 'Другое',
    icon: 'fa-credit-card',
    color: '#114d26',
    patterns: [
      /плата\s*за\s*обслуживан/i,
      /комиссия\s*за\s*обслуживан/i,
      /пакет\s*услуг.*обслуживание/i,
    ],
    defaultPrice: 150,
  },

  // --- Облачные сервисы и IT ---
  {
    name: 'Telegram Premium',
    category: 'Другое',
    icon: 'fa-paper-plane',
    color: '#24A1DE',
    patterns: [/telegram[\s*._-]*premium/i, /t\.me\/premium/i, /telegram\s*fzco/i],
    defaultPrice: 299,
  },
  {
    name: 'Apple Services (iCloud/App Store)',
    category: 'Другое',
    icon: 'fa-brands fa-apple',
    color: '#555555',
    patterns: [/apple\.com\/bill/i, /itunes/i, /icloud/i, /apple\s*services/i],
    defaultPrice: 169,
  },
  {
    name: 'Облако Mail.ru',
    category: 'Другое',
    icon: 'fa-cloud',
    color: '#005ff9',
    patterns: [/cloud\.mail\.ru/i, /облако[\s*._-]*mail/i, /mail\.ru\s*cloud/i],
    defaultPrice: 99,
  },
  {
    name: 'Яндекс 360 (Диск)',
    category: 'Другое',
    icon: 'fa-cloud',
    color: '#ffcc00',
    patterns: [/yandex[\s*._-]*360/i, /яндекс[\s*._-]*360/i, /диск\s*яндекс/i, /yandex[\s*._-]*disk/i],
    defaultPrice: 199,
  },
  {
    name: 'Google One / Storage',
    category: 'Другое',
    icon: 'fa-brands fa-google',
    color: '#4285f4',
    patterns: [/google[\s*._-]*one/i, /google[\s*._-]*storage/i, /google\*services/i],
    defaultPrice: 199,
  },
  {
    name: 'ChatGPT Plus (OpenAI)',
    category: 'Другое',
    icon: 'fa-robot',
    color: '#10a37f',
    patterns: [/openai/i, /chatgpt/i],
    defaultPrice: 1990,
  },

  // --- Мобильная связь (абонентская плата) ---
  {
    name: 'МТС (Абонентская плата)',
    category: 'Другое',
    icon: 'fa-mobile-screen',
    color: '#e30611',
    patterns: [/mts[\s*._-]*premium/i, /оплата[\s*._-]*мтс/i, /\bmts\b.*(связь|абонент|тариф)/i],
    defaultPrice: 450,
  },
  {
    name: 'МегаФон (Тариф)',
    category: 'Другое',
    icon: 'fa-mobile-screen',
    color: '#00b956',
    patterns: [/megafon.*(тариф|абонент)/i, /оплата.*мегафон/i],
    defaultPrice: 500,
  },
  {
    name: 'Билайн (Тариф)',
    category: 'Другое',
    icon: 'fa-mobile-screen',
    color: '#ffc800',
    patterns: [/beeline.*(тариф|абонент)/i, /оплата.*билайн/i],
    defaultPrice: 480,
  },
  {
    name: 'Tele2 / T2 (Тариф)',
    category: 'Другое',
    icon: 'fa-mobile-screen',
    color: '#1f2229',
    patterns: [/tele2.*(тариф|абонент)/i, /т2.*(тариф|абонент)/i],
    defaultPrice: 450,
  },
];

// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ АНАЛИЗА =====

/**
 * Парсер даты в миллисекунды с надежной поддержкой российского формата DD.MM.YYYY и ISO
 */
export function parseDateToMs(dateStr: string): number {
  if (!dateStr) return NaN;
  const str = String(dateStr).trim();
  const dotParts = str.match(/^(\d{2})[./](\d{2})[./](\d{4})$/);
  if (dotParts) {
    return new Date(Number(dotParts[3]), Number(dotParts[2]) - 1, Number(dotParts[1])).getTime();
  }
  const isoParts = str.match(/^(\d{4})[./-](\d{2})[./-](\d{2})$/);
  if (isoParts) {
    return new Date(Number(isoParts[1]), Number(isoParts[2]) - 1, Number(isoParts[3])).getTime();
  }
  const t = new Date(str).getTime();
  return isNaN(t) ? 0 : t;
}

/**
 * 1. Нормализация мерчанта/получателя
 * Очищает динамические идентификаторы авторизаций, даты и суммы из текста
 */
export function getMerchantFingerprint(desc: string): string {
  const lower = desc.toLowerCase();

  // Если это известный сервис — группируем строго по сервису
  for (const [key, name] of Object.entries(KNOWN_SUBSCRIPTIONS)) {
    if (lower.includes(key)) return `service:${name}`;
  }

  // Если распознано ФИО человека — группируем по человеку
  const person = extractPersonName(desc);
  if (person) return `person:${person.toLowerCase()}`;

  // Иначе очищаем от дат, времени и кодов авторизации
  return lower
    .replace(/сбербанк онлайн\.?\s*/i, '')
    .replace(/mcc\s*\d+/gi, '')
    .replace(/\b\d{2}\.\d{2}\.\d{4}\b/g, '')
    .replace(/\b\d{2}:\d{2}(:\d{2})?\b/g, '')
    .replace(/auth\s*[a-z0-9]+/gi, '')
    .replace(/[*#]\d+/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Извлечение ФИО или имени получателя из текста операции
 */
export function extractPersonName(text: string): string | null {
  const prefixMatch = text.match(/(?:клиенту(?:\s+сбербанка)?|сбп|перевод)\s*:\s*([А-ЯЁа-яё\s.]+)/i);
  if (prefixMatch && prefixMatch[1]) {
    const candidate = prefixMatch[1].trim();
    if (SBER_FIO_FULL.test(candidate) || SBER_NAME_WITH_INITIAL.test(candidate) || PATRONYMIC_REGEX.test(candidate)) {
      return candidate.replace(/[,;]+$/, '').trim();
    }
  }
  const fullMatch = text.match(SBER_FIO_FULL);
  if (fullMatch) return fullMatch[0].trim();
  const initialMatch = text.match(SBER_NAME_WITH_INITIAL);
  if (initialMatch) return initialMatch[0].trim();
  const patronymicMatch = text.match(PATRONYMIC_REGEX);
  if (patronymicMatch) return patronymicMatch[0].trim();
  return null;
}

/**
 * 2. Умная проверка периодичности с учетом смещения дат
 * Допускает плавающие даты:
 * - 6-16 дней для триала или недельных списаний (7-14 дней)
 * - 25-35 дней для стандартного месяца (разная длина месяцев + выходные)
 * - 36-45 дней при задержке ретрая списания банком
 */
export function hasRecurringTimeline(dates: number[]): boolean {
  if (dates.length < 2) return false;
  const sorted = [...dates].sort((a, b) => a - b);

  let validIntervals = 0;
  for (let i = 1; i < sorted.length; i++) {
    const diffDays = Math.round((sorted[i] - sorted[i - 1]) / (1000 * 60 * 60 * 24));

    if ((diffDays >= 6 && diffDays <= 16) || (diffDays >= 25 && diffDays <= 45)) {
      validIntervals++;
    }
  }

  // Если хотя бы 60% интервалов укладываются в подписной шаг
  return validIntervals / (sorted.length - 1) >= 0.6;
}

/**
 * 3. Умный анализ цен: триалы (1-15 ₽) и изменение стоимости (дороже/дешевле)
 */
export function analyzePricePattern(amounts: number[]) {
  // Выявляем пробные верификационные списания (1 ₽, 10 ₽, 11 ₽, 15 ₽)
  const trialCharges = amounts.filter((amt) => amt <= 15);
  const regularCharges = amounts.filter((amt) => amt > 15);

  const hasTrial = trialCharges.length > 0;

  // Проверяем, менялась ли цена среди регулярных списаний (например, было 299, стало 399)
  const uniqueRegularPrices = Array.from(new Set(regularCharges.map(Math.round)));
  const priceChanged = uniqueRegularPrices.length > 1;

  // Подписка считается валидной по суммам, если:
  // а) Есть хотя бы 1 регулярный платеж после триала
  // б) Либо регулярные платежи имеют характерную регулярность (повторение сумм)
  const isValidPricing =
    (regularCharges.length >= 1 && hasTrial) ||
    (regularCharges.length >= 2 && !hasTrial && uniqueRegularPrices.length <= 3) ||
    (amounts.length >= 2 && uniqueRegularPrices.length <= 2);

  const currentPrice =
    regularCharges.length > 0
      ? regularCharges[regularCharges.length - 1]
      : amounts[amounts.length - 1];

  return {
    hasTrial,
    priceChanged,
    isValidPricing,
    currentPrice,
  };
}

/**
 * 4. Скоринг операции: подписка против перевода человеку / родственнику
 */
export function classifyTransaction(tx: SberbankTransaction) {
  const desc = tx.description.toLowerCase();
  const mcc = tx.mcc?.trim();
  const reasons: string[] = [];
  let score = 0;

  let detectedServiceName: string | undefined;
  for (const [key, name] of Object.entries(KNOWN_SUBSCRIPTIONS)) {
    if (desc.includes(key)) {
      score += 85;
      detectedServiceName = name;
      reasons.push(`Распознан сервис: ${name}`);
      break;
    }
  }

  if (mcc) {
    if (P2P_MCC_CODES.has(mcc)) {
      score -= 75;
      reasons.push(`МСС ${mcc}: перевод физлицу (P2P)`);
    } else if (SUBSCRIPTION_MCC_CODES.has(mcc)) {
      score += 65;
      reasons.push(`МСС ${mcc}: подписной/медиа сервис`);
    }
  }

  for (const phrase of P2P_PHRASES) {
    if (desc.includes(phrase)) {
      score -= 75;
      reasons.push(`Назначение платежа физлицу: «${phrase}»`);
      break;
    }
  }

  const person = extractPersonName(tx.description);
  let detectedPersonName: string | undefined;
  if (person) {
    score -= 80;
    detectedPersonName = person;
    reasons.push(`Получатель — человек: ${person}`);
  }

  for (const fam of FAMILY_KEYWORDS) {
    const famRegex = new RegExp(`(?:^|[^А-ЯЁа-яёA-Za-z0-9])${fam}(?:[^А-ЯЁа-яёA-Za-z0-9]|$)`, 'i');
    if (famRegex.test(desc)) {
      score -= 50;
      reasons.push(`Семейный комментарий: «${fam}»`);
      break;
    }
  }

  // Триальные 1-15 ₽ практически никогда не отправляют родственникам как регулярный платеж
  if (tx.amount <= 15) {
    score += 40;
    reasons.push('Сумма 1-15 ₽ характерна для пробного периода или привязки карты');
  }

  return {
    isSubscription: score >= 20,
    isRelativeOrP2P: score <= -20,
    detectedPersonName,
    detectedServiceName,
    reasons,
  };
}

/**
 * ГЛАВНАЯ ФУНКЦИЯ ФИЛЬТРАЦИИ И КЛАССИФИКАЦИИ ВЫПИСКИ
 * Разделяет регулярные операции на коммерческие подписки и переводы родственникам (P2P/СБП)
 */
export function filterSubscriptionsFromStatement(transactions: SberbankTransaction[]) {
  const debits = transactions.filter((t) => t.amount > 0);

  // Группируем по мерчанту / человеку
  const groups = new Map<string, SberbankTransaction[]>();
  for (const tx of debits) {
    const key = getMerchantFingerprint(tx.description);
    const list = groups.get(key) || [];
    list.push(tx);
    groups.set(key, list);
  }

  const subscriptions: SubscriptionGroup[] = [];
  const relativesTransfers: SubscriptionGroup[] = [];

  for (const [, txList] of groups.entries()) {
    if (txList.length < 2) continue; // Минимум 2 списания для анализа регулярности

    // Хронологическая сортировка
    txList.sort((a, b) => parseDateToMs(a.date) - parseDateToMs(b.date));

    const dates = txList.map((t) => parseDateToMs(t.date)).filter((t) => !isNaN(t) && t > 0);
    const amounts = txList.map((t) => t.amount);

    // 1. Проверяем таймлайн (с учетом смещения дат)
    if (!hasRecurringTimeline(dates)) continue;

    // 2. Проверяем суммы (с учетом триалов 1-15 ₽ и изменений цены)
    const priceAnalysis = analyzePricePattern(amounts);
    if (!priceAnalysis.isValidPricing) continue;

    // 3. Классифицируем (подписка vs родственник)
    // Берем транзакцию с регулярной ценой, либо последнюю
    const representativeTx = txList.find((t) => t.amount > 15) || txList[txList.length - 1];
    const classification = classifyTransaction(representativeTx);

    const title =
      classification.detectedServiceName ||
      (classification.detectedPersonName
        ? `Перевод: ${classification.detectedPersonName}`
        : representativeTx.description);

    const group: SubscriptionGroup = {
      serviceTitle: title,
      currentPrice: priceAnalysis.currentPrice,
      hasTrial: priceAnalysis.hasTrial,
      priceChanged: priceAnalysis.priceChanged,
      priceHistory: amounts,
      totalSpent6Months: amounts.reduce((a, b) => a + b, 0),
      chargeCount: txList.length,
      transactions: txList,
      classification,
    };

    if (classification.isSubscription) {
      subscriptions.push(group);
    } else if (classification.isRelativeOrP2P) {
      relativesTransfers.push(group);
    }
  }

  subscriptions.sort((a, b) => b.currentPrice - a.currentPrice);
  relativesTransfers.sort((a, b) => b.currentPrice - a.currentPrice);

  return { subscriptions, relativesTransfers };
}

// ===== УТИЛИТЫ ЧТЕНИЯ И ИЗВЛЕЧЕНИЯ ИЗ ТАБЛИЦ =====

/**
 * Extracts a numeric amount from row values or string
 */
function extractAmount(val: any): number | null {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') {
    const abs = Math.abs(val);
    if (abs > 0 && abs < 1000000) return abs;
    return null;
  }
  const str = String(val).trim();
  const cleaned = str
    .replace(/[₽$€]|rub|руб|rur/gi, '')
    .replace(/\s/g, '')
    .replace(',', '.');

  const match = cleaned.match(/-?\d+(\.\d+)?/);
  if (match) {
    const num = Math.abs(parseFloat(match[0]));
    if (num >= 1 && num <= 500000 && num !== 2024 && num !== 2025 && num !== 2026) {
      return num;
    }
  }
  return null;
}

/**
 * Detects date string from row or cell
 */
function extractDate(val: any): string | null {
  if (!val) return null;
  const str = String(val).trim();
  const match = str.match(/\b(\d{2}[./-]\d{2}[./-]\d{2,4}|\d{4}[./-]\d{2}[./-]\d{2})\b/);
  if (match) {
    return match[1];
  }
  return null;
}

/**
 * Extracts MCC code if present in cell or row
 */
function extractMcc(val: any): string | undefined {
  if (!val) return undefined;
  const str = String(val).trim();
  const match = str.match(/\bmcc\s*[:=-]?\s*(\d{4})\b/i) || str.match(/\b(4899|5815|5816|5735|5968|7997|7372|6536|6538|4829|6540|6012)\b/);
  if (match) {
    return match[1];
  }
  return undefined;
}

/**
 * Decodes ArrayBuffer to string with automatic UTF-8 or Windows-1251 detection
 */
export function decodeBufferToString(buffer: ArrayBuffer): string {
  try {
    const utfDecoder = new TextDecoder('utf-8', { fatal: true });
    return utfDecoder.decode(buffer);
  } catch (e) {
    try {
      const winDecoder = new TextDecoder('windows-1251');
      return winDecoder.decode(buffer);
    } catch (e2) {
      return new TextDecoder().decode(buffer);
    }
  }
}

/**
 * Detects bank from text or metadata
 */
function detectBank(text: string): string {
  const lower = text.toLowerCase();
  if (lower.includes('сбербанк') || lower.includes('sber') || lower.includes('сбер')) {
    return 'СберБанк';
  }
  if (lower.includes('тинькофф') || lower.includes('т-банк') || lower.includes('t-bank') || lower.includes('tcs')) {
    return 'Т-Банк';
  }
  if (lower.includes('втб') || lower.includes('vtb')) {
    return 'ВТБ';
  }
  if (lower.includes('альфа') || lower.includes('alfa')) {
    return 'Альфа-Банк';
  }
  if (lower.includes('газпромбанк')) {
    return 'Газпромбанк';
  }
  if (lower.includes('райффайзен') || lower.includes('raiffeisen')) {
    return 'Райффайзенбанк';
  }
  if (lower.includes('озон') || lower.includes('ozon')) {
    return 'Озон Банк';
  }
  if (lower.includes('яндекс банк') || lower.includes('yandex bank')) {
    return 'Яндекс Банк';
  }
  return 'Банковская выписка';
}

// ===== ГЛАВНЫЙ КЛИЕНТСКИЙ ПАРСЕР ТАБЛИЦ ВЫПИСОК (SheetJS) =====

export async function parseStatementFileClient(
  file: File
): Promise<StatementAnalysisResult> {
  const buffer = await file.arrayBuffer();
  const filename = file.name.toLowerCase();

  let workbook: XLSX.WorkBook;
  let rawTextForMetadata = '';

  const isCsvOrText = filename.endsWith('.csv') || filename.endsWith('.txt');

  if (isCsvOrText) {
    const decodedText = decodeBufferToString(buffer);
    rawTextForMetadata = decodedText;
    workbook = XLSX.read(decodedText, { type: 'string' });
  } else {
    workbook = XLSX.read(buffer, { type: 'array' });
  }

  const allRows: any[][] = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (sheet) {
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: '',
      }) as any[][];
      allRows.push(...rows);
    }
  }

  if (allRows.length === 0) {
    throw new Error('Файл не содержит табличных данных или пуст.');
  }

  if (!rawTextForMetadata) {
    rawTextForMetadata = allRows
      .slice(0, 30)
      .map((r) => r.join(' '))
      .join('\n');
  }

  const bankDetected = detectBank(rawTextForMetadata);

  // Извлекаем все транзакции из строк
  const extractedTransactions: SberbankTransaction[] = [];
  const detectedAllDates: string[] = [];

  for (const row of allRows) {
    if (!Array.isArray(row) || row.length === 0) continue;
    const rowString = row.map((c) => String(c)).join(' ');
    if (rowString.trim().length < 4) continue;

    let rowDate: string | null = null;
    let rowAmount: number | null = null;
    let rowMcc: string | undefined = undefined;
    let descriptionCandidate = '';

    for (const cell of row) {
      const cellStr = String(cell).trim();
      if (!rowDate) {
        const d = extractDate(cellStr);
        if (d) rowDate = d;
      }
      if (!rowAmount) {
        const a = extractAmount(cell);
        if (a) rowAmount = a;
      }
      if (!rowMcc) {
        const m = extractMcc(cellStr);
        if (m) rowMcc = m;
      }
      if (cellStr.length > 3 && isNaN(Number(cellStr)) && !extractDate(cellStr)) {
        if (!descriptionCandidate || cellStr.length > descriptionCandidate.length) {
          descriptionCandidate = cellStr;
        }
      }
    }

    if (!rowMcc) {
      rowMcc = extractMcc(rowString);
    }

    if (rowDate && rowAmount && rowAmount > 0) {
      detectedAllDates.push(rowDate);
      extractedTransactions.push({
        date: rowDate,
        amount: rowAmount,
        description: descriptionCandidate || rowString.slice(0, 100),
        mcc: rowMcc,
      });
    }
  }

  // Запускаем алгоритм классификации и фильтрации переводов родственникам
  const { subscriptions: detectedSubGroups, relativesTransfers } =
    filterSubscriptionsFromStatement(extractedTransactions);

  // Преобразуем обнаруженные подписки в формат Subscription для интерфейса
  const subscriptions: Subscription[] = [];
  const processedNames = new Set<string>();

  for (const group of detectedSubGroups) {
    const rawTitle = group.serviceTitle;
    let matchedCatalog = SUBSCRIPTION_CATALOG.find((cat) =>
      cat.patterns.some((p) => p.test(rawTitle) || p.test(group.transactions[0]?.description || ''))
    );

    const displayName = matchedCatalog?.name || rawTitle;
    if (processedNames.has(displayName.toLowerCase())) continue;
    processedNames.add(displayName.toLowerCase());

    // Формируем понятное описание с деталями проверки
    const detailsParts: string[] = [];
    if (group.hasTrial) {
      detailsParts.push(`Пробный период (${Math.min(...group.priceHistory)} ₽)`);
    }
    if (group.priceChanged) {
      const min = Math.min(...group.priceHistory.filter((p) => p > 15));
      const max = Math.max(...group.priceHistory);
      if (min !== max) {
        detailsParts.push(`Смена тарифа (${min} ₽ ➔ ${max} ₽)`);
      }
    }
    detailsParts.push(`${group.chargeCount} списан.`);

    const clusterInfo = `${detailsParts.join(' • ')}: ${group.transactions
      .map((t) => t.description.slice(0, 35))
      .slice(0, 2)
      .join(' | ')}`;

    // Определение категории
    let category = matchedCatalog?.category || 'Другое';
    if (category === 'Другое') {
      const low = displayName.toLowerCase();
      if (/кино|film|tv|cinema|сериал|okko|ivi|kion|wink|premier/i.test(low)) category = 'Кино';
      else if (/музык|music|zvuk|звук|spotify/i.test(low)) category = 'Музыка';
      else if (/книг|book|read|литрес|litres|букмейт/i.test(low)) category = 'Книги';
    }

    subscriptions.push({
      name: displayName,
      category,
      clusterInfo,
      price: Math.round(group.currentPrice),
      icon: matchedCatalog?.icon || (category === 'Кино' ? 'fa-film' : category === 'Музыка' ? 'fa-music' : category === 'Книги' ? 'fa-book' : 'fa-receipt'),
      color: matchedCatalog?.color || '#21a038',
      frequency: 'Ежемесячно',
      detectedDates: group.transactions.map((t) => t.date).slice(0, 5),
      hasTrial: group.hasTrial,
      priceChanged: group.priceChanged,
      priceHistory: group.priceHistory,
      totalSpent6Months: group.totalSpent6Months,
      chargeCount: group.chargeCount,
      reasons: group.classification.reasons,
    });
  }

  // Также проверяем единичные списания из каталога (если загружена короткая выписка за 1 месяц)
  // При этом ОБЯЗАТЕЛЬНО проверяем, чтобы это НЕ был перевод человеку (P2P)
  for (const tx of extractedTransactions) {
    const classification = classifyTransaction(tx);
    if (classification.isRelativeOrP2P) continue; // Защита: никогда не добавляем переводы человеку

    for (const cat of SUBSCRIPTION_CATALOG) {
      const matches = cat.patterns.some((p) => p.test(tx.description));
      if (matches && !processedNames.has(cat.name.toLowerCase())) {
        processedNames.add(cat.name.toLowerCase());
        subscriptions.push({
          name: cat.name,
          category: cat.category,
          clusterInfo: `Найдено в операциях: ${tx.description.slice(0, 50)}`,
          price: tx.amount > 15 ? Math.round(tx.amount) : cat.defaultPrice,
          icon: cat.icon,
          color: cat.color,
          frequency: 'Ежемесячно',
          detectedDates: [tx.date],
          hasTrial: tx.amount <= 15,
        });
        break;
      }
    }
  }

  // Сортировка по стоимости
  subscriptions.sort((a, b) => b.price - a.price);

  let periodStr = 'По предоставленной выписке';
  if (detectedAllDates.length > 0) {
    const sorted = [...detectedAllDates].sort();
    periodStr = `${sorted[0]} — ${sorted[sorted.length - 1]}`;
  }

  const note =
    subscriptions.length > 0
      ? `Успешно выявлено ${subscriptions.length} регулярных подписок. Отделено переводов родственникам/P2P: ${relativesTransfers.length}.`
      : 'В предоставленной выписке не обнаружено регулярных периодических списаний.';

  return {
    subscriptions,
    relativesTransfers,
    statementSummary: {
      bankDetected,
      period: periodStr,
      totalTransactionsAnalyzed: Math.max(extractedTransactions.length, allRows.length),
      filteredRelativesCount: relativesTransfers.length,
      note,
    },
  };
}

// ===== КЛИЕНТСКИЕ ГЕНЕРАТОРЫ ЗАЯВЛЕНИЙ НА ОТКАЗ ОТ ПОДПИСОК =====

/**
 * Client-Side Legal Cancellation Letter Generator
 * Complies with Art. 32 of RF Law "On Protection of Consumer Rights" & Art. 782 of Civil Code
 */
export function generateClientCancellationLetter(
  serviceName: string,
  monthlyPrice: number
): string {
  const today = new Date().toLocaleDateString('ru-RU');
  return `В службу поддержки сервиса «${serviceName}»
От: [Ваши ФИО]
Эл. почта: [Ваш email, привязанный к аккаунту]
Телефон: [Ваш номер телефона]

ТРЕБОВАНИЕ ОБ ОТКАЗЕ ОТ ПОДПИСКИ И ОТЗЫВЕ СОГЛАСИЯ НА АВТОСПИСАНИЯ

Я, [Ваши ФИО], являюсь пользователем сервиса «${serviceName}». Настоящим уведомляю вас о своем отказе от платных услуг сервиса (стоимостью ${monthlyPrice} рублей в месяц) и требую:

1. Немедленно прекратить действие платной подписки на сервис «${serviceName}» в отношении моего аккаунта.
2. Отключить функцию автоматического продления подписки (автоплатеж / безакцептные списания).
3. Удалить из моей учетной записи все привязанные банковские карты и платежные реквизиты.
4. Прекратить обработку моих платежных данных в целях совершения любых последующих списаний.

ПРАВОВОЕ ОБОСНОВАНИЕ:
В соответствии со статьей 32 Закона РФ «О защите прав потребителей» и статьей 782 Гражданского кодекса РФ, потребитель вправе отказаться от исполнения договора о выполнении работ (оказании услуг) в любое время при условии оплаты исполнителю фактически понесенных им расходов.

Прошу в срок до 3 рабочих дней направить на мой адрес электронной почты письменное подтверждение расторжения договора и отвязки банковской карты.

В случае продолжения автоматических списаний денежных средств с моих банковских счетов после получения данного требования, я оставляю за собой право обратиться с жалобой в Роспотребнадзор, Банк России и правоохранительные органы.

Дата: ${today}
Подпись: ____________ / [Ваши ФИО]`;
}

/**
 * Client-Side Bulk Cancellation Letter Generator
 */
export function generateClientBulkCancellationLetter(
  services: Subscription[]
): string {
  const today = new Date().toLocaleDateString('ru-RU');
  const totalMonthly = services.reduce((acc, curr) => acc + curr.price, 0);
  const totalYearly = totalMonthly * 12;

  let letter = `Кому: Службам клиентской поддержки онлайн-сервисов
От: [Ваши ФИО]
Эл. почта: [Ваш email]
Телефон: [Ваш номер телефона]

СВОДНОЕ ТРЕБОВАНИЕ ОБ ОТКАЗЕ ОТ ПЛАТНЫХ ПОДПИСОК

Настоящим заявляю об отказе от продления платных подписок и автоматических списаний со следующих сервисов:

`;

  services.forEach((sub, idx) => {
    letter += `${idx + 1}. «${sub.name}» — ${sub.price} ₽/мес (${sub.category})\n`;
  });

  letter += `
ИТОГО ОБЩАЯ СУММА ЭКОНОМИИ:
• В месяц: ${totalMonthly.toLocaleString('ru-RU')} ₽
• В год: ${totalYearly.toLocaleString('ru-RU')} ₽

На основании ст. 32 Закона РФ «О защите прав потребителей» и ст. 782 ГК РФ требую:
1. Расторгнуть договоры оказания услуг по указанным подпискам.
2. Отключить безакцептное списание денежных средств (рекуррентные автоплатежи).
3. Удалить сохраненные реквизиты банковских карт из профилей.

Дата составления: ${today}
Подпись: ____________ / [Ваши ФИО]`;

  return letter;
}
