/**
 * Сбер Сканер Подписок — Client Engine
 * Real AI bank statement parsing powered by Gemini
 */

import * as XLSX from 'xlsx';
import {
  parseStatementFileClient,
  generateClientCancellationLetter,
  generateClientBulkCancellationLetter,
  Subscription,
  SubscriptionGroup,
} from './statementParser';

declare const Chart: any;

// Fallback pool for demo data preview
const subscriptionsPool: Subscription[] = [
  { name: "Netflix", category: "Кино", clusterInfo: "Сгруппировано из: NFLX, NETFLIX, Netflix.com", price: 999, icon: "fa-film", color: "#E50914" },
  { name: "Яндекс Плюс", category: "Музыка", clusterInfo: "Сгруппировано из: YNDX PLUS, Yandex*Plus", price: 299, icon: "fa-music", color: "#ffcc00" },
  { name: "VK Музыка", category: "Музыка", clusterInfo: "Сгруппировано из: VK PAY, BOOM MUSIC", price: 159, icon: "fa-headphones", color: "#0077FF" },
  { name: "ЛитРес", category: "Книги", clusterInfo: "Сгруппировано из: LITRES.RU", price: 399, icon: "fa-book", color: "#ff6600" },
  { name: "Okko", category: "Кино", clusterInfo: "Сгруппировано из: OKKO, Rambler", price: 499, icon: "fa-tv", color: "#6a0dad" },
  { name: "Apple Music", category: "Музыка", clusterInfo: "Сгруппировано из: APPLE.COM/BILL", price: 169, icon: "fa-apple", color: "#ff2d55" },
  { name: "MEGOGO", category: "Кино", clusterInfo: "Сгруппировано из: MEGOGO, Мегого", price: 399, icon: "fa-video", color: "#00b3b3" },
  { name: "Storytel", category: "Книги", clusterInfo: "Сгруппировано из: STORYTEL", price: 549, icon: "fa-book-open", color: "#ff8c42" },
  { name: "Облако Mail.ru", category: "Другое", clusterInfo: "Сгруппировано из: CLOUD.MAIL.RU", price: 99, icon: "fa-cloud", color: "#005ff9" },
  { name: "СберПрайм", category: "Другое", clusterInfo: "Сгруппировано из: SBPRIME, СберПрайм+", price: 399, icon: "fa-crown", color: "#21a038" }
];

let activeData: Subscription[] = [];
let currentRelativesTransfers: SubscriptionGroup[] = [];
let currentFilter = 'all';
let isAnalyzed = false;
let selectedSubs = new Set<string>();
let categoryChartInstance: any = null;
let topSubsChartInstance: any = null;
let currentEmailText = '';

// DOM Elements
const dropZone = document.getElementById('dropZone') as HTMLElement;
const fileInput = document.getElementById('fileInput') as HTMLInputElement;
const uploadContent = document.getElementById('uploadContent') as HTMLElement;
const loader = document.getElementById('loader') as HTMLElement;
const loaderText = document.getElementById('loaderText') as HTMLElement;
const progressFill = document.getElementById('progressFill') as HTMLElement;
const emptyState = document.getElementById('emptyState') as HTMLElement;
const subsList = document.getElementById('subsList') as HTMLElement;
const totalAmount = document.getElementById('totalAmount') as HTMLElement;
const filtersBar = document.getElementById('filtersBar') as HTMLElement;
const exportRow = document.getElementById('exportRow') as HTMLElement;
const modal = document.getElementById('llmModal') as HTMLElement;
const closeModal = document.getElementById('closeModal') as HTMLElement;
const emailContent = document.getElementById('emailContent') as HTMLElement;
const yearlySavings = document.getElementById('yearlySavings') as HTMLElement;
const header = document.getElementById('header') as HTMLElement;
const backToTop = document.getElementById('backToTop') as HTMLElement;
const toastContainer = document.getElementById('toastContainer') as HTMLElement;
const healthCard = document.getElementById('healthCard') as HTMLElement;
const healthRing = document.getElementById('healthRing') as HTMLElement;
const healthStatus = document.getElementById('healthStatus') as HTMLElement;
const healthDesc = document.getElementById('healthDesc') as HTMLElement;
const selectedSavings = document.getElementById('selectedSavings') as HTMLElement;
const savingsAmount = document.getElementById('savingsAmount') as HTMLElement;

// ===== TOAST SYSTEM =====
function showToast(message: string, type: 'success' | 'warning' | 'info' = 'success') {
  if (!toastContainer) return;
  const toast = document.createElement('div');
  toast.className = 'toast';
  const icon = type === 'success' ? 'fa-circle-check' : type === 'warning' ? 'fa-triangle-exclamation' : 'fa-info-circle';
  toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
  toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.remove();
  }, 4500);
}

// Expose functions to window for inline HTML onclick attributes
(window as any).showToast = showToast;
(window as any).loadDemoData = loadDemoData;
(window as any).checkBudget = checkBudget;
(window as any).filterSubs = filterSubs;
(window as any).toggleSubSelection = toggleSubSelection;
(window as any).exportReport = exportReport;
(window as any).openLLMModal = openLLMModal;
(window as any).copyEmail = copyEmail;
(window as any).generateBulkLetters = generateBulkLetters;

// ===== THEME TOGGLE =====
const themeToggle = document.getElementById('themeToggle');
if (themeToggle) {
  themeToggle.addEventListener('click', () => {
    document.body.classList.toggle('light-theme');
    const icon = document.querySelector('#themeToggle i');
    if (icon) {
      if (document.body.classList.contains('light-theme')) {
        icon.className = 'fa-solid fa-sun';
      } else {
        icon.className = 'fa-solid fa-moon';
      }
    }
    // Update charts with new theme colors
    if (isAnalyzed) {
      updateCharts();
    }
  });
}

// ===== SMOOTH SCROLL & NAV =====
document.querySelectorAll('.nav-link').forEach(link => {
  link.addEventListener('click', function (this: HTMLElement, e: Event) {
    e.preventDefault();
    const href = this.getAttribute('href');
    if (!href) return;
    const targetId = href.substring(1);
    const targetEl = document.getElementById(targetId);
    if (targetEl) {
      targetEl.scrollIntoView({ behavior: 'smooth' });
    }
    const navMenu = document.getElementById('navMenu');
    if (navMenu) {
      navMenu.classList.remove('open');
    }
  });
});

window.addEventListener('scroll', () => {
  if (window.scrollY > 50) header?.classList.add('scrolled');
  else header?.classList.remove('scrolled');
  if (window.scrollY > 600) backToTop?.classList.add('visible');
  else backToTop?.classList.remove('visible');

  const sections = ['top', 'howitworks', 'scanner', 'analytics'];
  let current = 'top';
  sections.forEach(sectionId => {
    const section = document.getElementById(sectionId);
    if (section && window.scrollY >= section.offsetTop - 120) current = sectionId;
  });
  document.querySelectorAll('.nav-link').forEach(link => {
    link.classList.remove('active');
    if (link.getAttribute('href') === '#' + current) link.classList.add('active');
  });
});

if (backToTop) {
  backToTop.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
}

const hamburgerBtn = document.getElementById('hamburgerBtn');
if (hamburgerBtn) {
  hamburgerBtn.addEventListener('click', () => {
    document.getElementById('navMenu')?.classList.toggle('open');
  });
}

// ===== INTERSECTION OBSERVER =====
const observer = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) entry.target.classList.add('visible');
  });
}, { threshold: 0.1 });
document.querySelectorAll('.reveal, .reveal-left, .reveal-right').forEach(el => observer.observe(el));

// ===== DRAG & DROP & FILE HANDLING =====
if (dropZone) {
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
    }, false);
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
  });

  dropZone.addEventListener('drop', (e: DragEvent) => {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  }, false);
}

if (fileInput) {
  fileInput.addEventListener('change', () => {
    if (fileInput.files && fileInput.files.length > 0) {
      handleFiles(fileInput.files);
    }
  }, false);
}

let lastUploadedFile: File | null = null;

function handleFiles(files: FileList) {
  const file = files[0];
  if (!file) return;
  lastUploadedFile = file;
  startRealAnalysis(file);
}

function formatFriendlyError(rawError: string): string {
  if (!rawError) return 'Не удалось проанализировать файл выписки.';
  try {
    if (rawError.trim().startsWith('{')) {
      const parsed = JSON.parse(rawError);
      if (parsed?.error?.message) {
        rawError = parsed.error.message;
      }
    }
  } catch (e) {
    // ignore
  }

  if (rawError.includes('503') || rawError.includes('UNAVAILABLE') || rawError.includes('high demand') || rawError.includes('высокий спрос')) {
    return 'Сервер AI испытывает временный пик нагрузки (код 503). Пожалуйста, нажмите «Повторить попытку» или используйте файл в формате Excel / CSV.';
  }
  if (rawError.includes('429') || rawError.includes('quota') || rawError.includes('Resource has been exhausted')) {
    return 'Превышен лимит запросов к AI. Пожалуйста, подождите минуту и повторите попытку.';
  }
  return rawError;
}

// Global retry function for UI button
(window as any).retryLastFile = () => {
  if (lastUploadedFile) {
    startRealAnalysis(lastUploadedFile);
  } else if (fileInput) {
    fileInput.click();
  }
};

// Global sample loader functions for testing client-side parsing
(window as any).loadSampleXlsx = loadSampleXlsx;
(window as any).loadSampleCsv = loadSampleCsv;
(window as any).toggleRelativesDetails = toggleRelativesDetails;

function renderRelativesAlert(relatives: SubscriptionGroup[]) {
  const alertEl = document.getElementById('relativesAlert');
  const countEl = document.getElementById('relativesCount');
  const detailsList = document.getElementById('relativesDetailsList');
  const titleEl = document.getElementById('relativesAlertTitle');

  if (!alertEl || !countEl || !detailsList) return;

  if (!relatives || relatives.length === 0) {
    alertEl.style.display = 'none';
    detailsList.innerHTML = '';
    return;
  }

  alertEl.style.display = 'block';
  countEl.innerText = `${relatives.length} переводов`;
  if (titleEl) {
    titleEl.innerText = `Отделены регулярные переводы родственникам / СБП (${relatives.length})`;
  }

  detailsList.innerHTML = relatives
    .map((r) => {
      const reasonsList = r.classification.reasons.join(', ');
      return `
      <div style="display: flex; justify-content: space-between; align-items: flex-start; padding: 8px 0; border-bottom: 1px solid rgba(255,255,255,0.06); gap: 12px;">
        <div>
          <div style="font-weight: 600; color: var(--text); font-size: 0.92rem;">
            <i class="fa-solid fa-user-check" style="color: var(--accent); margin-right: 6px;"></i> ${r.serviceTitle}
          </div>
          <div style="font-size: 0.78rem; color: var(--text-muted); margin-top: 3px;">
            ${reasonsList} • ${r.chargeCount} операций за период
          </div>
        </div>
        <div style="text-align: right; white-space: nowrap;">
          <div style="font-weight: 600; color: var(--text);">${r.currentPrice.toLocaleString('ru-RU')} ₽</div>
          <div style="font-size: 0.72rem; color: #21a038; font-weight: 500;">Близкие / P2P (не отменять)</div>
        </div>
      </div>
    `;
    })
    .join('');
}

function toggleRelativesDetails() {
  const detailsList = document.getElementById('relativesDetailsList');
  const chevron = document.getElementById('relativesChevron');
  if (!detailsList) return;
  const isShown = detailsList.style.display === 'block';
  detailsList.style.display = isShown ? 'none' : 'block';
  if (chevron) {
    chevron.className = isShown ? 'fa-solid fa-chevron-down' : 'fa-solid fa-chevron-up';
  }
}

async function loadSampleXlsx() {
  try {
    const wb = XLSX.utils.book_new();
    const wsData = [
      ['ПАО СБЕРБАНК — Отчет по счету дебетовой карты за 6 месяцев'],
      ['Период: 01.10.2023 — 31.03.2024'],
      ['Владелец счета: Иванов Иван Иванович'],
      [],
      ['Дата операции', 'Сумма в валюте операции', 'Категория', 'Описание операции', 'МСС'],
      // 1. Регулярные переводы маме (P2P - надежно отделяются от подписок!)
      ['10.01.2024', -15000.00, 'Переводы', 'Перевод клиенту Сбербанка: Мария Ивановна П. (маме)', '6536'],
      ['10.02.2024', -15000.00, 'Переводы', 'Перевод клиенту Сбербанка: Мария Ивановна П. (маме)', '6536'],
      ['10.03.2024', -15000.00, 'Переводы', 'Перевод клиенту Сбербанка: Мария Ивановна П. (маме)', '6536'],
      // 2. Регулярные переводы сыну (СБП - надежно отделяются от подписок!)
      ['15.01.2024', -5000.00, 'Переводы', 'Перевод через СБП: Иван Сергеевич К. (сыну на карманные)', '4829'],
      ['15.02.2024', -5000.00, 'Переводы', 'Перевод через СБП: Иван Сергеевич К. (сыну на карманные)', '4829'],
      ['15.03.2024', -5000.00, 'Переводы', 'Перевод через СБП: Иван Сергеевич К. (сыну на карманные)', '4829'],
      // 3. Подписка Яндекс Плюс с пробным периодом 1 ₽ и сменой тарифа 299 ₽ -> 399 ₽
      ['05.01.2024', -1.00, 'Развлечения', 'YNDX*PLUS KINOPOISK MOSCOW RUS', '4899'],
      ['05.02.2024', -299.00, 'Развлечения', 'Yandex Plus', '4899'],
      ['05.03.2024', -399.00, 'Развлечения', 'YNDX PLUS', '4899'],
      // 4. Подписка Telegram Premium
      ['14.01.2024', -299.00, 'Связь', 'Telegram Premium FZCO', '5816'],
      ['14.02.2024', -299.00, 'Связь', 'Telegram Premium', '5816'],
      ['14.03.2024', -299.00, 'Связь', 'Telegram Premium', '5816'],
      // 5. Книги ЛитРес
      ['18.01.2024', -399.00, 'Книги', 'LITRES.RU Литрес книги', '5815'],
      ['18.02.2024', -399.00, 'Книги', 'LITRES.RU', '5815'],
      ['18.03.2024', -399.00, 'Книги', 'LITRES.RU', '5815'],
      // 6. Плата за СМС-уведомления Сбербанка
      ['10.01.2024', -99.00, 'Услуги банка', 'Плата за уведомления об операциях', ''],
      ['10.02.2024', -99.00, 'Услуги банка', 'Плата за уведомления об операциях', ''],
      ['10.03.2024', -99.00, 'Услуги банка', 'Плата за уведомления об операциях', ''],
      // 7. Кино Okko
      ['22.01.2024', -499.00, 'Кино', 'OKKO.TV Онлайн-кинотеатр Okko', '4899'],
      ['22.02.2024', -499.00, 'Кино', 'OKKO.TV Онлайн-кинотеатр Okko', '4899'],
      // Разовые бытовые операции
      ['15.01.2024', -1250.00, 'Супермаркеты', 'Пятерочка 412', '5411'],
      ['28.02.2024', -2400.00, 'Транспорт', 'АЗС Лукойл 77', '5541']
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Выписка');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], 'vypiska_sberbank_sample.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await startRealAnalysis(file);
  } catch (err: any) {
    showToast(`Ошибка генерации тестового файла: ${err.message}`, 'warning');
  }
}

async function loadSampleCsv() {
  try {
    const csvText = `Дата операции;Сумма;Категория;Описание операции;МСС
10.01.2024;-15000,00;Переводы;Перевод клиенту Сбербанка: Мария Ивановна П. (маме);6536
10.02.2024;-15000,00;Переводы;Перевод клиенту Сбербанка: Мария Ивановна П. (маме);6536
10.03.2024;-15000,00;Переводы;Перевод клиенту Сбербанка: Мария Ивановна П. (маме);6536
15.01.2024;-5000,00;Переводы;Перевод через СБП: Иван Сергеевич К. (сыну);4829
15.02.2024;-5000,00;Переводы;Перевод через СБП: Иван Сергеевич К. (сыну);4829
15.03.2024;-5000,00;Переводы;Перевод через СБП: Иван Сергеевич К. (сыну);4829
05.01.2024;-1,00;Развлечения;YNDX*PLUS KINOPOISK;4899
05.02.2024;-299,00;Развлечения;Yandex Plus;4899
05.03.2024;-399,00;Развлечения;YNDX PLUS;4899
14.01.2024;-299,00;Связь;Telegram Premium;5816
14.02.2024;-299,00;Связь;Telegram Premium;5816
14.03.2024;-299,00;Связь;Telegram Premium;5816
10.01.2024;-99,00;Услуги банка;Плата за уведомления об операциях;
10.02.2024;-99,00;Услуги банка;Плата за уведомления об операциях;
10.03.2024;-99,00;Услуги банка;Плата за уведомления об операциях;
18.01.2024;-399,00;Книги;LITRES.RU;5815
18.02.2024;-399,00;Книги;LITRES.RU;5815
18.03.2024;-399,00;Книги;LITRES.RU;5815
20.01.2024;-499,00;Кино;OKKO;4899
20.02.2024;-499,00;Кино;OKKO;4899
15.01.2024;-1450,00;Супермаркеты;Супермаркет Перекресток;5411
`;
    const blob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });
    const file = new File([blob], 'vypiska_tbank_sample.csv', { type: 'text/csv' });
    await startRealAnalysis(file);
  } catch (err: any) {
    showToast(`Ошибка создания тестового CSV: ${err.message}`, 'warning');
  }
}

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      const base64Index = result.indexOf(';base64,');
      if (base64Index !== -1) {
        resolve(result.substring(base64Index + 8));
      } else {
        resolve(result);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsDataURL(file);
  });
}

// ===== CLIENT-SIDE STATEMENT ANALYSIS (SheetJS) =====
async function startRealAnalysis(file: File) {
  lastUploadedFile = file;
  if (uploadContent) uploadContent.style.display = 'none';
  if (loader) loader.style.display = 'flex';
  if (emptyState) emptyState.style.display = 'none';
  if (subsList) subsList.style.display = 'none';
  if (totalAmount) totalAmount.style.display = 'none';
  if (filtersBar) filtersBar.style.display = 'none';
  if (exportRow) exportRow.style.display = 'none';
  if (healthCard) healthCard.style.display = 'none';
  if (selectedSavings) selectedSavings.style.display = 'none';
  selectedSubs.clear();

  const fileNameLower = file.name.toLowerCase();
  const isTableFile =
    fileNameLower.endsWith('.xlsx') ||
    fileNameLower.endsWith('.xls') ||
    fileNameLower.endsWith('.csv') ||
    fileNameLower.endsWith('.txt') ||
    fileNameLower.endsWith('.tsv') ||
    file.type.includes('sheet') ||
    file.type.includes('excel') ||
    file.type.includes('csv');

  if (isTableFile) {
    // 100% CLIENT-SIDE PROCESSING VIA SHEETJS — NO BACKEND REQUIRED
    try {
      if (loaderText) loaderText.innerText = `Чтение таблицы через SheetJS (клиент)...`;
      if (progressFill) progressFill.style.width = '30%';

      await new Promise((r) => setTimeout(r, 120));

      if (loaderText) loaderText.innerText = `Распознавание банка и извлечение операций...`;
      if (progressFill) progressFill.style.width = '65%';

      // Run client-side SheetJS parser
      const result = await parseStatementFileClient(file);

      await new Promise((r) => setTimeout(r, 150));

      if (loaderText) loaderText.innerText = `Кластеризация подписок завершена!`;
      if (progressFill) progressFill.style.width = '100%';

      setTimeout(() => {
        if (loader) loader.style.display = 'none';
        if (progressFill) progressFill.style.width = '0%';

        const foundSubs: Subscription[] = result.subscriptions || [];
        currentRelativesTransfers = result.relativesTransfers || [];
        renderRelativesAlert(currentRelativesTransfers);

        if (foundSubs.length === 0) {
          activeData = [];
          if (uploadContent) {
            uploadContent.style.display = 'block';
            const h4 = uploadContent.querySelector('h4');
            if (h4) h4.innerText = 'Загрузить другой файл';
          }
          if (emptyState) {
            emptyState.style.display = 'block';
            const relMsg = currentRelativesTransfers.length > 0 
              ? `<p style="color: #21a038; margin-top: 6px;"><i class="fa-solid fa-user-shield"></i> Обнаружено ${currentRelativesTransfers.length} регулярных переводов родственникам (они исключены из отписок).</p>` 
              : '';
            emptyState.innerHTML = `
              <i class="fa-solid fa-file-circle-check" style="font-size: 3rem; margin-bottom: 15px; color: var(--accent);"></i>
              <h4 style="margin-bottom: 8px;">Файл "${file.name}" прочитан через SheetJS</h4>
              <p style="color: var(--text-muted); max-width: 450px; margin: 0 auto 15px;">Проанализировано ${result.statementSummary.totalTransactionsAnalyzed} операций. Регулярных платных подписок в выписке не обнаружено.</p>
              ${relMsg}
              <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; margin-top: 15px;">
                <button class="btn btn-primary btn-sm" onclick="window.loadSampleXlsx()"><i class="fa-solid fa-file-excel"></i> Проверить тестовый XLSX</button>
                <button class="btn btn-outline btn-sm" onclick="loadDemoData()"><i class="fa-solid fa-wand-magic-sparkles"></i> Демо-данные</button>
              </div>
            `;
          }
          showToast('В выписке не найдено периодических подписок', 'info');
          return;
        }

        activeData = foundSubs;
        selectedSubs.clear();
        renderResults();
        checkBudget();
        updateCharts();
        updateHealthScore();
        isAnalyzed = true;

        if (uploadContent) {
          uploadContent.style.display = 'block';
          const h4 = uploadContent.querySelector('h4');
          if (h4) h4.innerText = `Загрузить другую выписку (${file.name})`;
        }

        const bankName = result.statementSummary?.bankDetected ? ` (${result.statementSummary.bankDetected})` : '';
        showToast(`SheetJS успешно нашёл ${activeData.length} подписок${bankName} без бэкенда!`, 'success');

        const scannerSection = document.getElementById('scanner');
        if (scannerSection) {
          scannerSection.scrollIntoView({ behavior: 'smooth' });
        }
      }, 250);

      return;
    } catch (parseError: any) {
      console.error('Client-side SheetJS parsing error:', parseError);
      if (loader) loader.style.display = 'none';
      if (progressFill) progressFill.style.width = '0%';
      if (uploadContent) uploadContent.style.display = 'block';

      showToast(`Ошибка парсинга файла: ${parseError.message}`, 'warning');
      if (emptyState) {
        emptyState.style.display = 'block';
        emptyState.innerHTML = `
          <i class="fa-solid fa-triangle-exclamation" style="font-size: 3rem; margin-bottom: 15px; color: #f59e0b;"></i>
          <h4 style="margin-bottom: 8px;">Не удалось обработать файл "${file.name}"</h4>
          <p style="color: var(--text-muted); max-width: 480px; margin: 0 auto 20px;">${parseError.message || 'Ошибка чтения формата таблицы'}</p>
          <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
            <button class="btn btn-primary btn-sm" onclick="window.loadSampleXlsx()"><i class="fa-solid fa-file-excel"></i> Запустить тестовый XLSX</button>
            <button class="btn btn-outline btn-sm" onclick="loadDemoData()"><i class="fa-solid fa-wand-magic-sparkles"></i> Демо-данные</button>
          </div>
        `;
      }
      return;
    }
  }

  // Non-spreadsheet file (PDF or images)
  let progress = 10;
  if (progressFill) progressFill.style.width = '10%';
  if (loaderText) loaderText.innerText = `Чтение файла ${file.name}...`;

  const progressInterval = setInterval(() => {
    if (progress < 85) {
      progress += Math.floor(Math.random() * 8) + 4;
      if (progressFill) progressFill.style.width = `${progress}%`;
      if (progress > 30 && progress < 60) {
        loaderText.innerText = 'Распознавание структуры документа...';
      } else if (progress >= 60) {
        loaderText.innerText = 'Поиск регулярных списаний...';
      }
    }
  }, 400);

  try {
    const base64 = await readFileAsBase64(file);

    const response = await fetch('/api/analyze-statement', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filename: file.name,
        mimeType: file.type,
        base64,
        text: '',
      }),
    });

    clearInterval(progressInterval);

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}));
      throw new Error(errData.error || `Ошибка сервера: ${response.status}`);
    }

    const data = await response.json();
    if (progressFill) progressFill.style.width = '100%';
    if (loaderText) loaderText.innerText = 'Обработка завершена!';

    setTimeout(() => {
      if (loader) loader.style.display = 'none';
      if (progressFill) progressFill.style.width = '0%';

      const foundSubs: Subscription[] = data.subscriptions || [];
      if (foundSubs.length === 0) {
        activeData = [];
        if (uploadContent) uploadContent.style.display = 'block';
        if (emptyState) {
          emptyState.style.display = 'block';
          emptyState.innerHTML = `
            <i class="fa-solid fa-file-circle-check" style="font-size: 3rem; margin-bottom: 15px; color: var(--accent);"></i>
            <h4 style="margin-bottom: 8px;">Файл "${file.name}" обработан</h4>
            <p style="color: var(--text-muted); max-width: 450px; margin: 0 auto 15px;">Регулярных платных подписок в предоставленной выписке не обнаружено.</p>
            <button class="btn btn-outline btn-sm" onclick="loadDemoData()"><i class="fa-solid fa-wand-magic-sparkles"></i> Посмотреть демо-данные</button>
          `;
        }
        showToast('В выписке не найдено периодических подписок', 'info');
        return;
      }

      activeData = foundSubs;
      selectedSubs.clear();
      renderResults();
      checkBudget();
      updateCharts();
      updateHealthScore();
      isAnalyzed = true;

      if (uploadContent) {
        uploadContent.style.display = 'block';
        const h4 = uploadContent.querySelector('h4');
        if (h4) h4.innerText = `Загрузить другую выписку (${file.name})`;
      }

      const bankName = data.statementSummary?.bankDetected ? ` (${data.statementSummary.bankDetected})` : '';
      showToast(`Успешно распознано ${activeData.length} подписок${bankName}!`, 'success');
      document.getElementById('scanner')?.scrollIntoView({ behavior: 'smooth' });
    }, 300);

  } catch (error: any) {
    clearInterval(progressInterval);
    if (loader) loader.style.display = 'none';
    if (progressFill) progressFill.style.width = '0%';
    if (uploadContent) uploadContent.style.display = 'block';

    console.warn('Backend statement analysis fallback error:', error);
    const friendlyNotice = 'Для мгновенной работы прямо в браузере без бэкенда рекомендуем сохранить выписку из банка в формате Excel (.xlsx) или CSV.';
    showToast(friendlyNotice, 'info');

    if (emptyState) {
      emptyState.style.display = 'block';
      emptyState.innerHTML = `
        <i class="fa-solid fa-file-excel" style="font-size: 3rem; margin-bottom: 15px; color: #21a038;"></i>
        <h4 style="margin-bottom: 8px;">Используйте выгрузку XLSX или CSV</h4>
        <p style="color: var(--text-muted); max-width: 500px; margin: 0 auto 20px;">${friendlyNotice} Она парсится библиотекой SheetJS моментально, локально и со 100% приватностью.</p>
        <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
          <button class="btn btn-primary btn-sm" onclick="window.loadSampleXlsx()"><i class="fa-solid fa-play"></i> Попробовать тестовый XLSX</button>
          <button class="btn btn-outline btn-sm" onclick="loadDemoData()"><i class="fa-solid fa-wand-magic-sparkles"></i> Демо-данные</button>
        </div>
      `;
    }
  }
}

// ===== DEMO DATA =====
function loadDemoData() {
  const shuffled = [...subscriptionsPool].sort(() => 0.5 - Math.random());
  activeData = shuffled.slice(0, Math.floor(Math.random() * 4) + 4);
  currentRelativesTransfers = [];
  renderRelativesAlert([]);
  selectedSubs.clear();
  if (emptyState) emptyState.style.display = 'none';
  renderResults();
  checkBudget();
  updateCharts();
  updateHealthScore();
  isAnalyzed = true;
  showToast('Демо-данные загружены!', 'info');
  document.getElementById('scanner')?.scrollIntoView({ behavior: 'smooth' });
}

// ===== RENDER RESULTS =====
function renderResults(filter = currentFilter) {
  if (!subsList) return;
  subsList.innerHTML = '';
  let filtered = activeData.filter(sub => filter === 'all' || sub.category === filter);

  if (filtered.length === 0 && activeData.length > 0) {
    subsList.innerHTML = `<div style="text-align: center; color: var(--text-muted); padding: 30px 0;">В категории "${filter}" подписок не найдено</div>`;
  }

  filtered.forEach((sub, index) => {
    const item = document.createElement('div');
    item.className = 'sub-item';
    item.style.animationDelay = `${index * 0.08}s`;

    const isChecked = selectedSubs.has(sub.name) ? 'checked' : '';
    const trialBadge = sub.hasTrial ? `<span class="cluster-tag" style="background: rgba(33, 160, 56, 0.15); color: #21a038; border: 1px solid rgba(33, 160, 56, 0.3);">Пробный период</span>` : '';
    const priceChangeBadge = sub.priceChanged ? `<span class="cluster-tag" style="background: rgba(245, 158, 11, 0.15); color: #d97706; border: 1px solid rgba(245, 158, 11, 0.3);">Тариф менялся</span>` : '';

    item.innerHTML = `
      <div class="sub-info">
        <input type="checkbox" class="sub-checkbox" ${isChecked} data-name="${sub.name}" onchange="toggleSubSelection('${sub.name.replace(/'/g, "\\'")}', this.checked)">
        <div class="sub-icon" style="color: ${sub.color || 'var(--accent)'};"><i class="fa-solid ${sub.icon || 'fa-receipt'}"></i></div>
        <div class="sub-details">
          <h5>${sub.name} <span class="cluster-tag">AI Распознано</span> ${trialBadge} ${priceChangeBadge}</h5>
          <p>${sub.clusterInfo || 'Автоматически выявлено из выписки'}</p>
        </div>
      </div>
      <div class="sub-actions">
        <div class="sub-price">${sub.price} ₽/мес</div>
        <button class="btn-ai" onclick="openLLMModal('${sub.name.replace(/'/g, "\\'")}', ${sub.price})"><i class="fa-solid fa-wand-magic-sparkles"></i> Отписаться</button>
      </div>
    `;
    subsList.appendChild(item);
  });

  const fullTotal = activeData.reduce((acc, curr) => acc + curr.price, 0);
  if (totalAmount) {
    totalAmount.innerText = `Итого: ${fullTotal.toLocaleString('ru-RU')} ₽/мес`;
    totalAmount.style.display = 'block';
  }
  subsList.style.display = 'flex';
  if (filtersBar) filtersBar.style.display = 'flex';
  if (exportRow) exportRow.style.display = 'flex';
  updateSelectedSavings();
}

function filterSubs(category: string) {
  currentFilter = category;
  document.querySelectorAll('.filter-chip').forEach(chip => {
    const text = chip.textContent?.trim().toLowerCase();
    const target = category.toLowerCase();
    chip.classList.toggle('active', text === target || (category === 'all' && text === 'все'));
  });
  renderResults(category);
}

function toggleSubSelection(name: string, checked: boolean) {
  if (checked) selectedSubs.add(name);
  else selectedSubs.delete(name);
  updateSelectedSavings();
}

function updateSelectedSavings() {
  const selectedList = activeData.filter(sub => selectedSubs.has(sub.name));
  const sum = selectedList.reduce((acc, curr) => acc + curr.price, 0);
  if (selectedSavings && savingsAmount) {
    if (selectedList.length > 0) {
      selectedSavings.style.display = 'block';
      savingsAmount.innerText = sum.toLocaleString('ru-RU');
    } else {
      selectedSavings.style.display = 'none';
    }
  }
}

// ===== BUDGET CHECK =====
function checkBudget() {
  const limitInput = document.getElementById('budgetLimit') as HTMLInputElement;
  const limit = parseFloat(limitInput?.value) || 0;
  const total = activeData.reduce((acc, curr) => acc + curr.price, 0);
  const warningEl = document.getElementById('budgetWarning');
  if (!warningEl) return;
  warningEl.style.display = 'block';
  if (total > limit) {
    warningEl.style.color = 'var(--orange-500)';
    warningEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Превышение лимита на ${total - limit} ₽!`;
  } else {
    warningEl.style.color = 'var(--accent)';
    warningEl.innerHTML = `<i class="fa-solid fa-circle-check"></i> Расходы в пределах лимита`;
  }
}

// ===== HEALTH SCORE =====
function updateHealthScore() {
  if (activeData.length === 0 || !healthRing || !healthStatus || !healthDesc || !healthCard) return;
  const total = activeData.reduce((acc, curr) => acc + curr.price, 0);
  const limitInput = document.getElementById('budgetLimit') as HTMLInputElement;
  const limit = parseFloat(limitInput?.value) || 1500;
  const ratio = limit > 0 ? (total / limit) : 2;
  let score = 100 - Math.min(80, Math.floor(ratio * 50));
  if (score < 0) score = 0;
  healthRing.innerText = score + '%';
  healthRing.style.borderColor = score > 70 ? 'var(--accent)' : score > 40 ? 'var(--orange-500)' : '#ff4444';
  if (score > 70) {
    healthStatus.innerText = 'Отличное';
    healthDesc.innerText = 'Расходы на подписки под контролем';
  } else if (score > 40) {
    healthStatus.innerText = 'Среднее';
    healthDesc.innerText = 'Есть потенциал для экономии';
  } else {
    healthStatus.innerText = 'Требует внимания';
    healthDesc.innerText = 'Слишком много подписок, рекомендовано отписаться';
  }
  healthCard.style.display = 'block';
}

// ===== EXPORT =====
function exportReport() {
  if (activeData.length === 0) {
    showToast('Нет данных для экспорта', 'warning');
    return;
  }

  // Export to Excel XLSX using SheetJS on client
  try {
    const wb = XLSX.utils.book_new();
    const wsData = [
      ['Отчёт по регулярным подпискам и списаниям'],
      ['Дата анализа:', new Date().toLocaleDateString('ru-RU')],
      [],
      ['Название сервиса', 'Категория', 'Стоимость в месяц (руб.)', 'Стоимость в год (руб.)', 'Периодичность', 'Детали выписки'],
      ...activeData.map((s) => [
        s.name,
        s.category,
        s.price,
        s.price * 12,
        s.frequency || 'Ежемесячно',
        s.clusterInfo || '',
      ]),
      [],
      ['ИТОГО В МЕСЯЦ:', '', activeData.reduce((acc, c) => acc + c.price, 0), activeData.reduce((acc, c) => acc + c.price * 12, 0)]
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Подписки');
    XLSX.writeFile(wb, 'sber_subscriptions_report.xlsx');
    showToast('Отчёт успешно сохранён в Excel (XLSX)', 'success');
  } catch (e) {
    // Fallback JSON export
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(activeData, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute('href', dataStr);
    downloadAnchor.setAttribute('download', 'sber_subscriptions_report.json');
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showToast('Отчёт экспортирован в JSON');
  }
}

// ===== CLIENT-SIDE CANCELLATION LETTER MODAL =====
function openLLMModal(serviceName: string, monthlyPrice: number) {
  if (!modal || !emailContent || !yearlySavings) return;
  modal.classList.add('open');
  emailContent.innerHTML = '';
  emailContent.classList.add('typing-cursor');
  const savings = monthlyPrice * 12;
  yearlySavings.innerText = `${savings.toLocaleString('ru-RU')} ₽`;

  // Generate complete legal cancellation statement directly on client
  const letterText = generateClientCancellationLetter(serviceName, monthlyPrice);
  currentEmailText = letterText;
  typeOutText(letterText);
}

function typeOutText(fullText: string) {
  if (!emailContent) return;
  emailContent.textContent = '';
  let i = 0;
  const speed = 10;
  const interval = setInterval(() => {
    emailContent.textContent = fullText.substring(0, i);
    i += 3;
    if (i > fullText.length) {
      clearInterval(interval);
      emailContent.textContent = fullText;
      emailContent.classList.remove('typing-cursor');
    }
  }, speed);
}

function copyEmail() {
  navigator.clipboard.writeText(currentEmailText).then(() => {
    showToast('Заявление скопировано в буфер обмена');
  }).catch(() => {
    showToast('Не удалось скопировать', 'warning');
  });
}

function generateBulkLetters() {
  const selectedList = activeData.filter((sub) => selectedSubs.has(sub.name));
  if (selectedList.length === 0) {
    showToast('Выберите хотя бы одну подписку галочкой', 'warning');
    return;
  }
  const totalMonthly = selectedList.reduce((acc, curr) => acc + curr.price, 0);
  const totalYearly = totalMonthly * 12;
  if (!modal || !emailContent || !yearlySavings) return;

  modal.classList.add('open');
  emailContent.classList.add('typing-cursor');
  yearlySavings.innerText = `${totalYearly.toLocaleString('ru-RU')} ₽`;

  // Generate bulk legal statement directly on client
  const letterText = generateClientBulkCancellationLetter(selectedList);
  currentEmailText = letterText;
  typeOutText(letterText);
}

if (closeModal) {
  closeModal.onclick = () => modal?.classList.remove('open');
}
window.onclick = (e) => {
  if (e.target === modal) modal?.classList.remove('open');
};

// ===== CHARTS =====
function updateCharts() {
  if (typeof Chart === 'undefined') return;

  // 1. Category chart
  const categories: { [key: string]: number } = {};
  activeData.forEach(sub => {
    categories[sub.category] = (categories[sub.category] || 0) + sub.price;
  });
  const catLabels = Object.keys(categories);
  const catValues = Object.values(categories);
  const catColors = ['#21a038', '#f97316', '#0077FF', '#ffcc00', '#9de3a8', '#ff8c42', '#a855f7'];

  if (categoryChartInstance) categoryChartInstance.destroy();
  const canvas1 = document.getElementById('categoryChart') as HTMLCanvasElement;
  if (canvas1) {
    const ctx1 = canvas1.getContext('2d');
    if (ctx1) {
      categoryChartInstance = new Chart(ctx1, {
        type: 'pie',
        data: {
          labels: catLabels,
          datasets: [{
            data: catValues,
            backgroundColor: catColors.slice(0, catLabels.length),
            borderWidth: 2,
            borderColor: document.body.classList.contains('light-theme') ? '#fff' : '#041208',
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'bottom',
              labels: {
                color: getComputedStyle(document.body).getPropertyValue('--text'),
                font: { size: 12 },
              },
            },
            tooltip: {
              callbacks: {
                label: (ctx: any) => `${ctx.label}: ${ctx.raw} ₽/мес`,
              },
            },
          },
        },
      });
    }
  }

  // 2. Top subscriptions chart
  const topSubs = [...activeData].sort((a, b) => b.price - a.price).slice(0, 5);
  const subNames = topSubs.map(s => s.name);
  const subPrices = topSubs.map(s => s.price);
  const subColors = ['#42b856', '#21a038', '#176b34', '#114d26', '#0c331a'];

  if (topSubsChartInstance) topSubsChartInstance.destroy();
  const canvas2 = document.getElementById('topSubsChart') as HTMLCanvasElement;
  if (canvas2) {
    const ctx2 = canvas2.getContext('2d');
    if (ctx2) {
      topSubsChartInstance = new Chart(ctx2, {
        type: 'bar',
        data: {
          labels: subNames,
          datasets: [{
            label: 'Цена (₽/мес)',
            data: subPrices,
            backgroundColor: subColors,
            borderRadius: 8,
            borderSkipped: false,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { color: getComputedStyle(document.body).getPropertyValue('--text') },
              grid: { color: 'rgba(255,255,255,0.1)' },
            },
            x: {
              ticks: {
                color: getComputedStyle(document.body).getPropertyValue('--text'),
                maxRotation: 45,
                minRotation: 0,
              },
              grid: { display: false },
            },
          },
        },
      });
    }
  }
}

// Initial hide of dynamic cards
if (healthCard) healthCard.style.display = 'none';
if (selectedSavings) selectedSavings.style.display = 'none';
