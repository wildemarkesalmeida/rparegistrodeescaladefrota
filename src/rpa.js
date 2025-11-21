require('dotenv').config();
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { chromium } = require('playwright');

const BASE_URL = process.env.RPA_BASE_URL || 'http://transul.snk.ativy.com:40150/mge/system.jsp#app/YnIuY29tLnNhbmtoeWEubWVudS5hZGljaW9uYWwuQURfVEdQRVND';
const USERNAME = process.env.RPA_USERNAME;
const PASSWORD = process.env.RPA_PASSWORD;
const KEEP_BROWSER_OPEN = process.env.RPA_KEEP_BROWSER_OPEN === 'true';
const SCHEDULE_FILE = process.env.RPA_SCHEDULE_FILE || 'ESCALA DIARIA.xlsx';
const SCHEDULE_SHEET = process.env.RPA_SCHEDULE_SHEET;

function ensureCredentials() {
  if (!USERNAME || !PASSWORD) {
    throw new Error('Credenciais ausentes. Configure RPA_USERNAME e RPA_PASSWORD antes de executar o RPA.');
  }
}

function formatDateToBR(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear());
  return `${day}/${month}/${year}`;
}

function formatDateToSheetName(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear());
  return `${day}-${month}-${year}`;
}

function normalizeTurno(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return '';
  if (normalized.startsWith('d')) return 'Dia';
  if (normalized.startsWith('n')) return 'Noite';
  return String(value || '').trim();
}

function normalizeTipo(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return '';
  if (normalized.startsWith('f')) return 'Fichado';
  if (normalized.startsWith('d')) return 'Diarista';
  return String(value || '').trim();
}

function escapeForRegex(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function resolveSchedulePath() {
  if (path.isAbsolute(SCHEDULE_FILE)) {
    return SCHEDULE_FILE;
  }
  return path.resolve(process.cwd(), SCHEDULE_FILE);
}

function parseDateFromSheetName(sheetName) {
  if (!sheetName) return null;
  const match = sheetName.match(/(\d{2})[-_/](\d{2})[-_/](\d{4})/);
  if (!match) return null;
  const [, day, month, year] = match;
  const isoString = `${year}-${month}-${day}T00:00:00`;
  const parsed = new Date(isoString);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

async function loadScheduleWorkbook() {
  const schedulePath = resolveSchedulePath();
  if (!fs.existsSync(schedulePath)) {
    throw new Error(`Planilha de escala não encontrada em ${schedulePath}. Configure RPA_SCHEDULE_FILE corretamente.`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(schedulePath);
  return { workbook, schedulePath };
}

function determineStartSheetIndex(workbook) {
  if (SCHEDULE_SHEET) {
    const preferred = workbook.getWorksheet(SCHEDULE_SHEET);
    if (preferred) {
      return workbook.worksheets.indexOf(preferred);
    }
    console.warn(`Aba configurada em RPA_SCHEDULE_SHEET ("${SCHEDULE_SHEET}") não encontrada. Iniciando da primeira aba disponível.`);
  }
  const firstDateSheet = workbook.worksheets.find(ws => parseDateFromSheetName(ws.name));
  if (firstDateSheet) {
    return workbook.worksheets.indexOf(firstDateSheet);
  }
  return 0;
}

function buildWorksheetQueue(workbook, startIndex, referenceDate) {
  if (!workbook.worksheets.length) {
    return [];
  }

  const normalizedIndex = Math.max(0, Math.min(startIndex, workbook.worksheets.length - 1));
  const queue = [];
  const firstWorksheet = workbook.worksheets[normalizedIndex];
  queue.push({
    worksheet: firstWorksheet,
    worksheetName: firstWorksheet.name,
    scheduleDate: parseDateFromSheetName(firstWorksheet.name) || referenceDate
  });

  for (let i = normalizedIndex + 1; i < workbook.worksheets.length; i++) {
    const worksheet = workbook.worksheets[i];
    const parsedDate = parseDateFromSheetName(worksheet.name);
    if (parsedDate) {
      queue.push({
        worksheet,
        worksheetName: worksheet.name,
        scheduleDate: parsedDate
      });
    }
  }

  return queue;
}

function extractRecordsFromWorksheet(worksheet) {
  const headerRow = worksheet.getRow(1);
  const headerMap = {};
  headerRow.eachCell((cell, colNumber) => {
    const key = String(cell.text || cell.value || '').trim().toUpperCase();
    if (key) {
      headerMap[key] = colNumber;
    }
  });

  const requiredHeaders = ['PLACA', 'MOTORISTA', 'TURNO', 'TIPO'];
  const missingHeaders = requiredHeaders.filter(header => !headerMap[header]);
  if (missingHeaders.length) {
    throw new Error(`Planilha de escala está sem as colunas obrigatórias: ${missingHeaders.join(', ')}`);
  }

  const records = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const placa = row.getCell(headerMap.PLACA).text?.trim() || row.getCell(headerMap.PLACA).value?.toString().trim() || '';
    const motorista = row.getCell(headerMap.MOTORISTA).text?.trim() || row.getCell(headerMap.MOTORISTA).value?.toString().trim() || '';
    const rawTurno = row.getCell(headerMap.TURNO).text?.trim() || row.getCell(headerMap.TURNO).value?.toString().trim() || '';
    const rawTipo = row.getCell(headerMap.TIPO).text?.trim() || row.getCell(headerMap.TIPO).value?.toString().trim() || '';
    const turno = normalizeTurno(rawTurno);
    const tipo = normalizeTipo(rawTipo);

    if (placa && motorista) {
      records.push({ placa, motorista, turno, tipo });
    }
  });

  if (!records.length) {
    throw new Error(`Nenhum registro encontrado na aba ${worksheet.name}.`);
  }

  return records;
}

async function selectOptionFromCombo(page, frame, fieldName, value) {
  if (!value) return;
  const combo = frame.locator(`sk-combobox[sk-field-name="${fieldName}"]`);
  await combo.waitFor({ state: 'visible', timeout: 10000 });
  const textbox = combo.getByRole('textbox', { name: 'Select box' }).first();
  await textbox.click();
  await page.waitForTimeout(150);
  const option = frame.locator('.ui-select-choices-row div').filter({
    hasText: new RegExp(escapeForRegex(value), 'i')
  }).first();
  if (await option.count()) {
    await option.click();
  } else {
    if (value.length) {
      await page.keyboard.type(value);
    }
    await page.keyboard.press('Enter');
  }
}

async function selectOptionFromPesquisa(frame, fieldName, value) {
  if (!value) return;
  const container = frame.locator(`sk-pesquisa-input[sk-field-name="${fieldName}"]`);
  await container.scrollIntoViewIfNeeded();
  const descriptionInput = container.locator('sk-typeahead-input input');
  const codeInput = container.locator('sk-text-input input');

  if (await codeInput.count()) {
    await codeInput.fill('');
  }

  await descriptionInput.fill('');
  await descriptionInput.type(value, { delay: 50 });

  const suggestion = frame.locator('.ui-select-choices-row').filter({ hasText: value }).first();
  await suggestion.waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
  if (await suggestion.count()) {
    await suggestion.click();
  } else {
    await descriptionInput.press('Enter');
  }
}

async function setTurno(frame, turno) {
  const turnoCombo = frame.locator('sk-combobox[sk-field-name="TURNO"]').getByRole('textbox', { name: 'Select box' }).first();
  await turnoCombo.click();
  const option = frame.locator('.ui-select-choices-row.ng-scope .option').filter({ hasText: turno }).first();
  await option.click();
}

async function setMotorista(frame, motorista) {
  const motoristaInput = frame.locator('sk-pesquisa-input[sk-field-name="CODMOT"]').locator('sk-typeahead-input input');
  await motoristaInput.fill('');
  await motoristaInput.type(motorista, { delay: 50 });
  const suggestion = frame.locator('.ui-select-choices-row')
    .filter({ hasText: new RegExp(escapeForRegex(motorista), 'i') })
    .first();
  try {
    await suggestion.waitFor({ state: 'visible', timeout: 5000 });
    await suggestion.click();
  } catch {
    await motoristaInput.press('Tab');
  }
}

async function setVeiculo(frame, placa) {
  const veiculoInput = frame.locator('sk-pesquisa-input[sk-field-name="CODVEICULO"]').locator('sk-typeahead-input input');
  await veiculoInput.fill('');
  await veiculoInput.type(placa, { delay: 50 });
  const suggestion = frame.locator('.ui-select-choices-row')
    .filter({ hasText: new RegExp(escapeForRegex(placa), 'i') })
    .first();
  try {
    await suggestion.waitFor({ state: 'visible', timeout: 5000 });
    await suggestion.click();
  } catch {
    await veiculoInput.press('Tab');
  }
}

async function setTipo(frame, tipo) {
  const tipoCombo = frame.locator('sk-combobox[sk-field-name="TIPO"]').getByRole('textbox', { name: 'Select box' }).first();
  await tipoCombo.click();
  const option = frame.locator('.ui-select-choices-row.ng-scope .option').filter({ hasText: tipo }).first();
  await option.click();
}

async function clickGlyphButton(frame, glyph) {
  const button = frame.getByRole('button', { name: glyph });
  await button.waitFor({ state: 'visible', timeout: 10000 });
  await button.click();
}

async function confirmEntry(frame) {
  const confirmButton = frame.locator('#dynaform-content-002').getByRole('button', { name: '' });
  await confirmButton.waitFor({ state: 'visible', timeout: 10000 });
  await confirmButton.click();
}

async function waitForEnabledButton(frame, locator, timeout = 60000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    try {
      await locator.waitFor({ state: 'visible', timeout: 2000 });
      if (await locator.isEnabled()) {
        return;
      }
    } catch {
      // elemento pode ter sido recriado; tenta novamente
    }
    await frame.waitForTimeout(300);
  }
  throw new Error('Botão "Salvar" não ficou habilitado dentro do tempo limite.');
}

async function saveEntry(frame) {
  const saveButton = frame
    .locator('button[ng-click="save()"]')
    .filter({ hasNotText: /\bFiltro\b/i })
    .first();
  try {
    await waitForEnabledButton(frame, saveButton);
    await saveButton.click();
    await frame.waitForTimeout(300);
  } catch (error) {
    console.warn('Botão "Salvar" não habilitou a tempo. Tentando atalho F7.');
    await frame.press('body', 'F7');
  }
}

async function setEntryDate(frame, date) {
  const dateInput = frame.locator('input[sk-bind-pop-over-id="00P"]');
  await dateInput.waitFor({ state: 'visible', timeout: 10000 });
  await dateInput.fill('');
  await dateInput.fill(formatDateToBR(date));
  await dateInput.press('Enter');
}

async function acknowledgeDialogs(page) {
  const okButton = page.getByRole('button', { name: /^Ok$/i });
  if (await okButton.count()) {
    await okButton.click();
  }
}

async function waitForOverlayToHide(frame) {
  await frame.waitForTimeout(500);
  try {
    await frame.waitForSelector('.SimplePopupGlass', { state: 'hidden', timeout: 5000 });
  } catch {
    await frame.waitForTimeout(300);
  }
}

async function clickNovoRegistro(frame) {
  const selectors = [
    'button.btn-novo-registro',
    'sk-button:has-text("Novo registro")',
    'button:has-text("Novo registro")',
    'button:has-text("Novo")',
    'span:has-text("Novo registro")',
    'text=/Novo registro/i',
    'button[title*="Novo"]',
    'sk-button[title*="Novo"] button',
    'sk-button[aria-label*="Novo"] button'
  ];
  for (const selector of selectors) {
    const locator = frame.locator(selector).first();
    try {
      await locator.waitFor({ state: 'visible', timeout: 2000 });
      await locator.click();
      return true;
    } catch {
      // tenta próximo seletor
    }
  }
  return false;
}

async function startNewRegister(frame) {
  await waitForOverlayToHide(frame);
  const clicked = await clickNovoRegistro(frame);
  if (clicked) return;
  try {
    await clickGlyphButton(frame, '');
  } catch (error) {
    const visibleButtons = await frame
      .locator('button, sk-button')
      .first()
      .evaluate(el => ({
        text: el.innerText?.trim(),
        title: el.title,
        classes: el.className
      }))
      .catch(() => null);
    console.error('Botões visíveis antes do erro:', visibleButtons);
    throw new Error('Botão inicial "Novo registro" não encontrado.');
  }
}

async function prepareScheduleForDate(frame, page, scheduleDate) {
  await startNewRegister(frame);
  await setEntryDate(frame, scheduleDate);
  await clickGlyphButton(frame, '');
  await confirmEntry(frame);
  await saveEntry(frame);
  await acknowledgeDialogs(page);
}

async function run() {
  ensureCredentials();

  const referenceDate = new Date();
  const { workbook, schedulePath } = await loadScheduleWorkbook();
  const startIndex = determineStartSheetIndex(workbook);
  const sheetQueue = buildWorksheetQueue(workbook, startIndex, referenceDate);
  if (!sheetQueue.length) {
    throw new Error('Não foi possível identificar abas para processamento na planilha.');
  }

  console.log(`Planilha carregada: ${schedulePath}`);
  console.log(`Abas a processar: ${sheetQueue.map(sheet => sheet.worksheetName).join(', ')}`);

  const browser = await chromium.launch({
    headless: false,
    channel: 'chrome', // garante uso do Chrome instalado
    args: ['--start-maximized']
  });

  const context = await browser.newContext({ viewport: null });
  const page = await context.newPage();

  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });

  const usernameInput = page.getByRole('textbox', { name: 'Usuário' });
  await usernameInput.waitFor({ state: 'visible' });
  await usernameInput.fill(USERNAME);
  await page.getByRole('button', { name: 'Prosseguir' }).click();

  const passwordInput = page.getByRole('textbox', { name: 'Senha' });
  await passwordInput.waitFor({ state: 'visible' });
  await passwordInput.fill(PASSWORD);
  await page.getByRole('button', { name: 'Prosseguir' }).click();

  const searchInput = page.getByRole('textbox', { name: 'Pesquisar' });
  await searchInput.waitFor({ state: 'visible' });
  await searchInput.fill('ESCALA');

  const escalaOption = page.locator('span', { hasText: 'Escala Motoristas' }).first();
  await escalaOption.waitFor({ state: 'visible' });
  await escalaOption.click();

  const escalaFrameElement = await page.waitForSelector('iframe[title="Escala Motoristas"]', { timeout: 20000 });
  const escalaFrame = await escalaFrameElement.contentFrame();
  if (!escalaFrame) {
    throw new Error('Frame "Escala Motoristas" não carregado.');
  }

  for (const sheetInfo of sheetQueue) {
    const { worksheet, worksheetName, scheduleDate } = sheetInfo;
    const records = extractRecordsFromWorksheet(worksheet);
    console.log(`\n--- Processando aba ${worksheetName} ---`);
    console.log(`Data da escala: ${formatDateToBR(scheduleDate)}`);
    console.log(`Primeiro registro:`, records[0]);

    await prepareScheduleForDate(escalaFrame, page, scheduleDate);

    for (const record of records) {
      console.log(`Processando motorista: ${record.motorista} | Veículo: ${record.placa}`);
      const newDriverButton = escalaFrame
        .locator('button.btn-novo-registro')
        .filter({ hasText: /\bVeículos x Motoristas\b/i })
        .first();
      await newDriverButton.waitFor({ state: 'visible', timeout: 10000 });
      await newDriverButton.click();

      await setTurno(escalaFrame, record.turno);
      await setMotorista(escalaFrame, record.motorista);
      await setVeiculo(escalaFrame, record.placa);
      await setTipo(escalaFrame, record.tipo);
      await saveEntry(escalaFrame);
      await acknowledgeDialogs(page);
    }
  }

  if (KEEP_BROWSER_OPEN) {
    console.log('Fluxo concluído. Pressione Ctrl+C para fechar o navegador.');
    await new Promise(() => { /* mantém processo aberto */ });
  } else {
    await browser.close();
  }
}

run().catch(error => {
  console.error('Falha ao executar o fluxo RPA:', error);
  process.exit(1);
});

