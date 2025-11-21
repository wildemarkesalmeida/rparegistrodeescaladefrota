require('dotenv').config();
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { chromium } = require('playwright');

const BASE_URL = process.env.RPA_BASE_URL || 'http://transul.snk.ativy.com:50150/mge/system.jsp#app/YnIuY29tLnNhbmtoeWEubWVudS5hZGljaW9uYWwuQURfVEdQRVND';
const USERNAME = process.env.RPA_USERNAME;
const PASSWORD = process.env.RPA_PASSWORD;
const KEEP_BROWSER_OPEN = process.env.RPA_KEEP_BROWSER_OPEN === 'true';
const SCHEDULE_FILE = process.env.RPA_SCHEDULE_FILE || 'ESCALA DIARIA - 04-11.xlsx';
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

async function loadScheduleData(referenceDate = new Date()) {
  const schedulePath = resolveSchedulePath();
  if (!fs.existsSync(schedulePath)) {
    throw new Error(`Planilha de escala não encontrada em ${schedulePath}. Configure RPA_SCHEDULE_FILE corretamente.`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(schedulePath);

  const preferredSheetName = SCHEDULE_SHEET || formatDateToSheetName(referenceDate);
  let worksheet = workbook.getWorksheet(preferredSheetName);
  if (!worksheet) {
    worksheet = workbook.worksheets[0];
    console.warn(`Aba ${preferredSheetName} não encontrada. Utilizando primeira aba disponível: ${worksheet?.name ?? 'desconhecida'}.`);
  }
  if (!worksheet) {
    throw new Error('Planilha de escala não contém abas válidas.');
  }

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

  return { records, worksheetName: worksheet.name, schedulePath };
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

async function run() {
  ensureCredentials();

  const referenceDate = new Date();
  const { records, worksheetName, schedulePath } = await loadScheduleData(referenceDate);
  const scheduleDate =
    parseDateFromSheetName(worksheetName) ||
    parseDateFromSheetName(SCHEDULE_SHEET) ||
    referenceDate;
  const primaryRecord = records[0];
  console.log(`Planilha carregada: ${schedulePath}`);
  console.log(`Aba utilizada: ${worksheetName}`);
  console.log(`Data da escala: ${formatDateToBR(scheduleDate)}`);
  console.log(`Primeiro registro:`, primaryRecord);

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

  await escalaFrame.waitForTimeout(500);
  const novoRegistroButton =
    (await escalaFrame.$('text=/Novo registro/i')) ||
    (await escalaFrame.$('button:has-text("Novo registro")')) ||
    (await escalaFrame.$('span:has-text("Novo registro")'));
  if (!novoRegistroButton) {
    throw new Error('Botão inicial "Novo registro" não encontrado.');
  }
  await novoRegistroButton.click();
  const dateInput = escalaFrame.locator('input[sk-bind-pop-over-id="00P"]');
  await dateInput.waitFor({ state: 'visible' });
  await dateInput.fill('');
  await dateInput.fill(formatDateToBR(scheduleDate));
  await dateInput.press('Enter');

  await escalaFrame.getByRole('button', { name: '' }).click();
  await escalaFrame.locator('#dynaform-content-002').getByRole('button', { name: '' }).click();
  await escalaFrame.locator('sk-combobox[sk-field-name="TURNO"]').getByRole('textbox', { name: 'Select box' }).click();
  await escalaFrame.locator('div').filter({ hasText: primaryRecord.turno }).first().click();
  await setMotorista(escalaFrame, primaryRecord.motorista);
  await setVeiculo(escalaFrame, primaryRecord.placa);
  await escalaFrame.locator('sk-combobox[sk-field-name="TIPO"]').getByRole('textbox', { name: 'Select box' }).click();
  await escalaFrame.locator('div').filter({ hasText: primaryRecord.tipo }).first().click();

  await browser.close();
}

run().catch(error => {
  console.error('Falha ao executar o fluxo RPA:', error);
  process.exit(1);
});

