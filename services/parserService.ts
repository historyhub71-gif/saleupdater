import { ParsedSalesMessage } from '@/types/sales';

const DATE_PATTERN = /\b(\d{1,4}[./-]\d{1,2}[./-]\d{2,4})\b/;
const DSF_PATTERN = /(?:^|\n)\s*(?:d\s*\.?\s*s\s*\.?\s*f|dsf)\s*(?:name)?\s*[:=-]\s*([^\r\n]+)/i;
const TODAY_VALUE_PATTERN = /today(?:'s)?\s*(?:valunne|value|sales?|v)\s*[:=-]\s*([0-9,\s]+)/i;

export function parseSalesMessage(message: string): ParsedSalesMessage {
  // Normalize the WhatsApp text before applying the extraction rules.
  const normalizedMessage = normalizeMessage(message);

  if (!normalizedMessage) {
    throw new Error('Invalid WhatsApp message');
  }

  const dateMatch = normalizedMessage.match(DATE_PATTERN);
  const dsfMatch = normalizedMessage.match(DSF_PATTERN);
  const todayValueMatch = normalizedMessage.match(TODAY_VALUE_PATTERN);

  if (!dsfMatch?.[1] || !todayValueMatch?.[1]) {
    throw new Error('Invalid WhatsApp message');
  }

  const date = dateMatch?.[1]?.trim() || getTodayDateString();
  const dsf = cleanTextValue(dsfMatch[1]);
  const rawTodayValue = todayValueMatch[1].replace(/[^0-9]/g, '');
  const todayValue = Number(rawTodayValue);

  if (!date || !dsf || Number.isNaN(todayValue) || todayValue <= 0) {
    throw new Error('Invalid WhatsApp message');
  }

  return {
    date,
    dsf,
    todayValue,
  };
}

function getTodayDateString(): string {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();
  return `${day}.${month}.${year}`;
}

function normalizeMessage(message: string): string {
  return message
    .replace(/[\u200B-\u200D\uFEFF\u200E\u200F]/g, '') // Strip invisible direction & zero-width marks
    .replace(/[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]/g, ' ') // Convert all unicode spaces to standard spaces
    .replace(/\*/g, ' ')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n+/g, '\n')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .trim();
}

function cleanTextValue(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}
