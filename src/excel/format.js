// Turns cell values into the text a cell shows, using Excel number format codes.
import SSF from 'ssf';

export const NUMBER_FORMATS = [
  { label: 'General', fmt: 'General' },
  { label: 'Number', fmt: '#,##0.00' },
  { label: 'Currency', fmt: '$#,##0.00' },
  { label: 'Accounting', fmt: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
  { label: 'Short Date', fmt: 'm/d/yyyy' },
  { label: 'Long Date', fmt: 'dddd, mmmm d, yyyy' },
  { label: 'Time', fmt: 'h:mm:ss AM/PM' },
  { label: 'Percentage', fmt: '0.00%' },
  { label: 'Fraction', fmt: '# ?/?' },
  { label: 'Scientific', fmt: '0.00E+00' },
  { label: 'Text', fmt: '@' },
];

export function isError(value) {
  return value !== null && typeof value === 'object' && 'error' in value;
}

export function formatValue(value, numFmt = 'General') {
  if (value === null || value === undefined) return '';
  if (isError(value)) return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value;
  if (!Number.isFinite(value)) return '#NUM!';
  try {
    return SSF.format(numFmt || 'General', value);
  } catch {
    return SSF.format('General', value);
  }
}

// Text that fits `width` pixels: General numbers lose decimals first, then
// anything numeric becomes ### like Excel.
export function fitNumber(value, numFmt, width, measure) {
  const text = formatValue(value, numFmt);
  if (measure(text) <= width) return text;
  if ((!numFmt || numFmt === 'General') && Number.isFinite(value)) {
    for (let digits = 10; digits >= 1; digits--) {
      const shorter = Number(value.toPrecision(digits));
      let candidate = SSF.format('General', shorter);
      if (Math.abs(value) >= 1e11 || (Math.abs(value) < 1e-9 && value !== 0)) candidate = value.toExponential(Math.max(0, digits - 1)).replace('e', 'E');
      if (measure(candidate) <= width) return candidate;
    }
  }
  const hash = measure('#') || 7;
  return '#'.repeat(Math.max(1, Math.floor(width / hash)));
}

export function isDateFormat(numFmt) {
  if (!numFmt || numFmt === 'General') return false;
  try {
    return SSF.is_date(numFmt);
  } catch {
    return false;
  }
}

// What appears in the formula bar for a typed (non-formula) value.
export function editText(value, numFmt) {
  if (typeof value === 'number') {
    if (isDateFormat(numFmt)) {
      const hasTime = /[hs]/i.test(numFmt.replace(/"[^"]*"/g, ''));
      const hasDate = /[dy]/i.test(numFmt.replace(/"[^"]*"/g, '')) || /m{3,}/i.test(numFmt);
      if (hasDate && hasTime) return SSF.format('m/d/yyyy h:mm:ss AM/PM', value);
      if (hasTime && !hasDate) return SSF.format('h:mm:ss AM/PM', value);
      return SSF.format('m/d/yyyy', value);
    }
    if (numFmt && /%/.test(numFmt.replace(/"[^"]*"/g, ''))) return `${round(value * 100)}%`;
    return String(round(value));
  }
  return null;
}

function round(n) {
  return Number(n.toPrecision(15));
}

// Adds or removes one decimal place from a number format.
export function adjustDecimals(numFmt, value, delta) {
  let fmt = numFmt && numFmt !== 'General' ? numFmt : generalBase(value);
  const sections = fmt.split(';').map((section) => {
    let changed = false;
    return section.replace(/(#,##0|0)(\.0+)?/, (match, int, dec = '') => {
      if (changed) return match;
      changed = true;
      const places = Math.max(0, (dec ? dec.length - 1 : 0) + delta);
      return places ? `${int}.${'0'.repeat(places)}` : int;
    });
  });
  fmt = sections.join(';');
  return fmt;
}

function generalBase(value) {
  if (typeof value !== 'number') return '0';
  const text = String(round(value));
  const dot = text.indexOf('.');
  return dot < 0 || /e/i.test(text) ? '0' : `0.${'0'.repeat(text.length - dot - 1)}`;
}
