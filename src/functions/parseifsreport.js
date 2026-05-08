const { app } = require('@azure/functions');
const pdf = require('pdf-parse');

// --------------------------------------------------
// Helpers
// --------------------------------------------------
function cleanValue(value) {
    if (value === null || value === undefined) return null;

    const result = String(value)
        .replace(/\r/g, '')
        .replace(/[ \t]+/g, ' ')
        .replace(/\n{2,}/g, '\n')
        .trim();

    return result || null;
}

function normalizeLine(line) {
    return String(line || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function normalizeDate(value) {
    if (!value) return null;

    const raw = String(value).trim();

    let m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;

    m = raw.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;

    return cleanValue(value);
}

function toNumber(value) {
    if (value === null || value === undefined) return null;

    const normalized = String(value).replace(',', '.').trim();
    const num = Number(normalized);

    return Number.isFinite(num) ? num : null;
}

function splitPages(text) {
    return String(text || '')
        .split(/\f+/)
        .map(x => cleanValue(x))
        .filter(Boolean);
}

function normalizeCompactedLine(line) {
    let s = String(line || '').trim();

    s = s.replace(/(\d{2}-\d{2})(\d+[.,]\d+)$/g, '$1 $2');
    s = s.replace(/(\d{2}[-/]\d{2}[-/]\d{4})(\d+[.,]\d+)$/g, '$1 $2');

    return s;
}

function uniquePush(arr, item, keyBuilder) {
    const key = keyBuilder(item);
    if (!arr.some(x => keyBuilder(x) === key)) {
        arr.push(item);
    }
}

// --------------------------------------------------
// Labels
// --------------------------------------------------
const LABELS = {
    interventionDate: ['Intervention Date', 'Besøgs dato', 'Besøgsdato'],
    serviceType: ['Service Type', 'Servicetype'],
    technicianName: ['Technician Name', 'Teknikerens navn', 'Tekniker'],
    reference: ['Reference'],
    task: ['Task', 'Opgave'],
    customer: ['Customer', 'Kunde'],
    customerName: ['Customer Name', 'Kundenavn'],
    customerContact: ['Customer Contact', 'Kundekontakt'],
    description: ['Description', 'Beskrivelse'],
    product: ['Product', 'Produkt'],
    partsUsed: ['Parts Used', 'Brugte dele'],
    partsReturned: ['Parts Returned', 'Dele returneres'],
    labourAndExpense: ['Labour and Expenses', 'Labour and Expense', 'Arbejdskraft og Udgifter'],
    notes: ['Notes', 'Noter']
};

const ALL_SECTION_HEADINGS = [
    ...LABELS.customer,
    ...LABELS.customerName,
    ...LABELS.customerContact,
    ...LABELS.description,
    ...LABELS.product,
    ...LABELS.partsUsed,
    ...LABELS.partsReturned,
    ...LABELS.labourAndExpense,
    ...LABELS.notes
].map(normalizeLine);

// --------------------------------------------------
// Generic helpers
// --------------------------------------------------
function findLabelIndex(lines, aliases) {
    const normalizedAliases = aliases.map(normalizeLine);
    return lines.findIndex(line => normalizedAliases.includes(normalizeLine(line)));
}

function getValueAfterAnyLabel(lines, aliases) {
    const idx = findLabelIndex(lines, aliases);
    if (idx === -1) return null;
    return lines[idx + 1] || null;
}

function getSectionLines(lines, startAliases) {
    const start = findLabelIndex(lines, startAliases);
    if (start === -1) return [];

    let end = lines.length;

    for (let i = start + 1; i < lines.length; i++) {
        if (ALL_SECTION_HEADINGS.includes(normalizeLine(lines[i]))) {
            end = i;
            break;
        }
    }

    return lines.slice(start + 1, end).map(x => x.trim()).filter(Boolean);
}

function extractServiceNumberFromFileName(fileName) {
    if (!fileName) return null;

    const matches = [...String(fileName).matchAll(/\b(\d{6,})\b/g)].map(m => m[1]);

    if (matches.length >= 2) return matches[1];
    return matches[0] || null;
}

function extractServiceOrder(text, fileName) {
    const matches = [...String(text || '').matchAll(/Service Order:\s*(\d+)/gi)].map(m => m[1]);
    if (matches.length) return matches[0];

    return extractServiceNumberFromFileName(fileName);
}

function isArbejdstidVarenummer(varenummer) {
    return ['15EXP101', '15EXP201'].includes(String(varenummer || '').toUpperCase());
}

function isHeaderLine(line) {
    const n = normalizeLine(line);

    return [
        'item numberdescriptionserialquantity',
        'item numberdescriptionquantity',
        'serial numberdescriptionmodel',
        'descriptiondateamount',
        'varenummerbeskrivelseantal',
        'varenummerbeskrivelseserieantal',
        'serienummerbeskrivelsemodel',
        'beskrivelsedatobeløb',
        'item number',
        'serial number',
        'description',
        'quantity',
        'serial',
        'date',
        'amount',
        'varenummer',
        'serienummer',
        'beskrivelse',
        'antal',
        'serie',
        'dato',
        'beløb',
        'model'
    ].includes(n);
}

function isNoiseLine(line) {
    const n = normalizeLine(line);

    return [
        'signature name',
        'signature date',
        'unable to sign',
        'unwilling to sign',
        '̈'
    ].includes(n);
}

// --------------------------------------------------
// Kunde
// --------------------------------------------------
function parseKunde(lines, fileName, serviceOrder) {
    const kundeBlock = getSectionLines(lines, LABELS.customer);

    return {
        kunde: kundeBlock[0] || null,
        kundeadresse: kundeBlock.length > 1 ? kundeBlock.slice(1).join(', ') : null,
        kundenummer: getValueAfterAnyLabel(lines, LABELS.customerName),
        kundekontakt: getValueAfterAnyLabel(lines, LABELS.customerContact),
        besogsdato: normalizeDate(getValueAfterAnyLabel(lines, LABELS.interventionDate)),
        servicetype: getValueAfterAnyLabel(lines, LABELS.serviceType),
        servicenummer: serviceOrder || extractServiceNumberFromFileName(fileName),
        tekniker: getValueAfterAnyLabel(lines, LABELS.technicianName),
        reference: getValueAfterAnyLabel(lines, LABELS.reference),
        task: getValueAfterAnyLabel(lines, LABELS.task)
    };
}

// --------------------------------------------------
// Beskrivelse
// --------------------------------------------------
function parseBeskrivelse(lines) {
    const descriptionLines = getSectionLines(lines, LABELS.description);

    for (const line of descriptionLines) {
        if (!line) continue;
        if (isHeaderLine(line)) continue;
        if (isNoiseLine(line)) continue;
        if (/^service order:/i.test(line)) continue;
        if (/^\(utc/i.test(line)) continue;
        if (/^[A-ZÆØÅ.\- ]+A\/S$/i.test(line)) continue;

        return cleanValue(line);
    }

    return null;
}

// --------------------------------------------------
// Produkt
// --------------------------------------------------
function parseProdukt(lines) {
    const productLines = getSectionLines(lines, LABELS.product)
        .filter(x => !isHeaderLine(x))
        .filter(Boolean);

    if (!productLines.length) {
        return {
            serienummer: null,
            beskrivelse: null,
            model: null
        };
    }

    const line = productLines[0];

    const m = line.match(/^(\d+)\s*(.+?)([A-Z0-9_]+)$/);
    if (!m) {
        return {
            serienummer: null,
            beskrivelse: cleanValue(line),
            model: null
        };
    }

    return {
        serienummer: cleanValue(m[1]),
        beskrivelse: cleanValue(m[2]),
        model: cleanValue(m[3])
    };
}

// --------------------------------------------------
// Parts line parser
// --------------------------------------------------
function parseItemLine(rawLine) {
    const line = normalizeCompactedLine(rawLine);

    const match = line.match(/^((?:15[A-Z]{3}\d{3})|(?:\d+(?:\.\d+)+))(.*)$/i);
    if (!match) return null;

    const varenummer = match[1];
    const rest = cleanValue(match[2]);
    if (!rest) return null;

    const qtyMatch = rest.match(/(\d+(?:[.,]\d+)?)\s*$/);
    if (!qtyMatch) {
        return {
            varenummer,
            beskrivelse: rest,
            antal: null,
            rawLine
        };
    }

    const antal = toNumber(qtyMatch[1]);
    const beskrivelse = cleanValue(
        rest.substring(0, rest.length - qtyMatch[1].length).trim()
    );

    return {
        varenummer,
        beskrivelse,
        antal,
        rawLine
    };
}

// --------------------------------------------------
// Brugte dele
// --------------------------------------------------
function parseBrugteDele(lines, pageNumber) {
    const sectionLines = getSectionLines(lines, LABELS.partsUsed);

    const arbejdstid = [];
    const reservedele = [];

    for (const rawLine of sectionLines) {
        if (isHeaderLine(rawLine) || isNoiseLine(rawLine)) continue;

        const item = parseItemLine(rawLine);
        if (!item) continue;

        const obj = {
            pageNumber,
            varenummer: item.varenummer,
            beskrivelse: item.beskrivelse,
            antal: item.antal,
            rawLine: item.rawLine
        };

        if (isArbejdstidVarenummer(item.varenummer)) {
            uniquePush(arbejdstid, obj, x => `${x.pageNumber}|${x.varenummer}|${x.beskrivelse}|${x.antal}`);
        } else {
            uniquePush(reservedele, obj, x => `${x.pageNumber}|${x.varenummer}|${x.beskrivelse}|${x.antal}`);
        }
    }

    return {
        arbejdstid,
        reservedele
    };
}

// --------------------------------------------------
// Dele returneret
// --------------------------------------------------
function parseDeleReturneret(lines, pageNumber) {
    const sectionLines = getSectionLines(lines, LABELS.partsReturned);
    const results = [];

    for (const rawLine of sectionLines) {
        if (isHeaderLine(rawLine) || isNoiseLine(rawLine)) continue;
        if (/^service order:/i.test(rawLine)) continue;
        if (/^\(utc/i.test(rawLine)) continue;
        if (/^[A-ZÆØÅ.\- ]+A\/S$/i.test(rawLine)) continue;

        const item = parseItemLine(rawLine);
        if (!item) continue;

        uniquePush(
            results,
            {
                pageNumber,
                varenummer: item.varenummer,
                beskrivelse: item.beskrivelse,
                antal: item.antal,
                rawLine: item.rawLine
            },
            x => `${x.pageNumber}|${x.varenummer}|${x.beskrivelse}|${x.antal}`
        );
    }

    return {
        linjer: results
    };
}

// --------------------------------------------------
// Arbejdskraft og udgifter
// --------------------------------------------------
function parseArbejdskraftOgUdgifter(pageText, pageNumber) {
    const text = String(pageText || '');

    const startMatch = text.match(/(?:Labour and Expenses|Labour and Expense|Arbejdskraft og Udgifter)\s*/i);
    if (!startMatch) {
        return {
            linjer: [],
            raw: null
        };
    }

    const startIndex = startMatch.index + startMatch[0].length;
    let sectionText = text.substring(startIndex);

    const endMatch = sectionText.match(/(?:Parts Returned|Dele returneres|Notes|Noter|Description|Beskrivelse)\s*/i);
    if (endMatch) {
        sectionText = sectionText.substring(0, endMatch.index);
    }

    sectionText = cleanValue(sectionText);

    if (!sectionText) {
        return {
            linjer: [],
            raw: null
        };
    }

    const rawLines = sectionText
        .split('\n')
        .map(x => x.trim())
        .filter(Boolean);

    const results = [];

    for (let rawLine of rawLines) {
        if (isHeaderLine(rawLine) || isNoiseLine(rawLine)) continue;

        let line = normalizeCompactedLine(rawLine);

        const match = line.match(/^(.*?)(\d{2}[-/]\d{2}[-/]\d{4})\s*(\d+(?:[.,]\d+)?)$/);
        if (!match) continue;

        const beskrivelse = cleanValue(match[1]);
        const dato = normalizeDate(match[2]);
        const beloeb = toNumber(match[3]);

        if (!beskrivelse && !dato && beloeb === null) continue;

        uniquePush(
            results,
            {
                pageNumber,
                beskrivelse,
                dato,
                beloeb,
                rawLine
            },
            x => `${x.pageNumber}|${x.beskrivelse}|${x.dato}|${x.beloeb}`
        );
    }

    return {
        linjer: results,
        raw: sectionText
    };
}

// --------------------------------------------------
// Noter
// --------------------------------------------------
function parseNoter(lines) {
    const noteLines = getSectionLines(lines, LABELS.notes);

    const filtered = noteLines
        .map(x => x.trim())
        .filter(Boolean)
        .filter(x => !isHeaderLine(x))
        .filter(x => !isNoiseLine(x))
        .filter(x => !/^service order:/i.test(x))
        .filter(x => !/^\(utc/i.test(x))
        .filter(x => !/^[A-ZÆØÅ.\- ]+A\/S$/i.test(x))
        .filter(x => normalizeLine(x) !== 'parts returned');

    return filtered.length ? cleanValue(filtered.join('\n')) : null;
}

// --------------------------------------------------
// Valid visit page
// --------------------------------------------------
function isRealVisitPage(page) {
    return Boolean(
        page?.kunde?.reference ||
        page?.kunde?.task ||
        page?.kunde?.besogsdato ||
        page?.produkt?.serienummer ||
        page?.produkt?.model ||
        page?.brugteDele?.arbejdstid?.length ||
        page?.brugteDele?.reservedele?.length ||
        page?.arbejdskraftOgUdgifter?.linjer?.length
    );
}

// --------------------------------------------------
// Build one page
// --------------------------------------------------
function buildPage(pageText, pageNumber, fileName, serviceOrder) {
    const lines = String(pageText || '')
        .split('\n')
        .map(x => x.trim())
        .filter(Boolean);

    return {
        pageNumber,
        serviceOrder,
        kunde: parseKunde(lines, fileName, serviceOrder),
        beskrivelse: parseBeskrivelse(lines),
        produkt: parseProdukt(lines),
        brugteDele: parseBrugteDele(lines, pageNumber),
        deleReturneret: parseDeleReturneret(lines, pageNumber),
        arbejdskraftOgUdgifter: parseArbejdskraftOgUdgifter(pageText, pageNumber),
        noter: parseNoter(lines),
        rawText: cleanValue(pageText)
    };
}

// --------------------------------------------------
// Build document
// --------------------------------------------------
function buildDocument(text, numPages, fileName) {
    const serviceOrder = extractServiceOrder(text, fileName);
    const pageTexts = splitPages(text);

    const allPages = pageTexts.map((p, i) =>
        buildPage(p, i + 1, fileName, serviceOrder)
    );

    const records = allPages.filter(isRealVisitPage);

    const ignoredPages = allPages
        .filter(p => !isRealVisitPage(p))
        .map(p => ({
            pageNumber: p.pageNumber,
            reason: 'Ingen reelt besøg på siden'
        }));

    return {
        filnavn: fileName || null,
        serviceOrder,
        detectedPages: pageTexts.length,
        reportedPages: numPages,
        records,
        ignoredPages,
        pages: allPages
    };
}

// --------------------------------------------------
// Azure Function
// --------------------------------------------------
app.http('parseIfsReport', {
    methods: ['POST'],
    authLevel: 'function',
    handler: async (req, context) => {
        try {
            const body = await req.json();

            if (!body?.pdfBase64) {
                return {
                    status: 400,
                    jsonBody: {
                        ok: false,
                        error: 'pdfBase64 mangler'
                    }
                };
            }

            const buffer = Buffer.from(body.pdfBase64, 'base64');

            const data = await pdf(buffer, {
                pagerender: async function (pageData) {
                    const textContent = await pageData.getTextContent();

                    const text = textContent.items
                        .map(item => item.str)
                        .join('\n');

                    return text + '\f';
                }
            });

            const result = buildDocument(
                data.text || '',
                data.numpages || 0,
                body.fileName || null
            );

            return {
                status: 200,
                jsonBody: {
                    ok: true,
                    pages: data.numpages || 0,
                    records: result.records.length,
                    ignoredPages: result.ignoredPages.length,
                    result
                }
            };
        } catch (err) {
            context.error('Fejl i parseIfsReport:', err);

            return {
                status: 500,
                jsonBody: {
                    ok: false,
                    error: err.message
                }
            };
        }
    }
});
