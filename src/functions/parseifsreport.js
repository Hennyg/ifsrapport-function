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
// Labels  (dansk + engelsk)
// --------------------------------------------------
const LABELS = {
    interventionDate: ['Intervention Date', 'Besøgs dato', 'Besøgsdato'],
    serviceType:      ['Service Type', 'Servicetype'],
    technicianName:   ['Technician Name', 'Teknikerens navn', 'Tekniker'],
    reference:        ['Reference'],
    task:             ['Task', 'Opgave'],
    customer:         ['Customer', 'Kunde'],
    customerName:     ['Customer Name', 'Kundenavn'],
    customerContact:  ['Customer Contact', 'Kundekontakt'],
    description:      ['Description', 'Beskrivelse'],
    product:          ['Product', 'Produkt'],
    partsUsed:        ['Parts Used', 'Brugte dele'],
    partsReturned:    ['Parts Returned', 'Dele returneres'],
    labourAndExpense: ['Labour and Expenses', 'Labour and Expense', 'Arbejdskraft og Udgifter'],
    notes:            ['Notes', 'Noter']
};

// VIGTIGT: 'Description'/'Beskrivelse' er bevidst udeladt fra ALL_SECTION_HEADINGS.
// De optræder også som kolonneoverskrifter inde i Parts Used/Returned-sektionerne,
// og ville afskære sektionen for tidligt hvis de var inkluderet.
// getSectionLines stopper korrekt ved de øvrige sektioner.
const ALL_SECTION_HEADINGS = [
    ...LABELS.customer,
    ...LABELS.customerName,
    ...LABELS.customerContact,
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
        kunde:         kundeBlock[0] || null,
        kundeadresse:  kundeBlock.length > 1 ? kundeBlock.slice(1).join(', ') : null,
        kundenummer:   getValueAfterAnyLabel(lines, LABELS.customerName),
        kundekontakt:  getValueAfterAnyLabel(lines, LABELS.customerContact),
        besogsdato:    normalizeDate(getValueAfterAnyLabel(lines, LABELS.interventionDate)),
        servicetype:   getValueAfterAnyLabel(lines, LABELS.serviceType),
        servicenummer: serviceOrder || extractServiceNumberFromFileName(fileName),
        tekniker:      getValueAfterAnyLabel(lines, LABELS.technicianName),
        reference:     getValueAfterAnyLabel(lines, LABELS.reference),
        task:          getValueAfterAnyLabel(lines, LABELS.task)
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
        return { serienummer: null, beskrivelse: null, model: null };
    }

    // Forsøg 1: alt på én linje "0003275775 103 Astronaut A5 RH A5_RIGHT"
    const line = productLines[0];
    const m = line.match(/^([A-Z0-9-]+)\s+(.+?)\s+([A-Z0-9_]+)$/i);
    if (m) {
        return {
            serienummer: cleanValue(m[1]),
            beskrivelse: cleanValue(m[2]),
            model:       cleanValue(m[3])
        };
    }

    // Forsøg 2: tre separate linjer
    // linje 0: serienummer (kun tal/bogstaver, ingen mellemrum)
    // linje 1: beskrivelse
    // linje 2: model (ingen mellemrum, typisk A5_RIGHT, A4CU osv.)
    const sn    = productLines[0] || null;
    const besk  = productLines[1] || null;
    const model = productLines[2] || null;

    // Valider at serienummer ligner et serienummer (kun tal)
    const snLooksLikeSerialNr = sn && /^\d{7,}$/.test(sn.trim());
    // Valider at model ligner en model (ingen mellemrum, alphanumerisk+underscore)
    const modelLooksLikeModel = model && /^[A-Z0-9_]+$/i.test(model.trim());

    if (snLooksLikeSerialNr) {
        return {
            serienummer: cleanValue(sn),
            beskrivelse: besk ? cleanValue(besk) : null,
            model:       modelLooksLikeModel ? cleanValue(model) : null
        };
    }

    // Fallback
    return { serienummer: null, beskrivelse: cleanValue(line), model: null };
}

// --------------------------------------------------
// Parts line parser
// --------------------------------------------------
function parseItemLine(rawLine) {
    const line = normalizeCompactedLine(rawLine);

    // Matcher varenummer-formater:
    //   15EXP101, 15EXP201     (arbejdstid/kørsel)
    //   5.1003.0373.0           (reservedel: mindst 2 punktummer)
    //   151084                  (6+ cifre uden punktum)
    // IKKE: 0.50, 1.00, 2.00   (antal-tal med ét punktum afvises)
    const match = line.match(/^((?:15[A-Z]{3}\d{3})|(?:\d+(?:\.\d+){2,})|(?:\d{6,}))\s*(.*)$/i);
    if (!match) return null;

    const varenummer = match[1].trim();
    let rest = cleanValue(match[2]);
    if (!rest) return { varenummer, beskrivelse: null, antal: null, rawLine };

    // Fjern serie-felt "X" eller enkelt bogstav midt i linjen før antal
    rest = rest.replace(/\bX\b\s*/i, '').trim();

    // Antal sidst på linjen
    const qtyMatch = rest.match(/(\d+(?:[.,]\d+)?)\s*$/);
    if (!qtyMatch) {
        return { varenummer, beskrivelse: cleanValue(rest) || null, antal: null, rawLine };
    }

    const antal = toNumber(qtyMatch[1]);
    const beskrivelse = cleanValue(
        rest.substring(0, rest.length - qtyMatch[1].length).trim()
    ) || null;

    return { varenummer, beskrivelse, antal, rawLine };
}

// --------------------------------------------------
// Brugte dele
// --------------------------------------------------
function parseBrugteDele(lines, pageNumber) {
    const sectionLines = getSectionLines(lines, LABELS.partsUsed);

    const arbejdstid = [];
    const reservedele = [];

    let i = 0;
    while (i < sectionLines.length) {
        const rawLine = sectionLines[i];

        if (isHeaderLine(rawLine) || isNoiseLine(rawLine)) { i++; continue; }

        // Forsøg 1: varenummer + beskrivelse + [serial] + antal på én linje
        let item = parseItemLine(rawLine);

        // Forsøg 2: PDF har splittet linjen over flere linjer
        // f.eks. "15EXP101\nTimer hverdag 07-16\n1.00"
        if (item && item.beskrivelse === null && item.antal === null) {
            const nextLine  = sectionLines[i + 1] || '';
            const afterLine = sectionLines[i + 2] || '';

            const nextIsDesc   = nextLine  && !isHeaderLine(nextLine)  && !/^\d+(?:[.,]\d+)?$/.test(nextLine.trim());
            const afterIsAntal = afterLine && /^\d+(?:[.,]\d+)?$/.test(afterLine.trim());

            if (nextIsDesc && afterIsAntal) {
                item = {
                    varenummer:  item.varenummer,
                    beskrivelse: cleanValue(nextLine) || null,
                    antal:       toNumber(afterLine.trim()),
                    rawLine:     `${rawLine} ${nextLine} ${afterLine}`
                };
                i += 3;
            } else if (nextIsDesc) {
                const embedded = parseItemLine(`${item.varenummer} ${nextLine}`);
                if (embedded) {
                    item = { ...embedded, rawLine: `${rawLine} ${nextLine}` };
                } else {
                    item = { varenummer: item.varenummer, beskrivelse: cleanValue(nextLine), antal: null, rawLine: `${rawLine} ${nextLine}` };
                }
                i += 2;
            } else {
                i++;
            }
        } else if (item) {
            i++;
        } else {
            i++;
            continue;
        }

        if (!item) continue;

        const obj = {
            pageNumber,
            varenummer:  item.varenummer,
            beskrivelse: item.beskrivelse,
            antal:       item.antal,
            rawLine:     item.rawLine
        };

        const key = x => `${x.pageNumber}|${x.varenummer}|${x.beskrivelse}|${x.antal}`;

        if (isArbejdstidVarenummer(item.varenummer)) {
            uniquePush(arbejdstid, obj, key);
        } else {
            uniquePush(reservedele, obj, key);
        }
    }

    return { arbejdstid, reservedele };
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
                varenummer:  item.varenummer,
                beskrivelse: item.beskrivelse,
                antal:       item.antal,
                rawLine:     item.rawLine
            },
            x => `${x.pageNumber}|${x.varenummer}|${x.beskrivelse}|${x.antal}`
        );
    }

    return { linjer: results };
}

// --------------------------------------------------
// Arbejdskraft og udgifter
// Læser linjer via getSectionLines (samme teknik som parseBrugteDele)
// for at undgå problemer med sektionsopskæring i rå tekst.
// --------------------------------------------------
function parseArbejdskraftOgUdgifter(lines, pageNumber) {
    const sectionLines = getSectionLines(lines, LABELS.labourAndExpense);

    const results = [];

    let i = 0;
    while (i < sectionLines.length) {
        const rawLine = sectionLines[i];
        const line = normalizeCompactedLine(rawLine);

        if (isHeaderLine(rawLine) || isNoiseLine(rawLine)) { i++; continue; }
        if (/^service order:/i.test(rawLine)) { i++; continue; }
        if (/^\(utc/i.test(rawLine)) { i++; continue; }
        if (/^[A-ZÆØÅ.\- ]+A\/S$/i.test(rawLine)) { i++; continue; }

        // Forsøg 1: beskrivelse + dato + beløb på én linje
        const oneLiner = line.match(/^(.*?)(\d{2}[-/]\d{2}[-/]\d{4})\s*(\d+(?:[.,]\d+)?)$/);
        if (oneLiner) {
            const beskrivelse = cleanValue(oneLiner[1]) || null;
            const dato        = normalizeDate(oneLiner[2]);
            const beloeb      = toNumber(oneLiner[3]);

            if (beskrivelse || dato || beloeb !== null) {
                uniquePush(
                    results,
                    { pageNumber, beskrivelse, dato, beloeb, rawLine },
                    x => `${x.pageNumber}|${x.beskrivelse}|${x.dato}|${x.beloeb}`
                );
            }
            i++;
            continue;
        }

        // Forsøg 2: beskrivelse / dato / beløb på separate linjer
        const nextLine  = sectionLines[i + 1] || '';
        const afterLine = sectionLines[i + 2] || '';

        const isDate   = /^\d{2}[-/]\d{2}[-/]\d{4}$/.test(nextLine.trim());
        const isAmount = /^\d+(?:[.,]\d+)?$/.test(afterLine.trim());

        if (isDate && isAmount) {
            const beskrivelse = cleanValue(rawLine) || null;
            const dato        = normalizeDate(nextLine.trim());
            const beloeb      = toNumber(afterLine.trim());

            uniquePush(
                results,
                { pageNumber, beskrivelse, dato, beloeb, rawLine: `${rawLine} ${nextLine} ${afterLine}` },
                x => `${x.pageNumber}|${x.beskrivelse}|${x.dato}|${x.beloeb}`
            );
            i += 3;
            continue;
        }

        // Forsøg 3: beskrivelse+dato på én linje, beløb på næste
        const halfLiner = line.match(/^(.*?)(\d{2}[-/]\d{2}[-/]\d{4})\s*$/);
        if (halfLiner && isAmount) {
            const beskrivelse = cleanValue(halfLiner[1]) || null;
            const dato        = normalizeDate(halfLiner[2]);
            const beloeb      = toNumber(afterLine.trim());

            uniquePush(
                results,
                { pageNumber, beskrivelse, dato, beloeb, rawLine: `${rawLine} ${afterLine}` },
                x => `${x.pageNumber}|${x.beskrivelse}|${x.dato}|${x.beloeb}`
            );
            i += 2;
            continue;
        }

        i++;
    }

    return {
        linjer: results,
        raw:    sectionLines.length ? sectionLines.join('\n') : null
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
        kunde:                  parseKunde(lines, fileName, serviceOrder),
        beskrivelse:            parseBeskrivelse(lines),
        produkt:                parseProdukt(lines),
        brugteDele:             parseBrugteDele(lines, pageNumber),
        deleReturneret:         parseDeleReturneret(lines, pageNumber),
        arbejdskraftOgUdgifter: parseArbejdskraftOgUdgifter(lines, pageNumber),
        noter:                  parseNoter(lines),
        rawText:                cleanValue(pageText)
    };
}

// --------------------------------------------------
// Build document
// --------------------------------------------------
function buildDocument(text, numPages, fileName) {
    const serviceOrder = extractServiceOrder(text, fileName);
    const pageTexts    = splitPages(text);

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
        filnavn:       fileName || null,
        serviceOrder,
        detectedPages: pageTexts.length,
        reportedPages: numPages,
        records,
        ignoredPages,
        pages:         allPages
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
                    jsonBody: { ok: false, error: 'pdfBase64 mangler' }
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
                    ok:           true,
                    pages:        result.records.length,
                    records:      result.records.length,
                    ignoredPages: result.ignoredPages.length,
                    result: {
                        filnavn:       result.filnavn,
                        serviceOrder:  result.serviceOrder,
                        detectedPages: result.detectedPages,
                        reportedPages: result.reportedPages,
                        records:       result.records,
                        ignoredPages:  result.ignoredPages
                        // pages udeladt - undgår duplikerede data
                    }
                }
            };
        } catch (err) {
            context.error('Fejl i parseIfsReport:', err);

            return {
                status: 500,
                jsonBody: { ok: false, error: err.message }
            };
        }
    }
});
