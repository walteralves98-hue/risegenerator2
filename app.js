
/* ADRA RISE Activation Studio
   Self-contained DOCX/XLSX generation using local JSZip.
   PDF import requires locally installed PDF.js.
*/

const ids = [
  'country','customCountry','emergency','scenario','title',
  'onset','start','end','lead','areas','households','hhsize',
  'people','modality','needs','activities','targeting',
  'coordination','distribution','staffing','targetBudget',
  'unit','rounds','delivery','personnel','monitoring','other',
  'indirect','assumptions','c07','c08','feedback1','feedback2',
  'methods','safeguarding','risks','partners'
];

const $ = id => document.getElementById(id);
const v = id => $(id)?.value || '';
const n = id => Number(v(id)) || 0;
const money = x => '$' + Number(x || 0).toLocaleString(
  'en-US', {maximumFractionDigits: 2}
);

const country = () =>
  v('country') === 'Other / custom'
    ? (v('customCountry') || 'Country TBD')
    : v('country');

const projectTitle = () =>
  v('title') || country() + ' Emergency Activation';

const templates = [
  'RISE_Proposal_Template.docx',
  'RISE_Budget_Template.xlsx',
  'RISE_LogFrame_ITT_Template.xlsx',
  'RISE_ME_Plan_Template.docx',
  'RISE_Activation_Checklist.docx'
];

let importedNEMP = null;

function escapeHTML(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;',
    '"':'&quot;', "'":'&#39;'
  })[c]);
}

function xml(s) {
  return String(s ?? '').replace(/[<>&"']/g, c => ({
    '<':'&lt;', '>':'&gt;', '&':'&amp;',
    '"':'&quot;', "'":'&apos;'
  })[c]);
}

function tab(id) {
  document.querySelectorAll('.section').forEach(el =>
    el.classList.toggle('active', el.id === id)
  );
  document.querySelectorAll('[data-tab]').forEach(el =>
    el.classList.toggle('active', el.dataset.tab === id)
  );
  if (id === 'review') renderReview();
  window.scrollTo({top: 0, behavior: 'smooth'});
}

document.querySelectorAll('[data-tab]').forEach(el => {
  el.onclick = () => tab(el.dataset.tab);
});

function budget() {
  const transfers = n('households') * n('unit') * n('rounds');
  const direct = transfers + n('delivery') + n('personnel')
    + n('monitoring') + n('other');
  const indirect = direct * n('indirect') / 100;
  return {
    transfers, direct, indirect,
    total: direct + indirect,
    gap: n('targetBudget') - direct - indirect
  };
}

function issues() {
  const result = [];
  const required = [
    ['title','Response title'],
    ['onset','Disaster onset date'],
    ['start','Start date'],
    ['end','End date'],
    ['areas','Target areas'],
    ['needs','Current needs assessment'],
    ['coordination','Coordination arrangements'],
    ['distribution','Distribution plan'],
    ['staffing','Staffing plan'],
    ['assumptions','Budget justification'],
    ['methods','Monitoring methodology'],
    ['safeguarding','Safeguarding procedures'],
    ['risks','Risk assessment']
  ];

  for (const [id, label] of required) {
    if (!v(id).trim()) result.push(label + ' missing');
  }

  if (n('households') <= 0)
    result.push('Household target must be positive');

  if (n('people') <= 0)
    result.push('Unique beneficiary target required');

  if (v('start') && v('end') && v('end') < v('start'))
    result.push('End date precedes start date');

  if (v('feedback1') === v('feedback2'))
    result.push('Select two different feedback channels');

  if (budget().gap < 0)
    result.push('Estimated costs exceed target funding');

  if (!importedNEMP)
    result.push('No uploaded NEMP is linked to this activation');

  return result;
}

function update() {
  const b = budget();

  $('snapCountry').textContent = 'ADRA ' + country();
  $('snapPeople').textContent = n('people').toLocaleString();
  $('snapTarget').textContent = money(n('targetBudget'));
  $('snapTotal').textContent = money(b.total);
  $('snapGap').textContent =
    (b.gap < 0 ? '-' : '') + money(Math.abs(b.gap));

  const count = issues().length;
  $('bar').style.width = Math.max(0, 100 - count * 6) + '%';
  $('readiness').textContent = count + ' review items remain';

  if ($('review').classList.contains('active')) renderReview();
}

function renderReview() {
  const list = issues();

  $('issues').innerHTML = list.length
    ? '<div class="note"><b>' + list.length +
      ' items require review</b><ul>' +
      list.map(s => '<li>' + escapeHTML(s) + '</li>').join('') +
      '</ul></div>'
    : '<div class="note">Basic checks passed. Technical, finance, ' +
      'compliance and donor approval are still required.</div>';
}

ids.forEach(id => {
  $(id)?.addEventListener('input', update);
  $(id)?.addEventListener('change', update);
});

['households','hhsize'].forEach(id => {
  $(id).addEventListener('input', () => {
    $('people').value = n('households') * n('hhsize');
    update();
  });
});

$('templateLinks').innerHTML = templates.map(t =>
  '<a style="display:block;color:#087a87;margin:8px 0" href="templates/' +
  encodeURIComponent(t) + '" download>' + escapeHTML(t) + '</a>'
).join('');

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 3000);
}

function saveDraft() {
  const data = Object.fromEntries(ids.map(id => [id, v(id)]));
  download(
    new Blob([JSON.stringify(data, null, 2)], {
      type:'application/json'
    }),
    'RISE_activation_draft.json'
  );
}

async function loadDraft(input) {
  const file = input.files?.[0];
  if (!file) return;

  try {
    const data = JSON.parse(await file.text());
    ids.forEach(id => {
      if (data[id] !== undefined) $(id).value = data[id];
    });
    update();
    $('status').textContent = 'Draft loaded successfully.';
  } catch (e) {
    alert('Could not load draft: ' + e.message);
  }
}

/* ---------- UNIVERSAL NEMP IMPORT ---------- */

function cleanText(text) {
  return String(text || '')
    .replace(/\u0000/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

async function readExcelNEMP(buffer) {
  if (!window.JSZip)
    throw Error('Local ZIP library is missing.');

  const zip = await JSZip.loadAsync(buffer);
  const workbookFile = zip.file('xl/workbook.xml');

  if (!workbookFile)
    throw Error('This is not a valid Excel workbook.');

  const parser = new DOMParser();

  const shared = [];
  const sharedFile = zip.file('xl/sharedStrings.xml');

  if (sharedFile) {
    const doc = parser.parseFromString(
      await sharedFile.async('string'), 'application/xml'
    );
    for (const si of doc.getElementsByTagName('si')) {
      shared.push(
        Array.from(si.getElementsByTagName('t'))
          .map(t => t.textContent).join('')
      );
    }
  }

  let result = '';

  const sheetPaths = Object.keys(zip.files)
    .filter(path => /^xl\/worksheets\/sheet\d+\.xml$/.test(path))
    .sort((a, b) => {
      const x = Number(a.match(/sheet(\d+)/)[1]);
      const y = Number(b.match(/sheet(\d+)/)[1]);
      return x - y;
    });

  for (const path of sheetPaths) {
    const sheetXML = await zip.file(path).async('string');
    const doc = parser.parseFromString(sheetXML, 'application/xml');

    result += '\n[Worksheet: ' + path + ']\n';

    for (const row of doc.getElementsByTagName('row')) {
      const values = [];

      for (const cell of row.getElementsByTagName('c')) {
        const type = cell.getAttribute('t');
        let value = '';

        const vNode = cell.getElementsByTagName('v')[0];
        const inline = cell.getElementsByTagName('is')[0];

        if (type === 's' && vNode) {
          value = shared[Number(vNode.textContent)] || '';
        } else if (inline) {
          value = Array.from(inline.getElementsByTagName('t'))
            .map(t => t.textContent).join('');
        } else if (vNode) {
          value = vNode.textContent;
        }

        if (value) values.push(value);
      }

      if (values.length) result += values.join(' | ') + '\n';
    }
  }

  return result;
}

async function readPDFNEMP(buffer) {
  if (!window.pdfjsLib) {
    throw Error(
      'PDF reader is not installed locally. Add pdf.min.js and ' +
      'pdf.worker.min.js to the vendor folder before importing PDFs.'
    );
  }

  pdfjsLib.GlobalWorkerOptions.workerSrc =
    './vendor/pdf.worker.min.js';

  const pdf = await pdfjsLib.getDocument({
    data: new Uint8Array(buffer)
  }).promise;

  let result = '';

  for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
    const page = await pdf.getPage(pageNo);
    const content = await page.getTextContent();

    result += '\n[Page ' + pageNo + ']\n' +
      content.items.map(item => item.str || '').join(' ') + '\n';
  }

  return result;
}

async function readNEMP(input) {
  const file = input.files?.[0];
  if (!file) return;

  $('uploadStatus').textContent = 'Reading ' + file.name + '...';
  $('nempExtract').style.display = 'none';

  try {
    if (file.size > 25 * 1024 * 1024)
      throw Error('Maximum NEMP size is 25 MB.');

    const extension = file.name.split('.').pop().toLowerCase();
    const buffer = await file.arrayBuffer();

    let text;

    if (extension === 'xlsx') {
      text = await readExcelNEMP(buffer);
    } else if (extension === 'pdf') {
      text = await readPDFNEMP(buffer);
    } else {
      throw Error('Upload a PDF or Excel (.xlsx) NEMP.');
    }

    text = cleanText(text);

    if (text.length < 120)
      throw Error(
        'Insufficient extractable text. Scanned PDFs require OCR.'
      );

    const filename = file.name.replace(/\.[^.]+$/, '');

    const countryGuess = filename
      .replace(/^ADRA[\s_-]*/i, '')
      .replace(/^NEMP[\s_-]*/i, '')
      .replace(/[\s_-]*(NEMP|final|english|202\d).*$/i, '')
      .replace(/[_-]/g, ' ')
      .trim();

    const match = text.search(
      /(?:intervention scenario|scenario\s*1|response scenario)/i
    );

    const excerpt = match >= 0
      ? text.slice(match, match + 3500)
      : text.slice(0, 3500);

    importedNEMP = {
      name: file.name,
      text,
      excerpt,
      country: countryGuess,
      uploadedAt: new Date().toISOString()
    };

    $('detectedCountry').value = countryGuess;
    $('detectedIntervention').value = excerpt.slice(0, 1100);
    $('sourceExcerpt').value = excerpt;
    $('nempExtract').style.display = 'block';

    $('uploadStatus').textContent =
      'Extracted ' + text.length.toLocaleString() +
      ' characters. Review the information before applying.';

    update();
  } catch (error) {
    console.error(error);
    $('uploadStatus').textContent = 'Import failed: ' + error.message;
  }
}

function applyNEMP() {
  if (!importedNEMP) return;

  $('country').value = 'Other / custom';
  $('customCountry').value = $('detectedCountry').value.trim();

  const description = $('detectedIntervention').value.trim();

  if (description && !v('assumptions').trim()) {
    $('assumptions').value =
      'NEMP planning reference (verify current prices, targets ' +
      'and assumptions): ' + description.slice(0, 800);
  }

  $('nempNote').textContent =
    'Source: ' + importedNEMP.name +
    '. Historical NEMP values require verification against ' +
    'current needs assessments and market prices.';

  $('uploadStatus').textContent =
    'NEMP linked to activation. Complete and verify response inputs.';

  update();
}

/* ---------- WORD DOCUMENT GENERATOR ---------- */

function wordParagraph(text, style = '') {
  const lines = String(text ?? '').split('\n');

  return lines.map(line =>
    '<w:p>' +
    (style ? '<w:pPr><w:pStyle w:val="' + style +
      '"/></w:pPr>' : '') +
    '<w:r><w:t xml:space="preserve">' +
    xml(line || ' ') +
    '</w:t></w:r></w:p>'
  ).join('');
}

async function createWord(name, sections) {
  const zip = new JSZip();

  zip.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>'
  );

  zip.file('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'
  );

  zip.file('word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>'
  );

  zip.file('word/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="21"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1">' +
    '<w:name w:val="heading 1"/><w:rPr><w:b/><w:color w:val="153E52"/><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2">' +
    '<w:name w:val="heading 2"/><w:rPr><w:b/><w:color w:val="087A87"/><w:sz w:val="26"/></w:rPr></w:style>' +
    '</w:styles>'
  );

  let body =
    wordParagraph('ADRA INTERNATIONAL | RISE PILLAR 1', 'Heading1') +
    wordParagraph(name, 'Heading1') +
    wordParagraph(
      'Country: ' + country() +
      ' | Emergency: ' + v('emergency') +
      ' | Draft for technical and finance review'
    );

  for (const [heading, content] of sections) {
    body += wordParagraph(heading, 'Heading2');
    body += wordParagraph(content);
  }

  zip.file('word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' + body +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>' +
    '</w:sectPr></w:body></w:document>'
  );

  return zip.generateAsync({type:'blob'});
}

/* ---------- EXCEL WORKBOOK GENERATOR ---------- */

function columnName(index) {
  let name = '';
  while (index > 0) {
    index--;
    name = String.fromCharCode(65 + index % 26) + name;
    index = Math.floor(index / 26);
  }
  return name;
}

function excelSheet(rows) {
  const body = rows.map((row, ri) => {
    const cells = row.map((value, ci) => {
      const ref = columnName(ci + 1) + (ri + 1);

      if (typeof value === 'number' && Number.isFinite(value)) {
        return '<c r="' + ref + '"><v>' + value + '</v></c>';
      }

      return '<c r="' + ref + '" t="inlineStr"><is><t>' +
        xml(value ?? '') + '</t></is></c>';
    }).join('');

    return '<row r="' + (ri + 1) + '">' + cells + '</row>';
  }).join('');

  return '<?xml version="1.0" encoding="UTF-8"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' + body + '</sheetData></worksheet>';
}

async function createExcel(sheets) {
  const zip = new JSZip();
  const names = Object.keys(sheets);

  zip.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    names.map((_, i) =>
      '<Override PartName="/xl/worksheets/sheet' + (i + 1) +
      '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    ).join('') + '</Types>'
  );

  zip.file('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>'
  );

  zip.file('xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets>' +
    names.map((name, i) =>
      '<sheet name="' + xml(name) + '" sheetId="' + (i + 1) +
      '" r:id="rId' + (i + 1) + '"/>'
    ).join('') + '</sheets></workbook>'
  );

  zip.file('xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    names.map((_, i) =>
      '<Relationship Id="rId' + (i + 1) +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' +
      ' Target="worksheets/sheet' + (i + 1) + '.xml"/>'
    ).join('') + '</Relationships>'
  );

  names.forEach((name, i) => {
    zip.file(
      'xl/worksheets/sheet' + (i + 1) + '.xml',
      excelSheet(sheets[name])
    );
  });

  return zip.generateAsync({type:'blob'});
}

/* ---------- PROPOSAL CONTENT ---------- */

function proposalSections() {
  const b = budget();

  return [
    ['1. Project Summary',
      'Project: ' + projectTitle() +
      '\nCountry: ' + country() +
      '\nEmergency: ' + v('emergency') +
      '\nTarget areas: ' + v('areas') +
      '\nImplementation: ' + v('start') + ' to ' + v('end') +
      '\nHouseholds: ' + n('households') +
      '\nUnique beneficiaries (C00): ' + n('people') +
      '\nModality: ' + v('modality') +
      '\nEstimated budget: ' + money(b.total)
    ],
    ['2. Sub-Awards and Partners',
      v('partners') || 'To be confirmed.'
    ],
    ['3. Description of Need',
      v('needs') || 'Current needs assessment required.'
    ],
    ['4. Government and Humanitarian Coordination',
      v('coordination')
    ],
    ['5. Proposed Response Activities',
      'Objective: Save lives and reduce suffering among ' +
      'disaster-affected people.\n\n' +
      'Activities:\n' + v('activities') +
      '\n\nTargeting:\n' + v('targeting') +
      '\n\nC00: ' + n('people') +
      '\nC07 target: ' + n('c07') + '%' +
      '\nC08 target: ' + n('c08') + '%'
    ],
    ['6. Implementation Timeline',
      'Disaster onset: ' + v('onset') +
      '\nStart: ' + v('start') +
      '\nEnd: ' + v('end')
    ],
    ['7. Budget',
      'Assistance: ' + money(b.transfers) +
      '\nDirect costs: ' + money(b.direct) +
      '\nEstimated indirect: ' + money(b.indirect) +
      '\nTotal: ' + money(b.total) +
      '\n\nAssumptions:\n' + v('assumptions') +
      '\n\nFinance must validate the official RISE budget.'
    ],
    ['8. Implementation and Staffing',
      v('staffing') + '\n\nDistribution:\n' + v('distribution')
    ],
    ['9. International Travel',
      'Confirm any planned international travel and approvals.'
    ],
    ['10. Risk Assessment and Management', v('risks')],
    ['11. Sanctions and Diversion Risk',
      'Verify partner and vendor screening, diversion ' +
      'controls, mitigation and reporting requirements.\n' +
      v('risks')
    ],
    ['12. Accountability and Safeguarding',
      'Feedback channels: ' + v('feedback1') +
      ' and ' + v('feedback2') +
      '\nSafeguarding: ' + v('safeguarding')
    ],
    ['13. NEMP Source and Adaptation',
      importedNEMP
        ? 'Source NEMP: ' + importedNEMP.name +
          '\nReviewed intervention:\n' +
          v('detectedIntervention') +
          '\nHistorical figures must be independently verified.'
        : 'No source NEMP uploaded.'
    ]
  ];
}

function mealSections() {
  return [
    ['Project Information',
      projectTitle() + '\nADRA ' + country() +
      '\nDates: ' + v('start') + ' to ' + v('end')
    ],
    ['1. Participant Identification',
      'C00 target: ' + n('people') +
      '\nTargeting: ' + v('targeting')
    ],
    ['2. Baseline', v('methods')],
    ['3. Routine Monitoring', v('methods')],
    ['4. On-Site Monitoring', v('distribution')],
    ['5. Post-Distribution Monitoring', v('methods')],
    ['6. Endline', 'Define endline methodology and timing.'],
    ['7. Qualitative Data Collection',
      'Define interviews, focus groups and consent procedures.'
    ],
    ['8. After-Action Review',
      'Document lessons learned following implementation.'
    ],
    ['9. Real-Time Evaluation',
      'Determine whether a real-time evaluation is required.'
    ],
    ['10. Accountability and Feedback',
      'Channels: ' + v('feedback1') + ', ' + v('feedback2') +
      '\nSafeguarding: ' + v('safeguarding') +
      '\nC07 target: ' + n('c07') + '%' +
      '\nC08 target: ' + n('c08') + '%'
    ]
  ];
}

function budgetSheets() {
  const b = budget();

  return {
    'Activation Inputs': [
      ['Parameter','Value','Notes'],
      ['Country',country(),''],
      ['Project',projectTitle(),''],
      ['Emergency',v('emergency'),''],
      ['Households',n('households'),''],
      ['Household size',n('hhsize'),''],
      ['Unique individuals',n('people'),'C00'],
      ['Target funding',n('targetBudget'),'USD'],
      ['Unit assistance',n('unit'),'USD'],
      ['Rounds',n('rounds'),''],
      ['Indirect rate',n('indirect') / 100,'Planning only']
    ],
    'Planning Budget': [
      ['Cost category','Quantity','Unit cost','Total USD'],
      ['Assistance',n('households') * n('rounds'),
        n('unit'),b.transfers],
      ['Delivery',1,n('delivery'),n('delivery')],
      ['Personnel',1,n('personnel'),n('personnel')],
      ['Monitoring and AAP',1,n('monitoring'),n('monitoring')],
      ['Other direct costs',1,n('other'),n('other')],
      ['Total direct costs','','',b.direct],
      ['Estimated indirect','','',b.indirect],
      ['Estimated total','','',b.total],
      ['Target funding','','',n('targetBudget')],
      ['Funding gap','','',b.gap]
    ],
    'Assumptions': [
      ['Item','Detail'],
      ['Cost justification',v('assumptions')],
      ['Source NEMP',importedNEMP?.name || 'Not uploaded'],
      ['Finance note',
        'Planning estimates only. Complete the official RISE budget.']
    ]
  };
}

function indicatorSheets() {
  return {
    'Indicators': [
      ['Code','Indicator','Target','Unit','Data source'],
      ['C00','Direct project beneficiaries',
        n('people'),'Individuals','Registration records'],
      ['C07','People reporting they feel heard',
        n('c07') / 100,'Proportion','PDM / survey'],
      ['C08','Feedback acknowledged, actioned and resolved',
        n('c08') / 100,'Proportion','Feedback register']
    ],
    'Project': [
      ['Field','Value'],
      ['Project',projectTitle()],
      ['Country',country()],
      ['Start',v('start')],
      ['End',v('end')]
    ]
  };
}

function checklistSheets() {
  return {
    'Activation Checklist': [
      ['Phase','Requirement','Status'],
      ['Pre-submission','Rapid needs assessment','Pending'],
      ['Pre-submission','Government and DHR coordination','Pending'],
      ['Submission','RISE proposal','Draft'],
      ['Submission','Official detailed budget','Pending'],
      ['Submission','Distribution plan','Draft'],
      ['Before implementation','Indicator tracker','Draft'],
      ['Before implementation','M&E Plan Brief','Draft'],
      ['Before distribution','Feedback mechanisms','Pending'],
      ['Before distribution','Safeguarding and PSEA','Pending'],
      ['Before implementation','Security and diversion controls',
        'Pending']
    ]
  };
}

/* ---------- EXPORT COMPLETE PACKAGE ---------- */

async function exportAll() {
  if (!window.JSZip) {
    const message =
      'Local ZIP library is missing. Verify that ' +
      'vendor/jszip.min.js exists in GitHub.';
    $('status').textContent = message;
    alert(message);
    return;
  }

  const button = $('generate');
  button.disabled = true;
  $('status').textContent = 'Generating proposal package...';

  try {
    const zip = new JSZip();

    zip.file(
      '01_RISE_Activation_Proposal_DRAFT.docx',
      await createWord(
        'Pillar 1 Emergency Response Activation Proposal',
        proposalSections()
      )
    );

    zip.file(
      '02_RISE_ME_Plan_Brief_DRAFT.docx',
      await createWord(
        'Monitoring and Evaluation Plan Brief',
        mealSections()
      )
    );

    zip.file(
      '03_RISE_Budget_Planning_DRAFT.xlsx',
      await createExcel(budgetSheets())
    );

    zip.file(
      '04_RISE_Indicator_Tracking_DRAFT.xlsx',
      await createExcel(indicatorSheets())
    );

    zip.file(
      '05_RISE_Activation_Checklist_DRAFT.xlsx',
      await createExcel(checklistSheets())
    );

    zip.file(
      '06_Activation_Inputs.json',
      JSON.stringify(
        Object.fromEntries(ids.map(id => [id, v(id)])),
        null, 2
      )
    );

    if (importedNEMP) {
      zip.file(
        '07_NEMP_Extracted_Text.txt',
        importedNEMP.text
      );

      zip.file(
        '08_NEMP_Source_Metadata.json',
        JSON.stringify({
          filename: importedNEMP.name,
          uploadedAt: importedNEMP.uploadedAt,
          country: v('detectedCountry'),
          intervention: v('detectedIntervention')
        }, null, 2)
      );
    }

    zip.file(
      'READ_ME_FIRST.txt',
      'ADRA RISE Activation Studio\n\n' +
      'This package contains editable planning drafts.\n' +
      'These are not completed official donor templates.\n' +
      'All needs, targets, costs, indicators and assumptions ' +
      'require technical and finance review.\n' +
      'Obtain applicable approvals before submission.\n'
    );

    for (const name of templates) {
      try {
        const response = await fetch(
          './templates/' + encodeURIComponent(name)
        );

        if (response.ok) {
          zip.file(
            'Original_Templates/' + name,
            await response.arrayBuffer()
          );
        }
      } catch (error) {
        console.warn('Template unavailable:', name, error);
      }
    }

    const blob = await zip.generateAsync({
      type:'blob',
      compression:'DEFLATE',
      compressionOptions:{level:5}
    });

    download(
      blob,
      'ADRA_RISE_' +
      country().replace(/[^a-z0-9]+/gi, '_') +
      '_Activation_Package.zip'
    );

    $('status').textContent =
      'Proposal package generated. Review all documents ' +
      'before submission.';
  } catch (error) {
    console.error(error);
    $('status').textContent =
      'Generation failed: ' + error.message;
    alert('Generation failed: ' + error.message);
  } finally {
    button.disabled = false;
  }
}

/* ---------- INITIALIZATION ---------- */

$('nempNote').textContent =
  'Upload a PDF or Excel NEMP, review the extracted information, ' +
  'and adapt the response using current emergency evidence.';

update();
