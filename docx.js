// Minimal OpenXML DOCX writer. The ZIP contains ordinary, uncompressed OOXML parts.
const encoder = new TextEncoder();
const xml = value => String(value ?? '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[char]);

export function extractParagraphs(html) {
  const article = new DOMParser().parseFromString(`<article>${html}</article>`, 'text/html').querySelector('article');
  const paragraphs = [];
  const text = node => {
    const clone = node.cloneNode(true);
    clone.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
    return clone.textContent.trim();
  };
  const add = (node, style = 'Normal') => { if (node && text(node)) paragraphs.push({ text: text(node), style }); };
  add(article.querySelector('.doc-head > div:first-child'));
  add(article.querySelector('.doc-head .approval'));
  const title = article.querySelector('.doc-title');
  if (title) {
    const main = title.cloneNode(true);
    main.querySelector('small')?.remove();
    add(main, 'Title');
    add(title.querySelector('small'), 'Subtitle');
  }
  add(article.querySelector('.doc-place'), 'Center');
  article.querySelectorAll('.doc-section').forEach(section => {
    add(section.querySelector('h3'), 'Heading1');
    section.querySelectorAll('p, li').forEach(node => {
      const link = node.querySelector('a[href]');
      paragraphs.push({ text: text(node) + (link ? ` (${link.href})` : ''), style: 'Normal' });
    });
  });
  add(article.querySelector('.doc-sign'));
  add(article.querySelector('.doc-draft'), 'Center');
  return paragraphs;
}

function paragraphXml({ text, style }) {
  const pStyle = ['Normal', 'Title', 'Subtitle', 'Heading1', 'Center'].includes(style) ? style : 'Normal';
  const runs = String(text).split('\n').map((line, index) => `${index ? '<w:r><w:br/></w:r>' : ''}<w:r><w:t xml:space="preserve">${xml(line)}</w:t></w:r>`).join('');
  return `<w:p><w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr>${runs}</w:p>`;
}

function crc32(data) {
  let crc = -1;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ -1) >>> 0;
}

function zip(entries) {
  const chunks = [];
  const directory = [];
  let offset = 0;
  const write16 = (view, pos, value) => view.setUint16(pos, value, true);
  const write32 = (view, pos, value) => view.setUint32(pos, value, true);
  for (const [name, content] of entries) {
    const filename = encoder.encode(name);
    const data = encoder.encode(content);
    const checksum = crc32(data);
    const local = new Uint8Array(30 + filename.length);
    const head = new DataView(local.buffer);
    write32(head, 0, 0x04034b50); write16(head, 4, 20); write16(head, 6, 0x800);
    write32(head, 14, checksum); write32(head, 18, data.length); write32(head, 22, data.length);
    write16(head, 26, filename.length); local.set(filename, 30);
    chunks.push(local, data);
    const central = new Uint8Array(46 + filename.length);
    const record = new DataView(central.buffer);
    write32(record, 0, 0x02014b50); write16(record, 4, 20); write16(record, 6, 20); write16(record, 8, 0x800);
    write32(record, 16, checksum); write32(record, 20, data.length); write32(record, 24, data.length);
    write16(record, 28, filename.length); write32(record, 42, offset); central.set(filename, 46);
    directory.push(central);
    offset += local.length + data.length;
  }
  const directorySize = directory.reduce((sum, entry) => sum + entry.length, 0);
  const end = new Uint8Array(22);
  const trailer = new DataView(end.buffer);
  write32(trailer, 0, 0x06054b50); write16(trailer, 8, entries.length); write16(trailer, 10, entries.length);
  write32(trailer, 12, directorySize); write32(trailer, 16, offset);
  const result = new Uint8Array(offset + directorySize + end.length);
  let position = 0;
  for (const chunk of [...chunks, ...directory, end]) { result.set(chunk, position); position += chunk.length; }
  return result;
}

export function buildDocx(paragraphs) {
  const body = paragraphs.map(paragraphXml).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:jc w:val="center"/><w:spacing w:before="300" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:pPr><w:spacing w:before="280" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Center"><w:name w:val="Center"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>`;
  return zip([
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ['word/document.xml', documentXml], ['word/styles.xml', styles]
  ]);
}
