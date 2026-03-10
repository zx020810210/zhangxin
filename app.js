const templateInput = document.getElementById('templateFile');
const targetInput = document.getElementById('targetFile');
const convertBtn = document.getElementById('convertBtn');
const statusEl = document.getElementById('status');
const formatSummaryEl = document.getElementById('formatSummary');

const STYLE_FILES = ['word/styles.xml', 'word/fontTable.xml', 'word/numbering.xml'];
const THEME_FILES = ['word/theme/theme1.xml'];

function setStatus(text, type = '') {
  statusEl.textContent = text;
  statusEl.className = `status ${type}`.trim();
}

function setFormatSummary(lines, type = '') {
  if (!formatSummaryEl) return;
  formatSummaryEl.innerHTML = '';
  if (!lines.length) {
    formatSummaryEl.innerHTML = '<li>未识别到可展示的模板格式信息。</li>';
    formatSummaryEl.className = `format-summary ${type}`.trim();
    return;
  }

  const frag = document.createDocumentFragment();
  lines.forEach((line) => {
    const li = document.createElement('li');
    li.textContent = line;
    frag.appendChild(li);
  });
  formatSummaryEl.appendChild(frag);
  formatSummaryEl.className = `format-summary ${type}`.trim();
}

function uniq(items) {
  return [...new Set(items.filter(Boolean))];
}

function getCheckedOptions() {
  return new Set([...document.querySelectorAll('.options input[type="checkbox"]:checked')].map((item) => item.value));
}

function getTextFromZip(zip, path) {
  const file = zip.file(path);
  return file ? file.async('string') : Promise.resolve(null);
}

async function copyIfExists(fromZip, toZip, path) {
  const file = fromZip.file(path);
  if (!file) return false;
  const content = await file.async('uint8array');
  toZip.file(path, content);
  return true;
}

function normalizeRelationshipTarget(target) {
  if (!target) return target;
  return target.replace(/(^|\/)setting\.xml(\?|#|$)/i, '$1settings.xml$2');
}

function normalizeContentTypePartName(partName) {
  if (!partName) return partName;
  return partName.replace(/(^|\/)setting\.xml$/i, '$1settings.xml');
}

function normalizePartPath(basePath, target) {
  const normalizedTarget = normalizeRelationshipTarget(target || '');
  const baseDir = basePath.split('/').slice(0, -1);
  const raw = `${baseDir.join('/')}/${normalizedTarget}`.split('/');
  const normalized = [];

  for (const segment of raw) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      normalized.pop();
      continue;
    }
    normalized.push(segment);
  }

  return normalized.join('/');
}

async function loadDocxZip(file, label) {
  const arrayBuffer = await file.arrayBuffer();
  let zip;
  try {
    zip = await JSZip.loadAsync(arrayBuffer);
  } catch {
    throw new Error(`${label} 不是有效的 DOCX 压缩包。`);
  }

  if (!zip.file('[Content_Types].xml') || !zip.file('word/document.xml')) {
    throw new Error(`${label} 缺少 Word 主文档结构（word/document.xml）。`);
  }

  return zip;
}

function parseXml(xml, label) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  if (doc.querySelector('parsererror')) {
    throw new Error(`${label} XML 结构异常，无法解析。`);
  }
  return doc;
}

function findBodySectPr(bodyEl) {
  if (!bodyEl) return null;
  return [...bodyEl.children].find((node) => node.localName === 'sectPr') || null;
}

function replaceSectPr(templateDocXml, targetDocXml) {
  const templateDoc = parseXml(templateDocXml, '模板文档');
  const targetDoc = parseXml(targetDocXml, '待转换文档');
  const templateBody = templateDoc.getElementsByTagNameNS('*', 'body')[0];
  const targetBody = targetDoc.getElementsByTagNameNS('*', 'body')[0];

  if (!templateBody || !targetBody) {
    throw new Error('文档缺少 body 节点，无法同步页面设置。');
  }

  const templateSectPr = findBodySectPr(templateBody);
  if (!templateSectPr) return targetDocXml;

  const targetSectPr = findBodySectPr(targetBody);
  const importedSectPr = targetDoc.importNode(templateSectPr, true);
  if (targetSectPr) {
    targetBody.replaceChild(importedSectPr, targetSectPr);
  } else {
    targetBody.appendChild(importedSectPr);
  }

  return new XMLSerializer().serializeToString(targetDoc);
}

function getPartRelsPath(partPath) {
  const idx = partPath.lastIndexOf('/');
  const dir = idx > -1 ? partPath.slice(0, idx) : '';
  const name = idx > -1 ? partPath.slice(idx + 1) : partPath;
  return `${dir}/_rels/${name}.rels`;
}

async function copyPartWithDependencies(templateZip, targetZip, partPath, copiedParts, visitedParts = new Set()) {
  const normalizedPartPath = normalizePartPath('word/document.xml', partPath);
  if (visitedParts.has(normalizedPartPath)) return;
  visitedParts.add(normalizedPartPath);

  const partCopied = await copyIfExists(templateZip, targetZip, normalizedPartPath);
  if (!partCopied) return;
  copiedParts.add(normalizedPartPath);

  const partRelsPath = getPartRelsPath(normalizedPartPath);
  const partRelsXml = await getTextFromZip(templateZip, partRelsPath);
  if (!partRelsXml) return;

  const relsDoc = parseXml(partRelsXml, `${normalizedPartPath} 关系`);
  const relNodes = [...relsDoc.getElementsByTagNameNS('*', 'Relationship')];

  for (const relNode of relNodes) {
    const targetMode = relNode.getAttribute('TargetMode');
    const rawTarget = relNode.getAttribute('Target');
    const normalizedTarget = normalizeRelationshipTarget(rawTarget);
    if (normalizedTarget) relNode.setAttribute('Target', normalizedTarget);

    if (targetMode === 'External' || !normalizedTarget) continue;

    const childPartPath = normalizePartPath(normalizedPartPath, normalizedTarget);
    if (childPartPath.startsWith('word/')) {
      await copyPartWithDependencies(templateZip, targetZip, childPartPath, copiedParts, visitedParts);
    }
  }

  targetZip.file(partRelsPath, new XMLSerializer().serializeToString(relsDoc));
}

async function mergeContentTypes(templateZip, targetZip, copiedParts) {
  const [templateContentTypesXml, targetContentTypesXml] = await Promise.all([
    getTextFromZip(templateZip, '[Content_Types].xml'),
    getTextFromZip(targetZip, '[Content_Types].xml')
  ]);

  if (!templateContentTypesXml || !targetContentTypesXml) return;

  const templateDoc = parseXml(templateContentTypesXml, '模板内容类型');
  const targetDoc = parseXml(targetContentTypesXml, '待转换内容类型');
  const templateRoot = templateDoc.getElementsByTagName('Types')[0];
  const targetRoot = targetDoc.getElementsByTagName('Types')[0];
  if (!templateRoot || !targetRoot) return;

  const copiedPartNames = new Set([...copiedParts].map((path) => normalizeContentTypePartName(`/${path}`)));

  [...templateRoot.children].forEach((node) => {
    if (node.localName === 'Override') {
      const partName = normalizeContentTypePartName(node.getAttribute('PartName') || '');
      if (!copiedPartNames.has(partName)) return;
      const exists = [...targetRoot.children].some((targetNode) => (
        targetNode.localName === 'Override' && normalizeContentTypePartName(targetNode.getAttribute('PartName') || '') === partName
      ));
      if (!exists) {
        const imported = targetDoc.importNode(node, true);
        imported.setAttribute('PartName', partName);
        targetRoot.appendChild(imported);
      }
      return;
    }

    if (node.localName === 'Default') {
      const ext = (node.getAttribute('Extension') || '').toLowerCase();
      const exists = [...targetRoot.children].some((targetNode) => (
        targetNode.localName === 'Default' && (targetNode.getAttribute('Extension') || '').toLowerCase() === ext
      ));
      if (!exists) {
        targetRoot.appendChild(targetDoc.importNode(node, true));
      }
    }
  });

  [...targetRoot.getElementsByTagName('Override')].forEach((node) => {
    const partName = node.getAttribute('PartName');
    const normalized = normalizeContentTypePartName(partName);
    if (normalized !== partName) node.setAttribute('PartName', normalized);
  });

  targetZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(targetDoc));
}

function readStyleSummary(stylesXml) {
  if (!stylesXml) return [];
  const lines = [];
  const doc = parseXml(stylesXml, '模板样式');

  const defaultsNode = doc.getElementsByTagNameNS('*', 'docDefaults')[0];
  if (defaultsNode) {
    const rFonts = defaultsNode.getElementsByTagNameNS('*', 'rFonts')[0];
    const sizeNode = defaultsNode.getElementsByTagNameNS('*', 'sz')[0];
    const spacingNode = defaultsNode.getElementsByTagNameNS('*', 'spacing')[0];

    const font = rFonts?.getAttribute('w:eastAsia') || rFonts?.getAttribute('w:ascii') || rFonts?.getAttribute('eastAsia') || rFonts?.getAttribute('ascii');
    const sizeHalf = sizeNode?.getAttribute('w:val') || sizeNode?.getAttribute('val');
    const spacing = spacingNode?.getAttribute('w:line') || spacingNode?.getAttribute('line');

    if (font) lines.push(`默认字体：${font}`);
    if (sizeHalf) lines.push(`默认字号：${Number(sizeHalf) / 2} 磅`);
    if (spacing) lines.push(`默认行距：${spacing}（twips）`);
  }

  const tableStyles = uniq(
    [...doc.getElementsByTagNameNS('*', 'style')]
      .filter((node) => (node.getAttribute('w:type') || node.getAttribute('type')) === 'table')
      .map((node) => {
        const nameNode = node.getElementsByTagNameNS('*', 'name')[0];
        return nameNode?.getAttribute('w:val') || nameNode?.getAttribute('val') || node.getAttribute('w:styleId') || node.getAttribute('styleId');
      })
  );
  if (tableStyles.length) lines.push(`表格样式：${tableStyles.slice(0, 5).join('、')}${tableStyles.length > 5 ? ' 等' : ''}`);

  const captionStyles = uniq(
    [...doc.getElementsByTagNameNS('*', 'style')]
      .filter((node) => (node.getAttribute('w:type') || node.getAttribute('type')) === 'paragraph')
      .map((node) => {
        const styleId = (node.getAttribute('w:styleId') || node.getAttribute('styleId') || '').toLowerCase();
        const nameNode = node.getElementsByTagNameNS('*', 'name')[0];
        const styleName = (nameNode?.getAttribute('w:val') || nameNode?.getAttribute('val') || '').toLowerCase();
        if (/(caption|图题|表题|标题)/.test(styleId) || /(caption|图题|表题|标题)/.test(styleName)) {
          return nameNode?.getAttribute('w:val') || nameNode?.getAttribute('val') || node.getAttribute('w:styleId') || node.getAttribute('styleId');
        }
        return null;
      })
  );
  if (captionStyles.length) lines.push(`标题样式（图/表）：${captionStyles.slice(0, 5).join('、')}${captionStyles.length > 5 ? ' 等' : ''}`);

  return lines;
}

function readLayoutSummary(documentXml) {
  if (!documentXml) return [];
  const lines = [];
  const doc = parseXml(documentXml, '模板正文');
  const body = doc.getElementsByTagNameNS('*', 'body')[0];
  const sectPr = findBodySectPr(body);
  if (!sectPr) return lines;

  const pgSz = sectPr.getElementsByTagNameNS('*', 'pgSz')[0];
  const pgMar = sectPr.getElementsByTagNameNS('*', 'pgMar')[0];

  if (pgSz) {
    const w = pgSz.getAttribute('w:w') || pgSz.getAttribute('w');
    const h = pgSz.getAttribute('w:h') || pgSz.getAttribute('h');
    if (w && h) lines.push(`页面尺寸（twips）：宽 ${w} × 高 ${h}`);
  }

  if (pgMar) {
    const top = pgMar.getAttribute('w:top') || pgMar.getAttribute('top');
    const right = pgMar.getAttribute('w:right') || pgMar.getAttribute('right');
    const bottom = pgMar.getAttribute('w:bottom') || pgMar.getAttribute('bottom');
    const left = pgMar.getAttribute('w:left') || pgMar.getAttribute('left');
    lines.push(`页边距（twips）：上 ${top || '-'}，右 ${right || '-'}，下 ${bottom || '-'}，左 ${left || '-'}`);
  }

  return lines;
}

function hasPageField(xml) {
  if (!xml) return false;
  const doc = parseXml(xml, '页眉页脚');
  const instrText = [...doc.getElementsByTagNameNS('*', 'instrText')].some((node) => /\bPAGE\b/i.test(node.textContent || ''));
  const fldSimple = [...doc.getElementsByTagNameNS('*', 'fldSimple')].some((node) => /\bPAGE\b/i.test(node.getAttribute('w:instr') || node.getAttribute('instr') || ''));
  return instrText || fldSimple;
}

async function extractTemplateSummary(templateZip) {
  const [stylesXml, documentXml] = await Promise.all([
    getTextFromZip(templateZip, 'word/styles.xml'),
    getTextFromZip(templateZip, 'word/document.xml')
  ]);

  const lines = [];
  lines.push(...readStyleSummary(stylesXml));
  lines.push(...readLayoutSummary(documentXml));

  const headerFooterFiles = Object.keys(templateZip.files).filter((path) => /^word\/(header|footer)\d+\.xml$/.test(path));
  if (headerFooterFiles.length) {
    lines.push(`页眉页脚：共识别 ${headerFooterFiles.length} 个相关部件`);
    const xmlList = await Promise.all(headerFooterFiles.map((path) => getTextFromZip(templateZip, path)));
    if (xmlList.some((xml) => hasPageField(xml))) lines.push('页码：识别到 PAGE 字段');
  }

  return uniq(lines);
}

async function syncHeaderFooter(templateZip, targetZip) {
  const [templateRelsXml, targetRelsXml] = await Promise.all([
    getTextFromZip(templateZip, 'word/_rels/document.xml.rels'),
    getTextFromZip(targetZip, 'word/_rels/document.xml.rels')
  ]);

  if (!templateRelsXml || !targetRelsXml) {
    throw new Error('缺少 document.xml.rels，无法同步页眉页脚。');
  }

  const templateRelsDoc = parseXml(templateRelsXml, '模板文档关系');
  const targetRelsDoc = parseXml(targetRelsXml, '待转换文档关系');
  const templateRels = [...templateRelsDoc.getElementsByTagNameNS('*', 'Relationship')]
    .filter((item) => /\/(header|footer)$/.test(item.getAttribute('Type') || ''));

  const targetRelsRoot = targetRelsDoc.getElementsByTagNameNS('*', 'Relationships')[0];
  if (!targetRelsRoot) throw new Error('待转换文档关系结构异常。');

  [...targetRelsRoot.getElementsByTagNameNS('*', 'Relationship')]
    .filter((item) => /\/(header|footer)$/.test(item.getAttribute('Type') || ''))
    .forEach((item) => item.remove());

  const copiedParts = new Set();
  for (const rel of templateRels) {
    const relId = rel.getAttribute('Id');
    const relType = rel.getAttribute('Type');
    const relTarget = normalizeRelationshipTarget(rel.getAttribute('Target'));
    if (!relId || !relType || !relTarget) continue;

    const conflicted = [...targetRelsRoot.getElementsByTagNameNS('*', 'Relationship')].find((item) => item.getAttribute('Id') === relId);
    if (conflicted) conflicted.remove();

    const importedRel = targetRelsDoc.importNode(rel, true);
    importedRel.setAttribute('Target', relTarget);
    targetRelsRoot.appendChild(importedRel);

    const partPath = normalizePartPath('word/document.xml', relTarget);
    await copyPartWithDependencies(templateZip, targetZip, partPath, copiedParts);
  }

  await mergeContentTypes(templateZip, targetZip, copiedParts);

  if (copiedParts.size > 0) {
    targetZip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(targetRelsDoc));
  }
}

async function handleTemplateSelected() {
  const templateFile = templateInput.files[0];
  if (!templateFile) {
    setFormatSummary([], '');
    return;
  }

  try {
    const templateZip = await loadDocxZip(templateFile, '模板文档');
    const lines = await extractTemplateSummary(templateZip);
    setFormatSummary(lines, 'success');
  } catch (error) {
    console.error(error);
    setFormatSummary([`模板解析失败：${error.message || '文件不可读取'}`], 'error');
  }
}

async function convertDocx() {
  const templateFile = templateInput.files[0];
  const targetFile = targetInput.files[0];

  if (!window.JSZip) {
    setStatus('转换组件未加载成功，请刷新页面重试。', 'error');
    return;
  }

  if (!templateFile || !targetFile) {
    setStatus('请先选择模板文档和待转换文档。', 'error');
    return;
  }

  const options = getCheckedOptions();
  setStatus('正在处理文档，请稍候...');
  convertBtn.disabled = true;

  try {
    const [templateZip, targetZip] = await Promise.all([
      loadDocxZip(templateFile, '模板文档'),
      loadDocxZip(targetFile, '待转换文档')
    ]);

    if (options.has('styles')) {
      for (const path of STYLE_FILES) await copyIfExists(templateZip, targetZip, path);
    }

    if (options.has('theme')) {
      for (const path of THEME_FILES) await copyIfExists(templateZip, targetZip, path);
    }

    if (options.has('headerFooter')) {
      await syncHeaderFooter(templateZip, targetZip);
    }

    if (options.has('table')) {
      await copyIfExists(templateZip, targetZip, 'word/stylesWithEffects.xml');
      await copyIfExists(templateZip, targetZip, 'word/settings.xml');
    }

    if (options.has('pageLayout') || options.has('headerFooter')) {
      const [templateDocXml, targetDocXml] = await Promise.all([
        getTextFromZip(templateZip, 'word/document.xml'),
        getTextFromZip(targetZip, 'word/document.xml')
      ]);
      if (templateDocXml && targetDocXml) {
        targetZip.file('word/document.xml', replaceSectPr(templateDocXml, targetDocXml));
      }
    }

    const outputBlob = await targetZip.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    const targetName = targetFile.name.replace(/\.docx$/i, '');
    const blobUrl = URL.createObjectURL(outputBlob);

    link.href = blobUrl;
    link.download = `${targetName}_formatted.docx`;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => URL.revokeObjectURL(blobUrl), 1500);
    setStatus('转换完成，已开始下载新文档。', 'success');
  } catch (error) {
    console.error(error);
    setStatus(`转换失败：${error.message || '请确认 DOCX 文件可读取。'}`, 'error');
  } finally {
    convertBtn.disabled = false;
  }
}

templateInput.addEventListener('change', handleTemplateSelected);
convertBtn.addEventListener('click', convertDocx);
