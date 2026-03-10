const templateInput = document.getElementById('templateFile');
const targetInput = document.getElementById('targetFile');
const convertBtn = document.getElementById('convertBtn');
const statusEl = document.getElementById('status');

const STYLE_FILES = [
  'word/styles.xml',
  'word/fontTable.xml',
  'word/numbering.xml'
];

const THEME_FILES = [
  'word/theme/theme1.xml'
];

function setStatus(text, type = '') {
  statusEl.textContent = text;
  statusEl.className = `status ${type}`.trim();
}

function getCheckedOptions() {
  return new Set(
    [...document.querySelectorAll('.options input[type="checkbox"]:checked')]
      .map((item) => item.value)
  );
}

function getTextFromZip(zip, path) {
  const file = zip.file(path);
  return file ? file.async('string') : Promise.resolve(null);
}

async function copyIfExists(fromZip, toZip, path) {
  const file = fromZip.file(path);
  if (!file) return;
  const content = await file.async('uint8array');
  toZip.file(path, content);
}

function normalizePartPath(basePath, target) {
  const baseDir = basePath.split('/').slice(0, -1);
  const raw = `${baseDir.join('/')}/${target}`.split('/');
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

function normalizeRelationshipTarget(target) {
  if (!target) return target;
  return target.replace(/(^|\/)setting\.xml(\?|#|$)/i, '$1settings.xml$2');
}

function normalizeContentTypePartName(partName) {
  if (!partName) return partName;
  return partName.replace(/(^|\/)setting\.xml$/i, '$1settings.xml');
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
  if (!templateSectPr) {
    return targetDocXml;
  }

  const targetSectPr = findBodySectPr(targetBody);
  const importedSectPr = targetDoc.importNode(templateSectPr, true);

  if (targetSectPr) {
    targetBody.replaceChild(importedSectPr, targetSectPr);
  } else {
    targetBody.appendChild(importedSectPr);
  }

  return new XMLSerializer().serializeToString(targetDoc);
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
  if (!targetRelsRoot) {
    throw new Error('待转换文档关系结构异常。');
  }

  [...targetRelsRoot.getElementsByTagNameNS('*', 'Relationship')]
    .filter((item) => /\/(header|footer)$/.test(item.getAttribute('Type') || ''))
    .forEach((item) => item.remove());

  const copiedParts = new Set();
  for (const rel of templateRels) {
    const relId = rel.getAttribute('Id');
    const relType = rel.getAttribute('Type');
    const relTarget = normalizeRelationshipTarget(rel.getAttribute('Target'));
    if (!relId || !relType || !relTarget) continue;

    const conflicted = [...targetRelsRoot.getElementsByTagNameNS('*', 'Relationship')]
      .find((item) => item.getAttribute('Id') === relId);
    if (conflicted) conflicted.remove();

    const importedRel = targetRelsDoc.importNode(rel, true);
    importedRel.setAttribute('Target', relTarget);
    targetRelsRoot.appendChild(importedRel);

    const partPath = normalizePartPath('word/document.xml', relTarget);
    copiedParts.add(partPath);
    await copyIfExists(templateZip, targetZip, partPath);

    const partName = partPath.split('/').pop();
    const partRelsPath = `word/_rels/${partName}.rels`;
    const partRelsXml = await getTextFromZip(templateZip, partRelsPath);
    if (partRelsXml) {
      const partRelsDoc = parseXml(partRelsXml, `${partName} 关系`);
      [...partRelsDoc.getElementsByTagNameNS('*', 'Relationship')].forEach((item) => {
        const normalizedTarget = normalizeRelationshipTarget(item.getAttribute('Target'));
        if (normalizedTarget) item.setAttribute('Target', normalizedTarget);
      });
      targetZip.file(partRelsPath, new XMLSerializer().serializeToString(partRelsDoc));
    }
  }

  const templateFiles = Object.keys(templateZip.files);
  const dependentParts = templateFiles.filter((path) =>
    /^word\/(media|embeddings|drawings)\//.test(path)
  );
  for (const path of dependentParts) {
    await copyIfExists(templateZip, targetZip, path);
  }

  const [templateContentTypesXml, targetContentTypesXml] = await Promise.all([
    getTextFromZip(templateZip, '[Content_Types].xml'),
    getTextFromZip(targetZip, '[Content_Types].xml')
  ]);

  if (templateContentTypesXml && targetContentTypesXml) {
    const templateContentTypesDoc = parseXml(templateContentTypesXml, '模板内容类型');
    const targetContentTypesDoc = parseXml(targetContentTypesXml, '待转换内容类型');
    const templateRoot = templateContentTypesDoc.getElementsByTagName('Types')[0];
    const targetRoot = targetContentTypesDoc.getElementsByTagName('Types')[0];

    if (templateRoot && targetRoot) {
      const shouldKeep = (node) => {
        if (node.localName === 'Override') {
          const partName = node.getAttribute('PartName') || '';
          return /\/word\/(header|footer)\d+\.xml$/.test(partName);
        }
        if (node.localName === 'Default') {
          const ext = (node.getAttribute('Extension') || '').toLowerCase();
          return ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'wmf', 'emf', 'rels', 'xml', 'bin'].includes(ext);
        }
        return false;
      };

      [...templateRoot.children].filter(shouldKeep).forEach((node) => {
        const key = node.localName === 'Override'
          ? `${node.localName}:${normalizeContentTypePartName(node.getAttribute('PartName'))}`
          : `${node.localName}:${node.getAttribute('Extension')}`;
        const exists = [...targetRoot.children].some((targetNode) => {
          const targetKey = targetNode.localName === 'Override'
            ? `${targetNode.localName}:${normalizeContentTypePartName(targetNode.getAttribute('PartName'))}`
            : `${targetNode.localName}:${targetNode.getAttribute('Extension')}`;
          return targetKey === key;
        });
        if (!exists) {
          const importedNode = targetContentTypesDoc.importNode(node, true);
          if (importedNode.localName === 'Override') {
            importedNode.setAttribute(
              'PartName',
              normalizeContentTypePartName(importedNode.getAttribute('PartName'))
            );
          }
          targetRoot.appendChild(importedNode);
        }
      });

      [...targetRoot.getElementsByTagName('Override')].forEach((node) => {
        const partName = node.getAttribute('PartName');
        const normalizedPartName = normalizeContentTypePartName(partName);
        if (normalizedPartName !== partName) {
          node.setAttribute('PartName', normalizedPartName);
        }
      });
    }

    targetZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(targetContentTypesDoc));
  }

  if (copiedParts.size > 0) {
    targetZip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(targetRelsDoc));
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
      for (const path of STYLE_FILES) {
        await copyIfExists(templateZip, targetZip, path);
      }
    }

    if (options.has('theme')) {
      for (const path of THEME_FILES) {
        await copyIfExists(templateZip, targetZip, path);
      }
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

convertBtn.addEventListener('click', convertDocx);
