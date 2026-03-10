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

function replaceSectPr(templateDocXml, targetDocXml) {
  const sectPrRegex = /<w:sectPr[\s\S]*?<\/w:sectPr>/;
  const templateSectPr = templateDocXml.match(sectPrRegex);
  if (!templateSectPr) return targetDocXml;

  if (sectPrRegex.test(targetDocXml)) {
    return targetDocXml.replace(sectPrRegex, templateSectPr[0]);
  }

  return targetDocXml.replace('</w:body>', `${templateSectPr[0]}</w:body>`);
}

async function syncHeaderFooter(templateZip, targetZip) {
  const templateFiles = Object.keys(templateZip.files);
  const headerFooterPaths = templateFiles.filter((path) =>
    /^word\/(header|footer)\d+\.xml$/.test(path)
  );

  for (const path of headerFooterPaths) {
    await copyIfExists(templateZip, targetZip, path);
  }

  await copyIfExists(templateZip, targetZip, 'word/_rels/document.xml.rels');
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
