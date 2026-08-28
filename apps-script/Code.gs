const SHEET_ID = '1HcWx845mSpfJnwziugByJLn0Npc1aBSomOBa6esQWVM';
const GITHUB_REPOSITORY = 'Dougster00/TMG-Product-Catalogue';
const GITHUB_BRANCH = 'main';
const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TMG Secure Image Admin')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function uploadAndUpdate(payload) {
  verifyOwner_();

  if (!payload || !payload.productName || !payload.filename || !payload.base64 || !payload.mimeType) {
    throw new Error('Product name, filename and image are required.');
  }

  const productName = String(payload.productName).trim();
  const filename = safeFilename_(payload.filename);
  const mimeType = String(payload.mimeType).toLowerCase();
  if (!/^(image\/jpeg|image\/png|image\/webp)$/.test(mimeType)) {
    throw new Error('Only JPG, PNG and WebP images are supported.');
  }

  const imageBytes = Utilities.base64Decode(String(payload.base64));
  if (imageBytes.length > MAX_IMAGE_BYTES) {
    throw new Error('The image is larger than 10 MB.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const match = findProduct_(productName);
    uploadToGitHub_(filename, imageBytes);
    match.sheet.getRange(match.row, match.imageColumn).setValue(filename);
    SpreadsheetApp.flush();
    return {
      ok: true,
      message: 'Uploaded ' + filename + ' and updated ' + productName + '.',
      filename: filename,
      row: match.row
    };
  } finally {
    lock.releaseLock();
  }
}

function verifyOwner_() {
  const props = PropertiesService.getScriptProperties();
  const allowedEmail = String(props.getProperty('ALLOWED_EMAIL') || '').trim().toLowerCase();
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!allowedEmail) throw new Error('ALLOWED_EMAIL has not been configured.');
  if (!activeEmail || activeEmail !== allowedEmail) throw new Error('You are not authorised to use this uploader.');
  if (!props.getProperty('GITHUB_TOKEN')) throw new Error('GITHUB_TOKEN has not been configured.');
}

function findProduct_(productName) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('The Google Sheet is empty.');

  const headers = values[0].map(function(value) {
    return String(value).trim().toUpperCase();
  });
  let productColumn = headers.indexOf('PRODUCT') + 1;
  if (!productColumn) productColumn = headers.indexOf('PRODUCT NAME') + 1;
  const imageColumn = headers.indexOf('IMAGE') + 1;
  if (!productColumn || !imageColumn) {
    throw new Error('The sheet must contain PRODUCT and IMAGE column headings.');
  }

  const wanted = productName.toLowerCase();
  const rows = [];
  for (let row = 2; row <= values.length; row++) {
    if (String(values[row - 1][productColumn - 1]).trim().toLowerCase() === wanted) rows.push(row);
  }
  if (!rows.length) throw new Error('No exact product match was found for "' + productName + '".');
  if (rows.length > 1) throw new Error('More than one product matches "' + productName + '". Nothing was changed.');
  return { sheet: sheet, row: rows[0], imageColumn: imageColumn };
}

function uploadToGitHub_(filename, imageBytes) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const encodedPath = filename.split('/').map(encodeURIComponent).join('/');
  const url = 'https://api.github.com/repos/' + GITHUB_REPOSITORY + '/contents/' + encodedPath;

  const existing = UrlFetchApp.fetch(url + '?ref=' + encodeURIComponent(GITHUB_BRANCH), {
    method: 'get',
    headers: githubHeaders_(token),
    muteHttpExceptions: true
  });
  if (existing.getResponseCode() === 200) {
    throw new Error('A GitHub image named "' + filename + '" already exists. Choose another filename.');
  }
  if (existing.getResponseCode() !== 404) {
    throw new Error('GitHub could not check the filename: ' + githubError_(existing));
  }

  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: githubHeaders_(token),
    payload: JSON.stringify({
      message: 'Add catalogue image for ' + filename,
      content: Utilities.base64Encode(imageBytes),
      branch: GITHUB_BRANCH
    }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 201) {
    throw new Error('GitHub upload failed: ' + githubError_(response));
  }
}

function githubHeaders_(token) {
  return {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'TMG-Catalogue-Admin'
  };
}

function githubError_(response) {
  try {
    const parsed = JSON.parse(response.getContentText());
    return parsed.message || ('HTTP ' + response.getResponseCode());
  } catch (error) {
    return 'HTTP ' + response.getResponseCode();
  }
}

function safeFilename_(value) {
  const filename = String(value).trim().toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '');
  if (!/^[a-z0-9][a-z0-9._-]*\.(jpg|jpeg|png|webp)$/.test(filename)) {
    throw new Error('The filename must end in .jpg, .jpeg, .png or .webp.');
  }
  return filename;
}
