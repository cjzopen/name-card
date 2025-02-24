const container = document.getElementById('svg-container');

// 設置名片的大小
const cardWidth = 736; // 92mm * 8
const cardHeight = 448; // 56mm * 8

const gap = 10;
const startHeight = 120;

// 定義字體大小變數
const fontSize = {
  name: 40,
  englishName: 30,
  chineseTitle: 24,
  englishTitle: 24,
  department: 24,
  phone: 18,
  email: 18,
  address: 18
};

// 定義電話、手機和 email 的 SVG icon
const phoneIcon = `
  <svg width="20" height="20" viewBox="0 0 24 24">
    <path d="M6.62 10.79a15.053 15.053 0 006.59 6.59l2.2-2.2a1.003 1.003 0 011.11-.27c1.12.45 2.33.69 3.58.69.55 0 1 .45 1 1v3.5c0 .55-.45 1-1 1C10.29 21 3 13.71 3 4.5c0-.55.45-1 1-1H7.5c.55 0 1 .45 1 1 0 1.25.24 2.46.69 3.58.14.34.07.73-.27 1.11l-2.2 2.2z"/>
  </svg>
`;

const mobileIcon = `
  <svg width="20" height="20" viewBox="0 0 320 512">
    <path d="M0 64C0 28.7 28.7 0 64 0L256 0c35.3 0 64 28.7 64 64l0 384c0 35.3-28.7 64-64 64L64 512c-35.3 0-64-28.7-64-64L0 64zm64 96l0 64c0 17.7 14.3 32 32 32l128 0c17.7 0 32-14.3 32-32l0-64c0-17.7-14.3-32-32-32L96 128c-17.7 0-32 14.3-32 32zM80 352a24 24 0 1 0 0-48 24 24 0 1 0 0 48zm24 56a24 24 0 1 0 -48 0 24 24 0 1 0 48 0zm56-56a24 24 0 1 0 0-48 24 24 0 1 0 0 48zm24 56a24 24 0 1 0 -48 0 24 24 0 1 0 48 0zm56-56a24 24 0 1 0 0-48 24 24 0 1 0 0 48zm24 56a24 24 0 1 0 -48 0 24 24 0 1 0 48 0zM128 48c-8.8 0-16 7.2-16 16s7.2 16 16 16l64 0c8.8 0 16-7.2 16-16s-7.2-16-16-16l-64 0z"/>
  </svg>
`;

const emailIcon = `
  <svg width="20" height="20" viewBox="0 0 512 512">
    <path d="M64 112c-8.8 0-16 7.2-16 16l0 22.1L220.5 291.7c20.7 17 50.4 17 71.1 0L464 150.1l0-22.1c0-8.8-7.2-16-16-16L64 112zM48 212.2L48 384c0 8.8 7.2 16 16 16l384 0c8.8 0 16-7.2 16-16l0-171.8L322 328.8c-38.4 31.5-93.7 31.5-132 0L48 212.2zM0 128C0 92.7 28.7 64 64 64l384 0c35.3 0 64 28.7 64 64l0 256c0 35.3-28.7 64-64 64L64 448c-35.3 0-64-28.7-64-64L0 128z"/>
  </svg>
`;

// 繪製名片
function drawCard(svg, name, englishName, chineseTitle, englishTitle, department, departmentEn, phone, ext, mobile, email, addressCn, addressEn1, addressEn2) {
  // 繪製姓名
  svg.innerHTML += `<text x="40" y="${startHeight}" font-size="${fontSize.name}" fill="#000">${name}</text>`;
  svg.innerHTML += `<text x="${40 + name.length * fontSize.name + gap}" y="${startHeight}" font-size="${fontSize.englishName}" fill="#000">${englishName}</text>`;

  // 繪製中文職稱
  svg.innerHTML += `<text x="40" y="${startHeight + fontSize.name}" font-size="${fontSize.chineseTitle}" fill="#000">${chineseTitle}</text>`;
  svg.innerHTML += `<text x="${40 + chineseTitle.length * fontSize.chineseTitle + gap}" y="${startHeight + fontSize.name}" font-size="${fontSize.englishTitle}" fill="#000">${englishTitle}</text>`;

  // 繪製部門1和部門2，中間加上「/」
  svg.innerHTML += `<text x="40" y="280" font-size="${fontSize.department}" fill="#000">${department} / ${departmentEn}</text>`;

  // 繪製電話 icon 和電話
  svg.innerHTML += `<g transform="translate(40, 342)">${phoneIcon}</g>`;
  svg.innerHTML += `<text x="70" y="360" font-size="${fontSize.phone}" fill="#000">${phone} ext. ${ext}</text>`;
  svg.innerHTML += `<g transform="translate(330, 342)">${mobileIcon}</g>`;
  svg.innerHTML += `<text x="360" y="360" font-size="${fontSize.phone}" fill="#000">${mobile}</text>`;

  // 繪製 email icon 和 email
  svg.innerHTML += `<g transform="translate(40, 382)">${emailIcon}</g>`;
  svg.innerHTML += `<text x="70" y="400" font-size="${fontSize.email}" fill="#000">${email}</text>`;

  // 繪製地址
  svg.innerHTML += `<text x="40" y="440" font-size="${fontSize.address}" fill="#000">${addressCn}</text>`;
  svg.innerHTML += `<text x="40" y="460" font-size="${fontSize.address}" fill="#000">${addressEn1}</text>`;
  svg.innerHTML += `<text x="40" y="480" font-size="${fontSize.address}" fill="#000">${addressEn2}</text>`;
}

// 讀取 Excel 文件
function readExcel(file) {
  return fetch(file)
    .then(response => response.arrayBuffer())
    .then(data => XLSX.read(data, { type: 'array' }));
}

// 根據工作地點找到對應的電話和地址
function findDetailsByLocation(positionRows, location) {
  const positionRow = positionRows.find(posRow => posRow[0] === location);
  return positionRow ? {
    phone: positionRow[7],
    addressCn: positionRow[3],
    addressEn1: positionRow[4],
    addressEn2: positionRow[5]
  } : {};
}

// 處理每一行數據
function processRow(row, positionRows) {
  const name = row[3]; // D欄位（姓名）
  const englishName = row[4]; // E欄位（英文名）
  const department = row[5]; // F欄位（部門1）
  const departmentEn = row[6]; // G欄位（部門2）
  const chineseTitle = row[7]; // H欄位（中文職稱）
  const englishTitle = row[8]; // I欄位（英文職稱）
  const location = row[1]; // B欄位（工作地點）
  const ext = row[12]; // M欄位（分機）
  const mobile = row[13]; // N欄位（手機）
  const email = row[11]; // L欄位（email）

  const { phone, addressCn, addressEn1, addressEn2 } = findDetailsByLocation(positionRows, location);

  if (name && englishName && department && departmentEn && chineseTitle && englishTitle && phone) {
    // 創建新的 svg 元素
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", cardWidth);
    svg.setAttribute("height", cardHeight);
    svg.setAttribute("class", "name-card");
    container.appendChild(svg);

    drawCard(svg, name, englishName, chineseTitle, englishTitle, department, departmentEn, phone, ext, mobile, email, addressCn, addressEn1, addressEn2);
  }
}

// 主函數
async function main() {
  const dataWorkbook = await readExcel('data.xlsx');
  const dataSheet = dataWorkbook.Sheets[dataWorkbook.SheetNames[0]];
  const dataRows = XLSX.utils.sheet_to_json(dataSheet, { header: 1 });

  const positionWorkbook = await readExcel('position.xlsx');
  const positionSheet = positionWorkbook.Sheets[positionWorkbook.SheetNames[0]];
  const positionRows = XLSX.utils.sheet_to_json(positionSheet, { header: 1 });

  // 忽略第一行，從第二行開始讀取
  dataRows.slice(1).forEach(row => processRow(row, positionRows));
}

main();
