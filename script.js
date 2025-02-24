const container = document.getElementById('canvas-container');

// 設置名片的大小
const cardWidth = 736; // 92mm * 8
const cardHeight = 448; // 56mm * 8

// 設置統一的姓名寬度
const nameWidth = cardWidth / 4;

// 繪製名片
function drawCard(ctx, name, englishName, chineseTitle, englishTitle, department, departmentEn, phone, phoneIcon) {
  // 繪製名片背景
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cardWidth, cardHeight);

  // 繪製姓名
  ctx.fillStyle = '#000000';
  ctx.font = '80px "Microsoft YaHei UI"';
  ctx.textAlign = 'left';
  ctx.fillText(name, 40, 120);

  // 繪製英文名
  ctx.font = '40px "Microsoft YaHei UI"';
  ctx.fillText(englishName, 40 + ctx.measureText(name).width + 20, 120);

  // 繪製中文職稱
  ctx.font = '60px "Microsoft YaHei UI"';
  ctx.fillText(chineseTitle, 40, 200);

  // 繪製英文職稱
  ctx.font = '30px "Microsoft YaHei UI"';
  ctx.fillText(englishTitle, 40 + ctx.measureText(chineseTitle).width + 20, 200);

  // 繪製部門1和部門2，中間加上「/」
  ctx.font = '20px "Microsoft YaHei UI"';
  ctx.fillText(`${department} / ${departmentEn}`, 40, 280);

  // 繪製電話 icon
  const svg = new Blob([phoneIcon], { type: 'image/svg+xml' });
  const url = URL.createObjectURL(svg);
  const img = new Image();
  img.onload = () => {
    ctx.drawImage(img, 40, 330, 20, 20);
    URL.revokeObjectURL(url);

    // 繪製電話
    ctx.font = '20px "Microsoft YaHei UI"';
    ctx.fillText(phone, 70, 360);
  };
  img.src = url;
}

// 讀取 Excel 文件
function readExcel(file) {
  return fetch(file)
    .then(response => response.arrayBuffer())
    .then(data => XLSX.read(data, { type: 'array' }));
}

// 根據工作地點找到對應的電話
function findPhoneByLocation(positionRows, location) {
  const positionRow = positionRows.find(posRow => posRow[0] === location);
  return positionRow ? positionRow[7] : '';
}

// 處理每一行數據
function processRow(row, positionRows, phoneIcon) {
  const name = row[3]; // D欄位（姓名）
  const englishName = row[4]; // E欄位（英文名）
  const department = row[5]; // F欄位（部門1）
  const departmentEn = row[6]; // G欄位（部門2）
  const chineseTitle = row[7]; // H欄位（中文職稱）
  const englishTitle = row[8]; // I欄位（英文職稱）
  const location = row[1]; // B欄位（工作地點）

  const phone = findPhoneByLocation(positionRows, location);

  if (name && englishName && department && departmentEn && chineseTitle && englishTitle && phone) {
    // 創建新的 canvas 元素
    const canvas = document.createElement('canvas');
    canvas.width = cardWidth;
    canvas.height = cardHeight;
    container.appendChild(canvas);

    const ctx = canvas.getContext('2d');
    drawCard(ctx, name, englishName, chineseTitle, englishTitle, department, departmentEn, phone, phoneIcon);
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

  // 電話 icon 的 SVG
  const phoneIcon = `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
      <path d="M164.9 24.6c-7.7-18.6-28-28.5-47.4-23.2l-88 24C12.1 30.2 0 46 0 64C0 311.4 200.6 512 448 512c18 0 33.8-12.1 38.6-29.5l24-88c5.3-19.4-4.6-39.7-23.2-47.4l-96-40c-16.3-6.8-35.2-2.1-46.3 11.6L304.7 368C234.3 334.7 177.3 277.7 144 207.3L193.3 167c13.7-11.2 18.4-30 11.6-46.3l-40-96z"/>
    </svg>
  `;

  // 忽略第一行，從第二行開始讀取
  dataRows.slice(1).forEach(row => processRow(row, positionRows, phoneIcon));
}

main();
