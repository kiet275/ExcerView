const excelFileInput = document.getElementById("excelFile");
const searchInput = document.getElementById("search");
const tableContainer = document.getElementById("tableContainer");
const fileHistoryList = document.getElementById("fileHistory");

let originalData = [];
let backupData = [];

excelFileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const arrayBuffer = e.target.result;
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    originalData = json;
    backupData = JSON.parse(JSON.stringify(json));
    renderTable(originalData);

    saveToHistory(file.name, arrayBuffer);
    loadHistory();
  };

  reader.readAsArrayBuffer(file);
});

function renderTable(data) {
  let html = '<table border="1" cellspacing="0" cellpadding="5">';
  data.forEach((row, rowIndex) => {
    html += "<tr>";
    row.forEach((cell, colIndex) => {
      const tag = rowIndex === 0 ? "th" : "td";
      const content = cell !== undefined ? cell : "";
      const editable = rowIndex === 0 ? "" : 'contenteditable="true"';
      let isEdited = "";
      if (
        rowIndex > 0 &&
        backupData[rowIndex] &&
        backupData[rowIndex][colIndex] !== content
      ) {
        isEdited = 'class="edited"';
      }
      html += `<${tag} ${editable} data-row="${rowIndex}" data-col="${colIndex}" ${isEdited}>${content}</${tag}>`;
    });
    html += "</tr>";
  });
  html += "</table>";
  tableContainer.innerHTML = html;
}

// Ghi dữ liệu khi chỉnh sửa (chỉ cập nhật dữ liệu, không render lại)
tableContainer.addEventListener("input", (e) => {
  if (e.target.tagName.toLowerCase() === "td") {
    const row = parseInt(e.target.getAttribute("data-row"));
    const col = parseInt(e.target.getAttribute("data-col"));
    const value = e.target.innerText;
    originalData[row][col] = value;
  }
});

// Đánh dấu ô đã chỉnh sửa khi rời ô
tableContainer.addEventListener("focusout", (e) => {
  if (e.target.tagName.toLowerCase() === "td") {
    const row = parseInt(e.target.getAttribute("data-row"));
    const col = parseInt(e.target.getAttribute("data-col"));
    const currentValue = e.target.innerText;
    const originalValue = backupData[row]?.[col] ?? "";

    if (currentValue !== originalValue) {
      e.target.classList.add("edited");
    } else {
      e.target.classList.remove("edited");
    }
  }
});

// Tìm kiếm
searchInput.addEventListener("input", () => {
  const keyword = searchInput.value.toLowerCase();
  if (keyword.trim() === "") {
    renderTable(originalData);
    return;
  }

  const highlightedData = originalData.map((row) =>
    row.map((cell) => {
      if (typeof cell === "string" && cell.toLowerCase().includes(keyword)) {
        return cell.replace(
          new RegExp(`(${keyword})`, "gi"),
          "<mark>$1</mark>"
        );
      }
      return cell;
    })
  );

  renderTable(highlightedData);
});

// Lưu lịch sử
function saveToHistory(fileName, arrayBuffer) {
  const history = JSON.parse(localStorage.getItem("excelHistory")) || [];
  const base64 = arrayBufferToBase64(arrayBuffer);
  const filtered = history.filter((entry) => entry.name !== fileName);
  const updatedHistory = [{ name: fileName, data: base64 }, ...filtered].slice(
    0,
    5
  );
  localStorage.setItem("excelHistory", JSON.stringify(updatedHistory));
}

// Tải lịch sử
function loadHistory() {
  const history = JSON.parse(localStorage.getItem("excelHistory")) || [];
  fileHistoryList.innerHTML = "";
  history.forEach((entry) => {
    const li = document.createElement("li");
    li.textContent = entry.name;
    li.style.cursor = "pointer";
    li.addEventListener("click", () => {
      const arrayBuffer = base64ToArrayBuffer(entry.data);
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
        type: "array",
      });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      originalData = json;
      backupData = JSON.parse(JSON.stringify(json));
      renderTable(originalData);
    });
    fileHistoryList.appendChild(li);
  });
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  for (let b of bytes) {
    binary += String.fromCharCode(b);
  }
  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

// Xuất PDF
document.getElementById("exportPdfBtn").addEventListener("click", () => {
  const element = tableContainer;
  const opt = {
    margin: 0.5,
    filename: "bang_excel.pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: "in", format: "a4", orientation: "portrait" },
  };
  html2pdf().set(opt).from(element).save();
});

// Xuất Word
document.getElementById("exportWordBtn").addEventListener("click", () => {
  const header = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset='utf-8'></head><body>`;
  const footer = "</body></html>";
  const sourceHTML = header + tableContainer.innerHTML + footer;

  const sourceBlob = new Blob(["\ufeff", sourceHTML], {
    type: "application/msword",
  });

  const url = URL.createObjectURL(sourceBlob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "bang_excel.doc";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});

// Tải Excel đã chỉnh sửa với popup
document.getElementById("downloadExcelBtn").addEventListener("click", () => {
  Swal.fire({
    title: "Bạn muốn làm gì?",
    text: "Tải file đã chỉnh sửa hay hoàn tác thay đổi?",
    icon: "question",
    showCancelButton: true,
    confirmButtonText: "Tải về",
    cancelButtonText: "Hoàn tác",
  }).then((result) => {
    if (result.isConfirmed) {
      const worksheet = XLSX.utils.aoa_to_sheet(originalData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      XLSX.writeFile(workbook, "bang_excel_sua.xlsx");
    } else if (result.dismiss === Swal.DismissReason.cancel) {
      originalData = JSON.parse(JSON.stringify(backupData));
      renderTable(originalData);
      Swal.fire("Đã hoàn tác thay đổi", "", "success");
    }
  });
});

// Nút hoàn tác trực tiếp
document.getElementById("undoBtn").addEventListener("click", () => {
  originalData = JSON.parse(JSON.stringify(backupData));
  renderTable(originalData);
  Swal.fire("Đã hoàn tác thay đổi", "", "success");
});

// Tải lịch sử khi mở trang
loadHistory();
