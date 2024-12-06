// 매입처 발주 양식
import { columnHeader, supplierHeaders } from "./mapping.js";

document
  .getElementById("file-input")
  .addEventListener("change", async (event) => {
    const fileInput = document.getElementById("file-input");
    const file = fileInput.files[0];
    if (file) {
      console.log("File input changed!", event.target.files);
    } else {
      alert("파일을 선택해주세요.");
      return;
    }

    // 업로드 파일 표시
    const fileNameDisplay = document.getElementById("file-name");
    fileNameDisplay.textContent = file.name;
  });

document.getElementById("process-btn").addEventListener("click", () => {
  const fileInput = document.getElementById("file-input");
  const file = fileInput.files[0];
  const errorMessage = document.getElementById("error-message");

  if (!file) {
    errorMessage.textContent = "엑셀 파일을 업로드해주세요!";
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      processWorkbook(workbook);
      errorMessage.textContent = ""; // Clear 에러메세지
    } catch (error) {
      errorMessage.textContent = error;
    }
  };
  reader.readAsArrayBuffer(file);
});

// 개별 다운로드
function processWorkbook(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]]; // 첫 번째 시트 사용
  const data = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    dateNF: "yyyy.mm.dd hh:mm AM/PM",
    cellDates: true, // 날짜 데이터를 Date 객체로 변환
  }); // 헤더가 있는 데이터 추출
  console.log("headerData", data);
  const supplierData = categorizeBySupplier(data);
  console.log("supplierData", supplierData);

  const resultList = document.getElementById("result-list");
  resultList.innerHTML = ""; // 기존 리스트 내용을 클리어

  Object.keys(supplierData).forEach((supplier) => {
    const mappedData = mapDataToSupplierFormat(
      supplierData[supplier],
      supplier,
    );
    console.log("mappedData", mappedData);

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(mappedData, {
      dateNF: "yyyy.mm.dd hh:mm AM/PM",
      cellDates: true,
    });
    const fileName = `${supplier}_프라이스잇_발주서.xlsx`;
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, supplier);

    // 매입처 리스트 요소 생성
    // 다운로드 버튼 생성
    const button = document.createElement("button");
    button.textContent = `다운로드`;
    button.addEventListener("click", () => {
      XLSX.writeFile(newWorkbook, fileName);
    });

    // result-item 요소 생성
    const supplyDiv = document.createElement("div");
    supplyDiv.className = "result-item";
    supplyDiv.textContent = fileName;

    // result-item-container 요소 생성
    const supplyDivContainer = document.createElement("div");
    supplyDivContainer.className = "result-item-container";
    supplyDivContainer.textContent = supplier;

    // container 안에 요소 생성
    supplyDivContainer.appendChild(supplyDiv);
    supplyDivContainer.appendChild(button);

    // resultList에 요소 추가
    resultList.appendChild(supplyDivContainer);
  });
}

// 전체 다운로드
document.querySelector(".downloadAll-btn").addEventListener("click", () => {
  downloadAllFiles(); // 모든 매입처 데이터를 처리하고 ZIP 파일로 저장합니다.
});

function downloadAllFiles() {
  const zip = new JSZip(); // ZIP 객체 생성
  const fileInput = document.getElementById("file-input");
  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataHeader = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      dateNF: "yyyy.mm.dd hh:mm AM/PM",
      cellDates: true,
    });
    const supplierData = categorizeBySupplier(dataHeader);

    Object.keys(supplierData).forEach((supplier) => {
      const mappedData = mapDataToSupplierFormat(
        supplierData[supplier],
        supplier,
      );
      const newWorkbook = XLSX.utils.book_new();
      const newSheet = XLSX.utils.json_to_sheet(mappedData, {
        dateNF: "yyyy.mm.dd hh:mm:ss",
        cellDates: true,
      });
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, supplier);

      const wbout = XLSX.write(newWorkbook, {
        bookType: "xlsx",
        type: "binary",
      });
      zip.file(`${supplier}_프라이스잇_발주서.xlsx`, wbout, { binary: true });
    });

    zip.generateAsync({ type: "blob" }).then(function (content) {
      saveAs(content, "프라이스잇_발주서.zip");
    });
  };
  reader.readAsArrayBuffer(file);
}

// 매입처 분류 함수
function categorizeBySupplier(data) {
  const headers = data[0]; // 통합 발주서 column
  console.log("categorizeBySupplier_headers", headers);
  const supplierIndex = headers.findIndex((header) => header === "매입처"); // 매입처 열이 존재해야함.
  console.log("supplierIndex", supplierIndex);
  const categorizedData = {};

  data.slice(1).forEach((row) => {
    const supplier = row[supplierIndex];
    if (!categorizedData[supplier]) {
      categorizedData[supplier] = [];
    }
    categorizedData[supplier].push(row);
  });

  return categorizedData;
}

function mapDataToSupplierFormat(data, supplier) {
  const headers = supplierHeaders[supplier] || columnHeader.common; // mapping.js에 매입처 양식이 없으면 common 양식으로 작성.
  return data.map((row) => {
    const newRow = {};
    Object.keys(headers).forEach((key) => {
      const columnIndex = headers[key].charCodeAt(0) - "A".charCodeAt(0); // 65 - 65 = 0
      newRow[key] =
        typeof row[columnIndex] === "string"
          ? row[columnIndex].trim()
          : row[columnIndex]; // 셀 데이터에서 불필요한 공백 제거
    });
    return newRow;
  });
}
