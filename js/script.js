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
    header: 1, // 헤더로 첫 번째 행을 사용
    raw: false,
    // dateNF: "yyyy.mm.dd hh:mm AM/PM",
    cellDates: true, // 날짜 데이터를 Date 객체로 변환
  }); // 헤더가 있는 데이터 추출
  console.log("Data", data);

  // 헤더 분리
  const dataHeader = data.slice(0, 1);
  console.log("dataHeader", dataHeader);

  // 데이터 처리 시작 (헤더 제외)
  const formatData = data.slice(1).map((row, index) => {
    let dataCell = row[0]; // '주문일시'가 A열에 위치
    if (!(dataCell instanceof Date)) {
      row[0] = parseDateString(dataCell);
    } else {
      row[0] = formatDateString(dataCell);
    }
    // console.log(`Row ${index + 2} formattedDate: `, dataCell); // 날짜 데이터 변환 로깅
    return row;
  });
  console.log("formatData", formatData);

  // 헤더 + 포맷데이터, 매입처별 매핑을 위해서 합쳐야함.
  const formattedData = [...dataHeader, ...formatData];
  console.log("formattedData", formattedData);

  const supplierData = categorizeBySupplier(formattedData); // 여기서 데이터 분류
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
    const newSheet = XLSX.utils.json_to_sheet(
      mappedData,
      //   {
      //   dateNF: "yyyy.mm.dd hh:mm AM/PM",
      //   cellDates: true,
      // }
    );
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
    const uploadData = new Uint8Array(e.target.result);
    const workbook = XLSX.read(uploadData, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, {
      header: 1, // 헤더로 첫 번째 행을 사용
      raw: false,
      // dateNF: "yyyy.mm.dd hh:mm AM/PM",
      cellDates: true, // 날짜 데이터를 Date 객체로 변환
    }); // 헤더가 있는 데이터 추출
    console.log("Data", data);

    // 헤더 분리
    const dataHeader = data.slice(0, 1);
    console.log("dataHeader", dataHeader);

    // 데이터 처리 시작 (헤더 제외)
    const formatData = data.slice(1).map((row, index) => {
      let dataCell = row[0]; // '주문일시'가 A열에 위치
      if (!(dataCell instanceof Date)) {
        row[0] = parseDateString(dataCell);
      } else {
        row[0] = formatDateString(dataCell);
      }
      // console.log(`Row ${index + 2} formattedDate: `, dataCell); // 날짜 데이터 변환 로깅
      return row;
    });
    console.log("formatData", formatData);

    // 헤더 + 포맷데이터
    const formattedData = [...dataHeader, ...formatData];
    console.log("formattedData", formattedData);

    const supplierData = categorizeBySupplier(formattedData); // 여기서 데이터 분류
    console.log("supplierData", supplierData);

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

// 문자열 -> 날짜 형식으로 변환
// 날짜 형식 "MM/DD/YY"를 "YYYY-MM-DD"로 변환
function parseDateString(dateStr) {
  if (typeof dateStr === "string") {
    // 날짜와 시간을 공백으로 분리
    const [datePart, timePart] = dateStr.split(" ");
    let year, month, day;

    // 날짜 처리
    if (datePart.includes("/")) {
      // 날짜가 "MM/DD/YY" 형식인 경우
      const parts = datePart.split("/");
      if (parts.length === 3) {
        year = parseInt(parts[2], 10);
        year += year < 50 ? 2000 : 1900; // YY -> YYYY 변환, 50을 기준으로 2000 또는 1900을 더함(50년 이후는 모름.. 두자리 데이터의 한계...)
        month = parts[0].padStart(2, "0");
        day = parts[1].padStart(2, "0");
      }
    } else if (datePart.includes(".")) {
      // 날짜가 "YYYY.MM.DD" 형식인 경우
      const parts = datePart.split(".");
      if (parts.length === 3) {
        year = parts[0];
        month = parts[1].padStart(2, "0");
        day = parts[2].padStart(2, "0");
      }
    }

    // 시간 처리
    if (timePart) {
      let [hours, minutes] = timePart.split(":");
      // 시간 정규화 (24시간제 지원)
      if (timePart.toLowerCase().includes("pm") && hours !== "12") {
        hours = (parseInt(hours, 10) + 12).toString();
      }
      hours = hours.padStart(2, "0");
      minutes = minutes.replace(/[^0-9]/g, "").padStart(2, "0"); // AM/PM 문자 제거 및 포맷

      return `${year}-${month}-${day} ${hours}:${minutes}`;
    } else {
      return `${year}-${month}-${day}`;
    }
  }
  return dateStr; // 변환할 수 없는 형식은 원래 값을 반환
}

// 날짜 형식으로 들어올때, 원하는 날짜 형식으로 포맷팅
function formatDateString(date) {
  if (!(date instanceof Date)) {
    console.error("Expected a Date instance, received:", date);
    return date; // 만약 Date 인스턴스가 아니라면 원래 값을 반환
  }

  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, "0");
  const day = date.getDate().toString().padStart(2, "0");
  const hours = date.getHours().toString().padStart(2, "0");
  const minutes = date.getMinutes().toString().padStart(2, "0");

  return `${year}-${month}-${day} ${hours}:${minutes}`;
}
