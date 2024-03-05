const fs = require("fs");
const xlsx = require("xlsx");
const csv = require("csv");
const { createObjectCsvWriter } = require("csv-writer");

// 회사 정보 입력
const companyList = [
  { key: 3, co_name: "㈜한화 건설부문" },
  { key: 1, co_name: "bsw" },
  { key: 4, co_name: "우원개발" },
];

let currentDate = new Date();
let year = currentDate.getFullYear();
let month = currentDate.getMonth() + 1; // getMonth()는 0부터 시작하므로 1을 더합니다.
let day = currentDate.getDate();

// Excel 파일 경로
const folderPath = "../upload_folder";

// 파일 읽기 및 데이터 처리 함수
function readAndProcessFile(folderPath) {
  fs.readdir(folderPath, (err, files) => {
    if (err) {
      console.log(err, "파일을 찾지 못하였습니다");
      return;
    }

    files.sort((a, b) => {
      return (
        fs.statSync(folderPath + "/" + b).mtime.getTime() -
        fs.statSync(folderPath + "/" + a).mtime.getTime()
      );
    });

    // 최근 파일 가져오기
    const latestFilePath = folderPath + "/" + files[0];

    // Excel 파일 읽기
    const workbook = xlsx.readFile(latestFilePath);

    // 첫 번째 시트 선택
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 시작 row 설정 (0부터 시작)
    const startRow = 3;

    const jsonData = xlsx.utils.sheet_to_json(worksheet, { range: startRow });

    // 데이터 처리 함수 호출
    processJsonData(jsonData);
  });
}

// 데이터 처리 및 파일 쓰기 함수
function processJsonData(jsonData) {
  // Excel의 시리얼 번호를 날짜로 변환하는 함수
  function serialNumberToDate(serialNumber) {
    const millisecondsInOneDay = 24 * 60 * 60 * 1000;
    const excelStartDate = new Date(1899, 11, 30);
    const offsetDays = (serialNumber - 1) * millisecondsInOneDay;
    const date = new Date(excelStartDate.getTime() + offsetDays);
    return date;
  }

  const modifiedJsonData = jsonData.map((item) => {
    let bloodGroup = 0;
    if (item["혈액형\n그룹"] === "RH+") {
      bloodGroup = 1;
    } else if (item["혈액형\n그룹"] === "RH-") {
      bloodGroup = 2;
    }

    let bloodType = 0;
    if (item["혈액형\n타입"] === "A" || item["혈액형\n타입"] === "a") {
      bloodType = 1;
    } else if (item["혈액형\n타입"] === "B" || item["혈액형\n타입"] === "b") {
      bloodType = 2;
    } else if (item["혈액형\n타입"] === "O" || item["혈액형\n타입"] === "o") {
      bloodType = 3;
    } else if (item["혈액형\n타입"] === "AB" || item["혈액형\n타입"] === "ab") {
      bloodType = 4;
    }

    const bcAddress = item["Beacon MAC 번호"].replace(/:/g, "");

    const wkId = parseInt(item["뭘고"], 10); // 10진수로 변환

    const company = companyList.find((co) => co.co_name === item["소속사"]);
    const coId = company ? company.key : null; // 해당하는 기업이 없을 경우 null 반환

    // Excel의 시리얼 번호를 날짜로 변환
    const birthDateSerial = parseFloat(item["생년월일\n(YYYY-MM-DD)"]);
    const birthDate = serialNumberToDate(birthDateSerial);

    const paddedId = String(wkId).padStart(4, "0");
    const bc_index = `BC${paddedId}`;

    // 원하는 형식으로 날짜를 포맷합니다.
    const formattedBirthDate = `${birthDate.getFullYear()}-${(
      birthDate.getMonth() + 1
    )
      .toString()
      .padStart(2, "0")}-${birthDate.getDate().toString().padStart(2, "0")}`;

    const workerIdex = String(wkId).padStart(4, "0");
    const workerIndex = `WK${workerIdex}`;
    return {
      wk_id: wkId,
      wk_index: workerIndex,
      co_id: coId,
      wk_position: item["직위"],
      wk_name: item["이름"],
      wk_birth: formattedBirthDate,
      wk_blood_group: bloodGroup,
      wk_blood_type: bloodType,
      wk_nation: item["국적"],
      wk_phone: item["핸드폰번호\n(010-####-####)"],
      bc_index: bc_index,
      bc_address: bcAddress,
    };
  });

  // JSON 데이터를 파일로 저장
  const jsonFilePath = "./json_folder/file.json";
  fs.writeFileSync(
    jsonFilePath,
    JSON.stringify(modifiedJsonData, null, 2),
    "utf-8"
  );

  console.log("JSON 파일이 성공적으로 생성되었습니다.");

  // JSON 파일 읽기
  const jsonData1 = JSON.parse(
    fs.readFileSync(jsonFilePath, { encoding: "utf-8" })
  );
  createBeaconExcel(modifiedJsonData);
  createWorkerExcel(modifiedJsonData);
  // createCSV(modifiedJsonData);
}

function createBeaconExcel(modifiedJsonData) {
  const beaconData = modifiedJsonData.map((item, index) => {
    return {
      bc_id: item.wk_id,
      bc_index: item.bc_index,
      bc_management: item.wk_id,
      bc_address: item.bc_address,
      bc_used_type: 1,
      ts_index: "SITE0001",
    };
  });

  // Excel 시트 데이터로 변환
  const beaconsheet = xlsx.utils.json_to_sheet(beaconData);

  // Excel 워크북 생성
  const beaconBook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(beaconBook, beaconsheet, "Beacon Data");

  // Excel 파일 저장 경로
  let beaconCount = 1;

  let beaconExcelFilePath = `../transfered_folder/비콘_${year}_${month}_${day}_(${beaconCount}).xlsx`;

  // 파일 경로 생성
  while (fs.existsSync(beaconExcelFilePath)) {
    beaconCount++;
    beaconExcelFilePath = `../transfered_folder/비콘_${year}_${month}_${day}_(${beaconCount}).xlsx`;
  }

  // Excel 파일 쓰기
  xlsx.writeFile(beaconBook, beaconExcelFilePath);
  console.log("비콘 파일 생성완료");
}

function createWorkerExcel(modifiedJsonData) {
  const workerData = modifiedJsonData.map((item, index) => {
    return {
      wk_id: item.wk_id,
      wk_index: item.wk_index,
      co_id: item.co_id,
      wk_position: item.wk_position,
      wk_name: item.wk_name,
      wk_birth: item.wk_birth,
      wk_blood_group: item.wk_blood_group,
      wk_blood_type: item.wk_blood_type,
      wk_nation: item.wk_nation,
      wk_phone: item.wk_phone,
      bc_index: item.bc_index,
    };
  });
  // Excel 시트 데이터로 변환
  const workerSheet = xlsx.utils.json_to_sheet(workerData);

  // Excel 워크북 생성
  const workerBook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workerBook, workerSheet, "Worker Data");

  // Excel 파일 저장 경로
  let workerCount = 1;

  let workerExcelFilePath = `../transfered_folder/작업자_${year}_${month}_${day}_(${workerCount}).xlsx`;

  // 파일 경로 생성
  while (fs.existsSync(workerExcelFilePath)) {
    workerCount++;
    workerExcelFilePath = `../transfered_folder/작업자_${year}_${month}_${day}_(${workerCount}).xlsx`;
  }

  // Excel 파일 쓰기
  xlsx.writeFile(workerBook, workerExcelFilePath);
  console.log("worker 파일 생성완료");
}

function createCSV(modifiedJsonData) {
  // CSV 파일 저장 경로
  const csvFilePath = "../transfered_folder/test.csv";

  // CSV 파일 컬럼 설정
  const csvWriter = createObjectCsvWriter({
    path: csvFilePath,
    header: [
      { id: "wk_id", title: "Worker ID" },
      { id: "wk_index", title: "Worker Index" },
      { id: "co_id", title: "Company ID" },
      { id: "wk_position", title: "Position" },
      { id: "wk_name", title: "Name" },
      { id: "wk_birth", title: "Birth Date" },
      { id: "wk_blood_group", title: "Blood Group" },
      { id: "wk_blood_type", title: "Blood Type" },
      { id: "wk_nation", title: "Nationality" },
      { id: "wk_phone", title: "Phone Number" },
      { id: "bc_index", title: "Beacon Index" },
    ],
    encoding: "utf8",
  });

  // JSON 데이터를 CSV 파일로 쓰기
  csvWriter
    .writeRecords(modifiedJsonData)
    .then(() => console.log("CSV 파일이 성공적으로 생성되었습니다."))
    .catch((err) =>
      console.error("CSV 파일 생성 중 오류가 발생했습니다:", err)
    );
}

// 파일 읽기 및 데이터 처리 함수 호출
readAndProcessFile(folderPath);
