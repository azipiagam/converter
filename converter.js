#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Path tetap untuk memindahkan file yang sudah di-convert
const DONE_INPUT_PATH = "E:\\Azi\\NetBackup\\Done - Input";

// Konstanta untuk batasan file
const MAX_FILE_SIZE_MB = 99;
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;

// Fungsi untuk clean up header names
function cleanHeader(header) {
  if (typeof header !== "string") {
    header = String(header);
  }

  return header
    .toLowerCase() // Convert ke lowercase
    .replace(/[^a-z0-9]+/g, "_") // Ganti semua non-alphanumeric dengan underscore
    .replace(/^_+|_+$/g, "") // Hapus underscore di awal dan akhir
    .replace(/_+/g, "_"); // Ganti multiple underscore dengan single underscore
}

// Fungsi untuk mengkonversi Excel date number ke Date object
function excelDateToJSDate(excelDate) {
  // Excel menggunakan 1900-01-01 sebagai tanggal dasar (serial number 1)
  // Tapi Excel salah menganggap 1900 adalah tahun kabisat
  const excelEpoch = new Date(1899, 11, 30); // 30 Desember 1899
  return new Date(excelEpoch.getTime() + excelDate * 24 * 60 * 60 * 1000);
}

// Fungsi untuk mengecek apakah nilai adalah Excel date
function isExcelDate(value) {
  // Cek apakah value adalah number dan dalam range yang wajar untuk tanggal
  if (typeof value === "number" && value > 1 && value < 2958466) {
    // 1900-01-01 to 9999-12-31
    return true;
  }
  return false;
}

// Mapping bulan Indonesia ke angka
const indonesianMonths = {
  jan: 1,
  januari: 1,
  feb: 2,
  februari: 2,
  mar: 3,
  maret: 3,
  apr: 4,
  april: 4,
  mei: 5,
  may: 5,
  jun: 6,
  juni: 6,
  jul: 7,
  juli: 7,
  agu: 8,
  agustus: 8,
  sep: 9,
  september: 9,
  okt: 10,
  oktober: 10,
  nov: 11,
  november: 11,
  des: 12,
  desember: 12,
  dec: 12,
};

// Fungsi untuk mengecek tipe data temporal
function detectTemporalType(value) {
  if (typeof value !== "string") return null;

  const trimmed = value.trim();

  // Pattern untuk format Indonesia: "01 Agu 2025 23:51"
  const indonesianDateTimePattern =
    /^(\d{1,2})\s+(\w+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/i;

  // Pattern untuk waktu saja: "23:51" atau "23:51:30"
  const timeOnlyPattern = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/;

  // Pattern untuk tanggal standar
  const datePatterns = [
    { pattern: /^\d{4}-\d{1,2}-\d{1,2}$/, type: "date" }, // YYYY-MM-DD
    { pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/, type: "date" }, // MM/DD/YYYY atau DD/MM/YYYY
    { pattern: /^\d{1,2}-\d{1,2}-\d{4}$/, type: "date" }, // MM-DD-YYYY atau DD-MM-YYYY
    { pattern: /^\d{1,2}\.\d{1,2}\.\d{4}$/, type: "date" }, // DD.MM.YYYY
    { pattern: /^\d{4}\/\d{1,2}\/\d{1,2}$/, type: "date" }, // YYYY/MM/DD
  ];

  // Pattern untuk datetime standar
  const datetimePatterns = [
    {
      pattern: /^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    {
      pattern: /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
    {
      pattern: /^\d{1,2}-\d{1,2}-\d{4}\s+\d{1,2}:\d{1,2}(?::\d{1,2})?/,
      type: "datetime",
    },
  ];

  // Cek format Indonesia
  if (indonesianDateTimePattern.test(trimmed)) {
    const match = trimmed.match(indonesianDateTimePattern);
    if (match[4]) {
      // Ada komponen waktu
      return "datetime";
    } else {
      return "date";
    }
  }

  // Cek waktu saja
  if (timeOnlyPattern.test(trimmed)) {
    return "time";
  }

  // Cek datetime patterns
  for (const { pattern, type } of datetimePatterns) {
    if (pattern.test(trimmed)) {
      return type;
    }
  }

  // Cek date patterns
  for (const { pattern, type } of datePatterns) {
    if (pattern.test(trimmed)) {
      return type;
    }
  }

  return null;
}

// Fungsi untuk parse format Indonesia
function parseIndonesianDate(value) {
  const indonesianPattern =
    /^(\d{1,2})\s+(\w+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/i;
  const match = value.match(indonesianPattern);

  if (!match) return null;

  const day = parseInt(match[1]);
  const monthName = match[2].toLowerCase();
  const year = parseInt(match[3]);
  const hour = match[4] ? parseInt(match[4]) : 0;
  const minute = match[5] ? parseInt(match[5]) : 0;
  const second = match[6] ? parseInt(match[6]) : 0;

  const monthNumber = indonesianMonths[monthName];
  if (!monthNumber) return null;

  return new Date(year, monthNumber - 1, day, hour, minute, second);
}

// Fungsi untuk mengkonversi berbagai format tanggal ke format standar
function convertToStandardDate(value) {
  try {
    let dateObj;
    let formatType;

    // Jika value adalah Excel date number
    if (isExcelDate(value)) {
      dateObj = excelDateToJSDate(value);
      formatType = "datetime"; // Default untuk Excel date
    }
    // Jika value adalah string, deteksi tipe temporal
    else if (typeof value === "string") {
      formatType = detectTemporalType(value);

      if (!formatType) {
        return value; // Bukan format temporal, kembalikan nilai asli
      }

      const trimmed = value.trim();

      // Handle format Indonesia
      const indonesianDate = parseIndonesianDate(trimmed);
      if (indonesianDate) {
        dateObj = indonesianDate;
      }
      // Handle time only format
      else if (formatType === "time") {
        const timeMatch = trimmed.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (timeMatch) {
          const hour = parseInt(timeMatch[1]);
          const minute = parseInt(timeMatch[2]);
          const second = timeMatch[3] ? parseInt(timeMatch[3]) : 0;

          // Return formatted time only
          return `${String(hour).padStart(2, "0")}:${String(minute).padStart(
            2,
            "0"
          )}:${String(second).padStart(2, "0")}`;
        }
      }
      // Handle standard date formats
      else {
        dateObj = new Date(trimmed);

        // Jika parsing gagal, coba format Indonesia DD/MM/YYYY
        if (isNaN(dateObj.getTime())) {
          const parts = trimmed.split(/[\/\-\.]/);
          if (parts.length >= 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]);
            const year = parseInt(parts[2]);

            // Jika day > 12, kemungkinan format DD/MM/YYYY
            if (day > 12) {
              dateObj = new Date(year, month - 1, day);
            } else {
              // Coba kedua format
              dateObj = new Date(year, month - 1, day);
              if (isNaN(dateObj.getTime())) {
                dateObj = new Date(year, day - 1, month);
              }
            }
          }
        }
      }
    }
    // Jika value sudah berupa Date object
    else if (value instanceof Date) {
      dateObj = value;
      formatType = "datetime"; // Default
    } else {
      return value; // Bukan tanggal, kembalikan nilai asli
    }

    // Handle time only - sudah di-handle di atas
    if (formatType === "time") {
      return value; // Sudah di-return di atas
    }

    // Cek apakah dateObj valid
    if (!dateObj || isNaN(dateObj.getTime())) {
      return value; // Jika tidak valid, kembalikan nilai asli
    }

    // Format berdasarkan tipe
    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, "0");
    const day = String(dateObj.getDate()).padStart(2, "0");
    const hours = String(dateObj.getHours()).padStart(2, "0");
    const minutes = String(dateObj.getMinutes()).padStart(2, "0");
    const seconds = String(dateObj.getSeconds()).padStart(2, "0");

    if (formatType === "date") {
      // Cek apakah ada komponen waktu yang tidak nol
      if (
        dateObj.getHours() === 0 &&
        dateObj.getMinutes() === 0 &&
        dateObj.getSeconds() === 0
      ) {
        return `${year}-${month}-${day}`;
      } else {
        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
      }
    } else {
      // datetime format
      return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    }
  } catch (error) {
    console.warn(
      `⚠️  Gagal konversi tanggal untuk nilai: ${value}`,
      error.message
    );
    return value; // Kembalikan nilai asli jika gagal
  }
}

// Fungsi untuk memproses dan mengkonversi data
function processRowData(row) {
  const processedRow = {};

  Object.keys(row).forEach((key) => {
    const cleanKey = cleanHeader(key);
    let originalValue = row[key];

    // Handle special case untuk "--" atau dash variants
    if (typeof originalValue === "string") {
      const trimmedValue = originalValue.trim();
      if (
        trimmedValue === "--" ||
        trimmedValue === "—" ||
        trimmedValue === "−"
      ) {
        processedRow[cleanKey] = "";
        console.log(
          `🔄 Konversi dash: "${originalValue}" → "" di kolom "${cleanKey}"`
        );
        return;
      }
    }

    // Konversi tanggal jika terdeteksi
    const convertedValue = convertToStandardDate(originalValue);

    // Log jika ada konversi tanggal (untuk debugging)
    if (convertedValue !== originalValue) {
      const temporalType = detectTemporalType(originalValue);
      if (temporalType) {
        console.log(
          `📅 Konversi ${temporalType}: "${originalValue}" → "${convertedValue}" di kolom "${cleanKey}"`
        );
      }
    }

    processedRow[cleanKey] = convertedValue;
  });

  return processedRow;
}

// Fungsi untuk mendapatkan ukuran file dalam bytes
function getFileSizeInBytes(filePath) {
  try {
    const stats = fs.statSync(filePath);
    return stats.size;
  } catch (error) {
    return 0;
  }
}

// Fungsi untuk format ukuran file
function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// Fungsi untuk memindahkan file ke Done - Input
function moveFileToProcessed(filePath) {
  try {
    // Pastikan folder Done - Input ada
    if (!fs.existsSync(DONE_INPUT_PATH)) {
      fs.mkdirSync(DONE_INPUT_PATH, { recursive: true });
      console.log(`📁 Folder Done - Input dibuat: ${DONE_INPUT_PATH}`);
    }

    const fileName = path.basename(filePath);
    const destinationPath = path.join(DONE_INPUT_PATH, fileName);

    // Cek jika file sudah ada di destination
    if (fs.existsSync(destinationPath)) {
      const baseName = path.basename(fileName, path.extname(fileName));
      const ext = path.extname(fileName);
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      const newFileName = `${baseName}_${timestamp}${ext}`;
      const newDestinationPath = path.join(DONE_INPUT_PATH, newFileName);

      fs.renameSync(filePath, newDestinationPath);
      console.log(
        `📦 File dipindahkan ke: ${newDestinationPath} (renamed karena duplikat)`
      );
      return newDestinationPath;
    } else {
      fs.renameSync(filePath, destinationPath);
      console.log(`📦 File dipindahkan ke: ${destinationPath}`);
      return destinationPath;
    }
  } catch (error) {
    console.error(`❌ Gagal memindahkan file ${filePath}:`, error.message);
    return null;
  }
}

// Fungsi untuk convert single file
function convertXlsToJsonl(
  inputPath,
  outputPath = null,
  moveAfterConvert = true
) {
  try {
    // Normalize input path
    inputPath = path.resolve(inputPath);

    // Baca file Excel
    console.log(`📖 Membaca file: ${inputPath}`);
    const workbook = XLSX.readFile(inputPath);

    // Ambil sheet pertama (atau bisa dimodif untuk pilih sheet)
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert ke JSON dengan opsi untuk mempertahankan tipe data
    const rawJsonData = XLSX.utils.sheet_to_json(worksheet, {
      raw: false, // Jangan convert semua ke string
      dateNF: "yyyy-mm-dd", // Format tanggal default
      cellDates: true, // Parse tanggal sebagai Date object
    });

    console.log(`🔄 Memproses ${rawJsonData.length} baris data...`);

    // Proses setiap row untuk konversi tanggal dan clean headers
    const jsonData = rawJsonData.map((row, index) => {
      try {
        return processRowData(row);
      } catch (error) {
        console.warn(`⚠️  Error memproses baris ${index + 1}:`, error.message);
        // Fallback: clean headers saja tanpa konversi tanggal
        const cleanedRow = {};
        Object.keys(row).forEach((key) => {
          const cleanKey = cleanHeader(key);
          cleanedRow[cleanKey] = row[key];
        });
        return cleanedRow;
      }
    });

    // Convert array ke JSON Lines format (satu object per line)
    const jsonlData = jsonData.map((row) => JSON.stringify(row)).join("\n");

    // Tentukan output path
    if (!outputPath) {
      const inputName = path.basename(inputPath, path.extname(inputPath));

      // Cek apakah ada folder output di working directory
      const outputDir = path.join(process.cwd(), "output");
      if (fs.existsSync(outputDir) && fs.statSync(outputDir).isDirectory()) {
        outputPath = path.join(outputDir, `${inputName}.jsonl`);
      } else {
        // Fallback ke folder yang sama dengan input
        const inputDir = path.dirname(inputPath);
        outputPath = path.join(inputDir, `${inputName}.jsonl`);
      }
    } else {
      // Handle jika outputPath adalah folder, bukan file
      outputPath = path.resolve(outputPath);

      if (fs.existsSync(outputPath) && fs.statSync(outputPath).isDirectory()) {
        // Jika outputPath adalah folder, buat nama file
        const inputName = path.basename(inputPath, path.extname(inputPath));
        outputPath = path.join(outputPath, `${inputName}.jsonl`);
      } else if (
        !outputPath.endsWith(".jsonl") &&
        !outputPath.endsWith(".json")
      ) {
        // Jika tidak ada ekstensi dan folder belum ada, anggap sebagai folder
        const inputName = path.basename(inputPath, path.extname(inputPath));
        // Buat folder jika belum ada
        fs.mkdirSync(outputPath, { recursive: true });
        outputPath = path.join(outputPath, `${inputName}.jsonl`);
      }
    }

    // Pastikan output directory ada
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    // Tulis file JSON Lines
    fs.writeFileSync(outputPath, jsonlData, "utf8");
    console.log(`✅ Berhasil convert: ${outputPath}`);
    console.log(`📊 Total records: ${jsonData.length}`);

    // Pindahkan file Excel ke Done - Input jika conversion berhasil
    if (moveAfterConvert) {
      const movedPath = moveFileToProcessed(inputPath);
      if (movedPath) {
        return { success: true, outputPath, movedPath, jsonData };
      } else {
        return { success: true, outputPath, movedPath: null, jsonData };
      }
    }

    return { success: true, outputPath, jsonData };
  } catch (error) {
    console.error(`❌ Error converting ${inputPath}:`, error.message);
    return { success: false, error: error.message };
  }
}

// Fungsi untuk merge JSONL files dengan batasan ukuran
function mergeJsonlFiles(jsonlFiles, outputDir, baseName = "merged") {
  try {
    console.log(`\n🔄 Memulai proses merge ${jsonlFiles.length} file JSONL...`);
    console.log(`📏 Batasan ukuran per file: ${MAX_FILE_SIZE_MB}MB`);

    let currentBatch = [];
    let currentSize = 0;
    let batchNumber = 1;
    const mergedFiles = [];
    let totalRecords = 0;

    for (const filePath of jsonlFiles) {
      console.log(`📖 Membaca: ${path.basename(filePath)}`);

      const fileContent = fs.readFileSync(filePath, "utf8");
      const fileSize = Buffer.byteLength(fileContent, "utf8");

      console.log(`   Ukuran: ${formatFileSize(fileSize)}`);

      // Jika file ini akan membuat batch melebihi batas, simpan batch saat ini
      if (
        currentSize + fileSize > MAX_FILE_SIZE_BYTES &&
        currentBatch.length > 0
      ) {
        const mergedFilePath = saveBatch(
          currentBatch,
          outputDir,
          baseName,
          batchNumber,
          currentSize
        );
        mergedFiles.push({
          path: mergedFilePath,
          size: currentSize,
          fileCount: currentBatch.length,
        });

        // Reset untuk batch baru
        currentBatch = [];
        currentSize = 0;
        batchNumber++;
      }

      // Tambahkan file ke batch saat ini
      currentBatch.push(fileContent);
      currentSize += fileSize;

      // Hitung jumlah records
      const recordCount = fileContent
        .split("\n")
        .filter((line) => line.trim()).length;
      totalRecords += recordCount;

      console.log(`   Records: ${recordCount}`);
    }

    // Simpan batch terakhir jika ada
    if (currentBatch.length > 0) {
      const mergedFilePath = saveBatch(
        currentBatch,
        outputDir,
        baseName,
        batchNumber,
        currentSize
      );
      mergedFiles.push({
        path: mergedFilePath,
        size: currentSize,
        fileCount: currentBatch.length,
      });
    }

    // Hapus file JSONL individual setelah merge berhasil
    console.log(`\n🗑️  Menghapus file JSONL individual...`);
    let deletedCount = 0;
    for (const filePath of jsonlFiles) {
      try {
        fs.unlinkSync(filePath);
        deletedCount++;
        console.log(`   ✅ Dihapus: ${path.basename(filePath)}`);
      } catch (error) {
        console.warn(
          `   ⚠️  Gagal hapus: ${path.basename(filePath)} - ${error.message}`
        );
      }
    }

    console.log(`\n🎉 Merge selesai!`);
    console.log(`📊 Summary:`);
    console.log(`   Total file merged: ${mergedFiles.length}`);
    console.log(`   Total records: ${totalRecords.toLocaleString()}`);
    console.log(
      `   File individual dihapus: ${deletedCount}/${jsonlFiles.length}`
    );

    console.log(`\n📁 File hasil merge:`);
    mergedFiles.forEach((file, index) => {
      console.log(`   ${index + 1}. ${path.basename(file.path)}`);
      console.log(`      Ukuran: ${formatFileSize(file.size)}`);
      console.log(`      File count: ${file.fileCount}`);
    });

    return {
      success: true,
      mergedFiles: mergedFiles.map((f) => f.path),
      totalRecords,
      deletedCount,
    };
  } catch (error) {
    console.error(`❌ Error merging files:`, error.message);
    return { success: false, error: error.message };
  }
}

// Fungsi helper untuk menyimpan batch
function saveBatch(
  batchContent,
  outputDir,
  baseName,
  batchNumber,
  currentSize
) {
  const paddedNumber = String(batchNumber).padStart(3, "0");
  const mergedFileName = `${baseName}_${paddedNumber}.jsonl`;
  const mergedFilePath = path.join(outputDir, mergedFileName);

  const mergedContent = batchContent.join("\n");
  fs.writeFileSync(mergedFilePath, mergedContent, "utf8");

  const recordCount = mergedContent
    .split("\n")
    .filter((line) => line.trim()).length;

  console.log(`💾 Batch ${batchNumber} disimpan: ${mergedFileName}`);
  console.log(`   Ukuran: ${formatFileSize(currentSize)}`);
  console.log(`   Records: ${recordCount.toLocaleString()}`);
  console.log(`   File count: ${batchContent.length}`);

  return mergedFilePath;
}

// Fungsi untuk convert multiple files di folder dengan merge
function convertFolderWithMerge(
  folderPath,
  customOutputDir = null,
  moveAfterConvert = true
) {
  try {
    const files = fs.readdirSync(folderPath);
    const xlsFiles = files.filter(
      (file) =>
        path.extname(file).toLowerCase() === ".xls" ||
        path.extname(file).toLowerCase() === ".xlsx"
    );

    if (xlsFiles.length === 0) {
      console.log("❌ Tidak ada file .xls/.xlsx ditemukan di folder ini");
      return;
    }

    // Sort files numerically if they follow pattern like 1.xlsx, 2.xlsx, etc.
    xlsFiles.sort((a, b) => {
      const aNum = parseInt(path.basename(a, path.extname(a)));
      const bNum = parseInt(path.basename(b, path.extname(b)));
      if (!isNaN(aNum) && !isNaN(bNum)) {
        return aNum - bNum;
      }
      return a.localeCompare(b);
    });

    console.log(`📁 Ditemukan ${xlsFiles.length} file Excel di: ${folderPath}`);

    // Tentukan output directory
    let outputDir;
    if (customOutputDir) {
      outputDir = path.resolve(customOutputDir);
    } else {
      // Default ke folder "Output" di dalam input folder
      outputDir = path.join(folderPath, "Output");
    }

    // Cek atau buat folder output
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    console.log(`📤 Output akan disimpan ke: ${outputDir}`);
    if (moveAfterConvert) {
      console.log(
        `📦 File Excel yang berhasil akan dipindahkan ke: ${DONE_INPUT_PATH}`
      );
    }

    // Buat folder temp untuk JSONL individual
    const tempDir = path.join(outputDir, ".temp");
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    let successCount = 0;
    let movedCount = 0;
    const results = [];
    const jsonlFiles = [];

    console.log(`\n🔄 === TAHAP 1: CONVERT EXCEL KE JSONL ===`);

    // Convert semua file Excel ke JSONL individual
    xlsFiles.forEach((file, index) => {
      const inputPath = path.join(folderPath, file);
      const fileName = path.basename(file, path.extname(file));
      const tempOutputPath = path.join(tempDir, `${fileName}.jsonl`);

      console.log(`\n📄 Processing ${index + 1}/${xlsFiles.length}: ${file}`);
      const result = convertXlsToJsonl(
        inputPath,
        tempOutputPath,
        moveAfterConvert
      );
      results.push({ file, result });

      if (result.success) {
        successCount++;
        jsonlFiles.push(tempOutputPath);
        if (result.movedPath) {
          movedCount++;
        }
      }
    });

    console.log(
      `\n✅ Tahap 1 selesai: ${successCount}/${xlsFiles.length} file berhasil diconvert`
    );

    if (jsonlFiles.length === 0) {
      console.log(
        "❌ Tidak ada file JSONL yang berhasil dibuat untuk di-merge"
      );
      // Hapus temp folder
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch (error) {
        console.warn(`⚠️  Gagal hapus temp folder: ${error.message}`);
      }
      return;
    }

    // Merge JSONL files
    console.log(`\n🔄 === TAHAP 2: MERGE JSONL FILES ===`);
    const folderBaseName = path.basename(folderPath);
    const mergeResult = mergeJsonlFiles(jsonlFiles, outputDir, folderBaseName);

    // Hapus temp folder
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
      console.log(`🗑️  Temp folder dihapus`);
    } catch (error) {
      console.warn(`⚠️  Gagal hapus temp folder: ${error.message}`);
    }

    if (mergeResult.success) {
      console.log(`\n🎉 PROSES SELESAI!`);
      console.log(`📊 Summary total:`);
      console.log(
        `   Excel files processed: ${successCount}/${xlsFiles.length}`
      );
      console.log(`   Merged JSONL files: ${mergeResult.mergedFiles.length}`);
      console.log(
        `   Total records: ${mergeResult.totalRecords.toLocaleString()}`
      );

      if (moveAfterConvert) {
        console.log(`   Excel files moved: ${movedCount}/${successCount}`);
      }

      // Show failed moves if any
      const failedMoves = results.filter(
        (r) => r.result.success && r.result.movedPath === null
      );
      if (failedMoves.length > 0) {
        console.log(
          `⚠️  ${failedMoves.length} file Excel berhasil diconvert tapi gagal dipindahkan:`
        );
        failedMoves.forEach((fm) => console.log(`     - ${fm.file}`));
      }
    } else {
      console.error(`❌ Merge gagal: ${mergeResult.error}`);
    }
  } catch (error) {
    console.error("❌ Error processing folder:", error.message);
  }
}

// Fungsi untuk convert multiple files di folder tanpa merge (legacy)
function convertFolder(
  folderPath,
  customOutputDir = null,
  moveAfterConvert = true
) {
  try {
    const files = fs.readdirSync(folderPath);
    const xlsFiles = files.filter(
      (file) =>
        path.extname(file).toLowerCase() === ".xls" ||
        path.extname(file).toLowerCase() === ".xlsx"
    );

    if (xlsFiles.length === 0) {
      console.log("❌ Tidak ada file .xls/.xlsx ditemukan di folder ini");
      return;
    }

    console.log(`📁 Ditemukan ${xlsFiles.length} file Excel di: ${folderPath}`);

    // Tentukan output directory
    let outputDir;
    if (customOutputDir) {
      outputDir = path.resolve(customOutputDir);
    } else {
      // Default ke folder output di project
      outputDir = path.join(process.cwd(), "output");
    }

    // Cek atau buat folder output
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`📁 Folder output dibuat: ${outputDir}`);
    }

    console.log(`📤 Output akan disimpan ke: ${outputDir}`);
    if (moveAfterConvert) {
      console.log(
        `📦 File yang berhasil akan dipindahkan ke: ${DONE_INPUT_PATH}`
      );
    }

    let successCount = 0;
    let movedCount = 0;
    const results = [];

    xlsFiles.forEach((file) => {
      const inputPath = path.join(folderPath, file);
      const fileName = path.basename(file, path.extname(file));
      const outputPath = path.join(outputDir, `${fileName}.jsonl`);

      console.log(`\n🔄 Processing: ${file}`);
      const result = convertXlsToJsonl(inputPath, outputPath, moveAfterConvert);
      results.push({ file, result });

      if (result.success) {
        successCount++;
        if (result.movedPath) {
          movedCount++;
        }
      }
    });

    console.log(
      `\n🎉 Selesai! ${successCount}/${xlsFiles.length} file berhasil diconvert`
    );
    if (moveAfterConvert) {
      console.log(
        `📦 ${movedCount}/${successCount} file berhasil dipindahkan ke Done - Input`
      );
    }

    // Show summary of any failed moves
    const failedMoves = results.filter(
      (r) => r.result.success && r.result.movedPath === null
    );
    if (failedMoves.length > 0) {
      console.log(
        `⚠️  ${failedMoves.length} file berhasil diconvert tapi gagal dipindahkan:`
      );
      failedMoves.forEach((fm) => console.log(`   - ${fm.file}`));
    }
  } catch (error) {
    console.error("❌ Error membaca folder:", error.message);
  }
}

// Main function
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log(`
🔄 XLS to JSON Lines Converter with Smart Merger (Max ${MAX_FILE_SIZE_MB}MB per file)

Cara pakai:
  node converter.js <input-folder>                       # Convert & merge semua file di folder
  node converter.js <input-folder> <output-folder>       # Convert & merge dengan custom output folder
  node converter.js <file.xls>                           # Convert single file (no merge)
  node converter.js <file.xls> <output.jsonl>            # Convert single file dengan custom output
  node converter.js --no-merge <input-folder>            # Convert tanpa merge (individual JSONL)
  node converter.js --no-move <input-folder>             # Convert & merge tanpa memindahkan Excel files
  node converter.js --help                               # Show help

Contoh:
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli"
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli" "C:\\Users\\PIAGAM\\Downloads\\Juli\\jsonl"
  node converter.js data.xlsx
  node converter.js --no-merge "C:\\Data\\Excel\\"
  node converter.js --no-move "C:\\Data\\Excel\\"

Features:
  📦 MERGE: Menggabungkan multiple JSONL dengan batasan ${MAX_FILE_SIZE_MB}MB per file
  📅 AUTO-CONVERT: Tanggal ke format yyyy-mm-dd hh:mm:ss
  📁 DYNAMIC PATH: Input dan output folder yang fleksibel
  🏷️  AUTO-NAMING: File merged menggunakan nama folder + nomor urut
  📂 DEFAULT OUTPUT: Jika tidak ada output path, buat folder "Output" di input folder
  📦 AUTO-MOVE: File Excel yang berhasil dipindahkan ke: ${DONE_INPUT_PATH}
        `);
    return;
  }

  if (args[0] === "--help" || args[0] === "-h") {
    console.log(`
🔄 XLS to JSON Lines Converter with Smart Merger - Help

MODES:
  1. FOLDER MODE (with merge): node converter.js <input-folder> [output-folder]
     - Convert semua Excel files di folder
     - Merge hasil JSONL dengan batasan ${MAX_FILE_SIZE_MB}MB per file
     - Default output: <input-folder>/Output/

  2. SINGLE FILE MODE: node converter.js <file.xlsx> [output.jsonl]
     - Convert single file tanpa merge
     - Standard behavior seperti versi lama

  3. NO-MERGE MODE: node converter.js --no-merge <input-folder> [output-folder]
     - Convert folder tapi tanpa merge (individual JSONL files)
     - Berguna jika ingin JSONL terpisah

  4. NO-MOVE MODE: node converter.js --no-move <input-folder> [output-folder]
     - Convert & merge tapi tidak pindahkan Excel files ke Done-Input

FLAGS:
  --no-merge    : Convert folder tanpa merge (individual JSONL files)
  --no-move     : Tidak memindahkan Excel files setelah conversion
  --help, -h    : Show this help

EXAMPLES:
  # Merge mode dengan dynamic paths
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli"
  node converter.js "C:\\Users\\PIAGAM\\Downloads\\Juli" "D:\\Results\\Merged\\"
  
  # Single file mode
  node converter.js "C:\\Data\\report.xlsx"
  node converter.js "C:\\Data\\report.xlsx" "D:\\Results\\report.jsonl"
  
  # No merge mode (individual files)
  node converter.js --no-merge "C:\\Excel\\Data\\" "D:\\Individual\\"
  
  # No move mode (keep Excel files in place)
  node converter.js --no-move "C:\\Excel\\Data\\" "D:\\Results\\"

MERGE BEHAVIOR:
  📏 Max size per merged file: ${MAX_FILE_SIZE_MB}MB
  🔢 File naming: {folder-name}_001.jsonl, {folder-name}_002.jsonl, dst.
  📊 Otomatis hitung total records dan file statistics
  🗑️  Individual JSONL files dihapus otomatis setelah merge berhasil
  📋 Sort files numerically (1.xlsx, 2.xlsx, 10.xlsx, dst.)

OUTPUT PATH LOGIC:
  - Jika tidak ada output path: buat folder "Output" di input folder
  - Jika output path adalah file: gunakan sebagai file output
  - Jika output path adalah folder: simpan di folder tersebut
  - Otomatis create directories yang belum ada

SUPPORTED FORMATS:
  📁 Input: .xls, .xlsx files
  📄 Output: .jsonl (JSON Lines format)
  📅 Date formats: Excel dates, Indonesian format, ISO, dll.
  🔤 Headers: Auto-clean (lowercase, underscore)

FEATURES:
  ✅ Smart date/time detection dan conversion
  ✅ Clean header names (lowercase + underscore)
  ✅ Batch processing dengan progress indicators
  ✅ Error handling dan recovery
  ✅ File size monitoring dan batching
  ✅ Auto-move processed files ke Done-Input
  ✅ Duplicate filename handling dengan timestamp
  ✅ Memory efficient processing
  ✅ Cross-platform path handling
        `);
    return;
  }

  // Check for flags
  let noMerge = false;
  let moveAfterConvert = true;
  let actualArgs = [...args];

  if (args[0] === "--no-merge") {
    noMerge = true;
    actualArgs = args.slice(1);
    console.log("🚫 Mode: Tidak akan merge file JSONL (individual files)");
  } else if (args[0] === "--no-move") {
    moveAfterConvert = false;
    actualArgs = args.slice(1);
    console.log(
      "🚫 Mode: Excel files tidak akan dipindahkan setelah conversion"
    );
  }

  if (actualArgs.length === 0) {
    console.error("❌ Setelah flag, harus ada input file/folder");
    return;
  }

  const inputPath = actualArgs[0];
  const outputPath = actualArgs[1];

  // Cek apakah path ada
  if (!fs.existsSync(inputPath)) {
    console.error("❌ File atau folder tidak ditemukan:", inputPath);
    return;
  }

  // Normalize paths
  const resolvedInputPath = path.resolve(inputPath);
  const stats = fs.statSync(resolvedInputPath);

  if (stats.isDirectory()) {
    // Convert semua file di folder
    let outputFolder = outputPath;

    // Jika tidak ada output path, buat folder "Output" di dalam input folder
    if (!outputFolder) {
      outputFolder = path.join(resolvedInputPath, "Output");
      console.log(
        `📁 Output path tidak disebutkan, akan menggunakan: ${outputFolder}`
      );
    }

    if (noMerge) {
      // Mode tanpa merge (legacy behavior)
      convertFolder(resolvedInputPath, outputFolder, moveAfterConvert);
    } else {
      // Mode dengan merge (new behavior)
      convertFolderWithMerge(resolvedInputPath, outputFolder, moveAfterConvert);
    }
  } else if (stats.isFile()) {
    // Convert single file (tidak ada merge untuk single file)
    const ext = path.extname(resolvedInputPath).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") {
      const result = convertXlsToJsonl(
        resolvedInputPath,
        outputPath,
        moveAfterConvert
      );
      if (result.success && moveAfterConvert && !result.movedPath) {
        console.log("⚠️  File berhasil diconvert tapi gagal dipindahkan");
      }
    } else {
      console.error("❌ File harus berformat .xls atau .xlsx");
    }
  }
}

// Jalankan program
if (require.main === module) {
  main();
}
