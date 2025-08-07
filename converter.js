#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Fungsi untuk clean up header names
function cleanHeader(header) {
  if (typeof header !== "string") {
    return String(header).toLowerCase().replace(/\s+/g, "_");
  }
  return header.toLowerCase().replace(/\s+/g, "_");
}

// Fungsi untuk convert single file
function convertXlsToJsonl(inputPath, outputPath = null) {
  try {
    // Baca file Excel
    console.log(`📖 Membaca file: ${inputPath}`);
    const workbook = XLSX.readFile(inputPath);

    // Ambil sheet pertama (atau bisa dimodif untuk pilih sheet)
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert ke JSON Lines (JSONL) dengan header yang sudah di-clean
    const rawJsonData = XLSX.utils.sheet_to_json(worksheet);

    // Clean up headers dan rebuild objects
    const jsonData = rawJsonData.map((row) => {
      const cleanedRow = {};
      Object.keys(row).forEach((key) => {
        const cleanKey = cleanHeader(key);
        cleanedRow[cleanKey] = row[key];
      });
      return cleanedRow;
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
    }

    // Tulis file JSON Lines
    fs.writeFileSync(outputPath, jsonlData, "utf8");
    console.log(`✅ Berhasil convert: ${outputPath}`);
    console.log(`📊 Total records: ${jsonData.length}`);

    return outputPath;
  } catch (error) {
    console.error(`❌ Error converting ${inputPath}:`, error.message);
    return null;
  }
}

// Fungsi untuk convert multiple files di folder
function convertFolder(folderPath, customOutputDir = null) {
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

    let successCount = 0;
    xlsFiles.forEach((file) => {
      const inputPath = path.join(folderPath, file);
      const fileName = path.basename(file, path.extname(file));
      const outputPath = path.join(outputDir, `${fileName}.jsonl`);

      const result = convertXlsToJsonl(inputPath, outputPath);
      if (result) successCount++;
    });

    console.log(
      `\n🎉 Selesai! ${successCount}/${xlsFiles.length} file berhasil diconvert`
    );
  } catch (error) {
    console.error("❌ Error membaca folder:", error.message);
  }
}

// Main function
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log(`
🔄 XLS to JSON Lines Converter

Cara pakai:
  node converter.js <file.xls>                           # Convert single file
  node converter.js <file.xls> <output.jsonl>           # Convert dengan custom output
  node converter.js <input-folder>                       # Convert semua file di folder
  node converter.js <input-folder> <output-folder>       # Convert dengan custom output folder
  node converter.js --help                               # Show help

Contoh:
  node converter.js data.xls
  node converter.js data.xls hasil.jsonl
  node converter.js ./excel-files/
  node converter.js C:\\Data\\Excel\\ D:\\Results\\
  node converter.js ../company-data/ ~/Desktop/output/
        `);
    return;
  }

  if (args[0] === "--help" || args[0] === "-h") {
    console.log(`
🔄 XLS to JSON Lines Converter - Help

Commands:
  Single file: node converter.js input.xls [output.jsonl]
  Folder:      node converter.js <input-folder> [output-folder]
  Help:        node converter.js --help

Examples:
  node converter.js data.xls
  node converter.js C:\\Excel\\data.xls D:\\Output\\result.jsonl
  node converter.js ./input-folder/ ./output-folder/
  node converter.js C:\\Company\\Reports\\ D:\\Processed\\
  node converter.js ../external-data/ ~/Desktop/results/

Features:
  ✅ Support .xls dan .xlsx
  ✅ Convert to JSON Lines format (JSONL)
  ✅ Clean headers (lowercase + underscore)
  ✅ Batch convert untuk folder
  ✅ Custom input/output paths (dalam atau luar project)
  ✅ Auto-create output directories
  ✅ Error handling yang baik
  ✅ Shows record count
        `);
    return;
  }

  const inputPath = args[0];
  const outputPath = args[1];

  // Cek apakah path ada
  if (!fs.existsSync(inputPath)) {
    console.error("❌ File atau folder tidak ditemukan:", inputPath);
    return;
  }

  // Cek apakah itu folder atau file
  const stats = fs.statSync(inputPath);

  if (stats.isDirectory()) {
    // Convert semua file di folder
    const outputFolder = args[1]; // Optional output folder
    convertFolder(inputPath, outputFolder);
  } else if (stats.isFile()) {
    // Convert single file
    const ext = path.extname(inputPath).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") {
      convertXlsToJsonl(inputPath, outputPath);
    } else {
      console.error("❌ File harus berformat .xls atau .xlsx");
    }
  }
}

// Jalankan program
if (require.main === module) {
  main();
}
