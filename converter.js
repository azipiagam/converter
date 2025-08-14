#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Path tetap untuk memindahkan file yang sudah di-convert
const DONE_INPUT_PATH = "E:\\Azi\\NetBackup\\Done - Input";

// Fungsi untuk clean up header names
function cleanHeader(header) {
  if (typeof header !== "string") {
    return String(header).toLowerCase().replace(/\s+/g, "_");
  }
  return header.toLowerCase().replace(/\s+/g, "_");
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
        return { success: true, outputPath, movedPath };
      } else {
        return { success: true, outputPath, movedPath: null };
      }
    }

    return { success: true, outputPath };
  } catch (error) {
    console.error(`❌ Error converting ${inputPath}:`, error.message);
    return { success: false, error: error.message };
  }
}

// Fungsi untuk convert multiple files di folder
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
🔄 XLS to JSON Lines Converter (with Auto-Move)

Cara pakai:
  node converter.js <file.xls>                           # Convert single file + move
  node converter.js <file.xls> <output.jsonl>           # Convert dengan custom output + move
  node converter.js <input-folder>                       # Convert semua file di folder + move
  node converter.js <input-folder> <output-folder>       # Convert dengan custom output folder + move
  node converter.js --no-move <file/folder>             # Convert tanpa memindahkan file
  node converter.js --help                               # Show help

Contoh:
  node converter.js data.xls
  node converter.js data.xls hasil.jsonl
  node converter.js ./excel-files/
  node converter.js C:\\Data\\Excel\\ D:\\Results\\
  node converter.js --no-move data.xls                   # Tidak pindahkan file

📦 File yang berhasil diconvert akan dipindahkan ke: ${DONE_INPUT_PATH}
        `);
    return;
  }

  if (args[0] === "--help" || args[0] === "-h") {
    console.log(`
🔄 XLS to JSON Lines Converter (with Auto-Move) - Help

Commands:
  Single file: node converter.js input.xls [output.jsonl]
  Folder:      node converter.js <input-folder> [output-folder]
  No Move:     node converter.js --no-move <input> [output]
  Help:        node converter.js --help

Examples:
  node converter.js data.xls
  node converter.js C:\\Excel\\data.xls D:\\Output\\result.jsonl
  node converter.js ./input-folder/ ./output-folder/
  node converter.js C:\\Company\\Reports\\ D:\\Processed\\
  node converter.js ../external-data/ ~/Desktop/results/
  node converter.js --no-move data.xls                   # Tanpa pindah file

Features:
  ✅ Support .xls dan .xlsx
  ✅ Convert to JSON Lines format (JSONL)
  ✅ Clean headers (lowercase + underscore)
  ✅ Batch convert untuk folder
  ✅ Custom input/output paths (dalam atau luar project)
  ✅ Auto-create output directories
  ✅ Auto-move processed files ke: ${DONE_INPUT_PATH}
  ✅ Handle duplicate files dengan timestamp
  ✅ Error handling yang baik
  ✅ Shows record count
  ✅ Option --no-move untuk disable auto-move
        `);
    return;
  }

  // Check for --no-move flag
  let moveAfterConvert = true;
  let actualArgs = [...args];

  if (args[0] === "--no-move") {
    moveAfterConvert = false;
    actualArgs = args.slice(1);
    console.log("🚫 Mode: File tidak akan dipindahkan setelah conversion");
  }

  if (actualArgs.length === 0) {
    console.error("❌ Setelah --no-move, harus ada input file/folder");
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
    const outputFolder = actualArgs[1]; // Optional output folder
    convertFolder(resolvedInputPath, outputFolder, moveAfterConvert);
  } else if (stats.isFile()) {
    // Convert single file
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
