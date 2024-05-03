import fs from "fs";
import fsp from "fs/promises";
import rl from "readline";
import path from "path";
import url from "url";
const __filename = url.fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

import express from "express";
import fileUpload from "express-fileupload";
import unzipper from "unzipper";
import exceljs from "exceljs";

const app = express();

app.use(fileUpload());

app.get("/", (req, res) => {
    res.sendFile(__dirname + "/index.html");
});

function unzip(zipFilePath, destination) {
    const stream = fs
        .createReadStream(zipFilePath)
        .pipe(unzipper.Extract({ path: destination }));

    return new Promise((resolve, reject) => {
        stream.on("finish", () => resolve());
        stream.on("error", (error) => reject(error));
    });
}

app.post("/upload", async (req, res) => {
    if (!req.files || !req.files.file) {
        return res.status(400).send("No files were uploaded.");
    }

    const uploadedZipFile = req.files.file;
    const excludeFiles = req.body.excludeFiles?.split(",").map((s) => s.trim());

    const tempFolderPath = __dirname + "/temp_folder/";
    const zipFilePath = tempFolderPath + "upload.zip";

    try {
        await fsp.mkdir(tempFolderPath);
        await uploadedZipFile.mv(zipFilePath);
        await unzip(zipFilePath, tempFolderPath);

        const files = await fsp.readdir(tempFolderPath, { recursive: true });

        // const wb = {
        //     cols: [
        //         {header: "folder1", key: "folder1"},
        //         {header: "folder2", key: "folder2"},
        //     ],
        //     sheets: [
        //         {name: "file1", keys: {folder1: [1,2,3], folder2: [4,5,6]}},
        //         {name: "file2", keys: {folder1: [1,2,3], folder2: [4,5,6]}},
        //     ]
        // };

        const wb = {
            cols: [],
            sheets: [],
        };

        for (const file of files) {
            if (file.endsWith(".json")) {
                const folderName = file.split("\\")[0];
                const fileName = file.split("\\")[1].replace(".json", "");

                if (excludeFiles.includes(fileName)) {
                    continue;
                }

                const existFolderNameInd = wb.cols.findIndex(
                    (col) => col.key === folderName
                );
                if (existFolderNameInd < 0) {
                    wb.cols.push({ header: folderName, key: folderName });
                }

                const imageKeys = [];
                const fileStream = fs.createReadStream(tempFolderPath + file);
                const lines = rl.createInterface({
                    input: fileStream,
                    crlfDelay: Infinity,
                });
                for await (const line of lines) {
                    const found = line.match(
                        /"http(s?):\/\/img-prod-cms(.*)"/g
                    );
                    if (found) {
                        const url = found[0].replaceAll(`"`, "");
                        const imageKeyWithQueryParams = url
                            .split("/")
                            .reverse()[0];
                        const imageKey = imageKeyWithQueryParams.split("?")[0];
                        imageKeys.push(imageKey);
                    }
                }
                const uniqueImageKeys = Array.from(new Set(imageKeys));

                const sheetName = fileName;
                const existFileNameInd = wb.sheets.findIndex(
                    (sheet) => sheet.name === sheetName
                );
                if (existFileNameInd < 0) {
                    wb.sheets.push({
                        name: sheetName,
                        keys: { [folderName]: uniqueImageKeys },
                    });
                } else {
                    wb.sheets[existFileNameInd].keys[folderName] =
                        uniqueImageKeys;
                }
            }
        }

        const workbook = new exceljs.Workbook();

        for (const sheet of wb.sheets) {
            const worksheet = workbook.addWorksheet(sheet.name, {
                properties: { defaultColWidth: 15 },
            });

            const sortedCols = [...wb.cols].sort((x, y) =>
                x.header.localeCompare(y.header)
            );

            let colsIndex = 1;
            sortedCols.forEach((col) => {
                const headerCell = worksheet.getRow(1).getCell(colsIndex);
                headerCell.value = col.header;
                headerCell.alignment = {
                    vertical: "middle",
                    horizontal: "center",
                };
                headerCell.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };
                worksheet.mergeCells(1, colsIndex, 1, colsIndex + 1);

                const subHeaderCell1 = worksheet.getRow(2).getCell(colsIndex);
                subHeaderCell1.value = "RT";
                subHeaderCell1.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };

                const subHeaderCell2 = worksheet
                    .getRow(2)
                    .getCell(colsIndex + 1);
                subHeaderCell2.value = "AEM";
                subHeaderCell2.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };

                if (Object.keys(sheet.keys).includes(col.header)) {
                    const imageKeys = sheet.keys[col.header];
                    for (let i = 0; i < imageKeys.length; i++) {
                        worksheet.getRow(3 + i).getCell(colsIndex).value =
                            imageKeys[i];
                    }
                }

                colsIndex += 2;
            });
        }

        const downloadExcelFileName = tempFolderPath + "output.xlsx";
        await workbook.xlsx.writeFile(downloadExcelFileName);
        res.download(downloadExcelFileName);
    } catch (err) {
        console.error(err);
        res.status(500).send("Error processing file");
    } finally {
        await fsp.rm(tempFolderPath, { recursive: true });
    }
});

app.listen(3000, () => {
    console.log("Server is running on port 3000");
});
