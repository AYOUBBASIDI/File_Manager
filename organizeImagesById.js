const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

function processImages(excelFilePath, imagesFolderPath, outputFolderPath) {
    const workbook = XLSX.readFile(excelFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);
    const totalItems = data.length;
    let processedItems = 0;

    data.forEach(item => {
        const id = item['id'] + '';
        const files = fs.readdirSync(imagesFolderPath);

        const matchingImages = files.filter(file => {
            const fileName = path.basename(file, path.extname(file));
            return fileName.includes(id);
        });
        if (matchingImages.length > 0) {
            const outputFolder = path.join(outputFolderPath, id);
            if (!fs.existsSync(outputFolder)) {
                fs.mkdirSync(outputFolder);
            }

            matchingImages.forEach(image => {
                const sourcePath = path.join(imagesFolderPath, image);
                const destinationPath = path.join(outputFolder, image);
                fs.copyFileSync(sourcePath, destinationPath);
            });

            console.log(`Images with ID ${id} processed.`);
        } else {
            console.log(`Images with ID ${id} not found. Skipping.`);
        }

        processedItems++;
        const progress = (processedItems / totalItems) * 100;
        console.log(`Progress: ${progress.toFixed(2)}%`);
    });
}

const excelFilePath = 'storage/excel/input2.xlsx';
const imagesFolderPath = 'Z:\\images';
const outputFolderPath = 'Z:\\images_filtered';

processImages(excelFilePath, imagesFolderPath, outputFolderPath);
