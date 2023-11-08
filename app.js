const express = require('express');
const fs = require('fs-extra');
const xlsx = require('xlsx');
const path = require('path');
const multer = require('multer');
const bodyParser = require('body-parser');
const PORT = process.env.PORT || 3000;

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.use(express.static('public'));


app.use('/assets', express.static(path.join(__dirname, 'assets')));

app.get('/', (req, res) => {
  res.render('index');
});

app.get('/success', (req, res) => {
  res.render('success'); // Render the success.ejs view
});

app.post('/upload', upload.single('excelFile'), (req, res) => {


  if (!req.file || !fs.existsSync(req.file.path)) {
    return res.render('index', { error: 'Uploaded file not found. Please try again.' });
  }

  const { storageFolder, outputExcelPath } = req.body;

  if (!storageFolder || storageFolder == "") {
    return res.render('index', { error: 'Storage folder not found. Please try again.' });
  }

  if (!outputExcelPath || outputExcelPath == "") {
    return res.render('index', { error: 'Output folder not found. Please try again.' });
  }

  const outputFolder = outputExcelPath +'/generated_folders';
  const inputExcelPath = path.join(__dirname, 'uploads', req.file.filename);
  const workbook = xlsx.readFile(inputExcelPath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const filesNotFound = []; // Array to store names of files not found

  // Rest of your processing code...
  // Create a directory for generated folders
  fs.ensureDirSync(outputFolder);

// Initialize a variable to store the output data for the Excel file
const outputData = [];

// Iterate through Excel rows and process the data
const excelData = xlsx.utils.sheet_to_json(worksheet, { header: ['id', 'images'] });

excelData.forEach(row => {
  const id = row.id;
  const images = row.images.split(',');

  // Create a directory for the current ID
  const idFolderPath = `${outputFolder}/${id}`;
  fs.ensureDirSync(idFolderPath);

  // Move images to the corresponding ID folder and update the output data
  const rowData = {
    id: id,
    images: row.images,
  };

  images.forEach(image => {
    const imageName = image.trim();
    const sourcePath = `${storageFolder}/${imageName}`;
    const destinationPath = `${idFolderPath}/${imageName}`;

    if (fs.existsSync(sourcePath)){
      fs.copyFileSync(sourcePath, destinationPath);
      console.log(`Moved ${imageName} to folder ${id}`);
    }else{
      console.log(`File ${imageName} not found at ${sourcePath}`);
    }

    // Add the image name to the output data row
    rowData[`Image ${images.indexOf(image) + 1}`] = path.resolve(destinationPath);
  });

  images.forEach(image => {
    const imageName = image.trim();
    const sourcePath = `${storageFolder}/${imageName}`;

    if (!fs.existsSync(sourcePath)) {
      filesNotFound.push(imageName);
      console.log(`File ${imageName} not found at ${sourcePath}`);
    }
  });

  // Add the processed row data to the output data array
  outputData.push(rowData);
});

fs.writeFileSync(outputFolder+'/not_found.txt', filesNotFound.join('\n'));

// Create a new workbook and write data to the output Excel file
const outputWorkbook = xlsx.utils.book_new();
const outputWorksheet = xlsx.utils.json_to_sheet(outputData);
xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Sheet1');
xlsx.writeFile(outputWorkbook, outputFolder+'/output.xlsx');

  // Send a response to the user when processing is completed
   
    res.redirect('/success') 
});



app.listen(PORT, () => {
  console.log('Server started');
});
