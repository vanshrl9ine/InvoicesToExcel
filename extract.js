const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');
const fs = require('fs');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');

const INPUT_FOLDER = './inputs/';
const OUTPUT_CSV = './combined_data.csv';
const OUTPUT_XLSX = './combined_data.xlsx';
const DELAY_MS = 200; // Delay in milliseconds between API requests

// Remove the output files if they already exist
if (fs.existsSync(OUTPUT_CSV)) fs.unlinkSync(OUTPUT_CSV);
if (fs.existsSync(OUTPUT_XLSX)) fs.unlinkSync(OUTPUT_XLSX);

const credentials = PDFServicesSdk.Credentials
  .serviceAccountCredentialsBuilder()
  .fromFile('pdfservices-api-credentials.json')
  .build();

const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);
const inputOptions = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
  .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT)
  .build();

let extractedTexts = [];

fs.readdir(INPUT_FOLDER, (err, files) => {
  if (err) {
    console.log(err);
    return;
  }

  let fileCount = files.length;
  let processedCount = 0;

  const processFile = (fileIndex) => {
    if (fileIndex >= fileCount) {
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.json_to_sheet(extractedTexts);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Extracted Data');
      XLSX.writeFile(workbook, OUTPUT_XLSX);

      console.log(`Successfully combined extracted data into ${OUTPUT_XLSX}`);
      return;
    }

    const file = files[fileIndex];
    const inputPath = INPUT_FOLDER + file;
    const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew();
    const input = PDFServicesSdk.FileRef.createFromLocalFile(
      inputPath,
      PDFServicesSdk.ExtractPDF.SupportedSourceFormat.pdf
    );

    extractPDFOperation.setInput(input);
    extractPDFOperation.setOptions(inputOptions);

    extractPDFOperation.execute(executionContext)
      .then(result => result.saveAsFile(`./ExtractedTextInfo_${fileIndex}.zip`))
      .then(() => {
        let zip = new AdmZip(`./ExtractedTextInfo_${fileIndex}.zip`);
        let jsondata = zip.readAsText('structuredData.json');
        let data = JSON.parse(jsondata);
        let extractedText = '';

        data.elements.forEach(element => {
          extractedText += element.Text + '\n';
        });

        extractedTexts.push({ 'Extracted Text': extractedText });

        fs.unlinkSync(`./ExtractedTextInfo_${fileIndex}.zip`);

        processedCount++;

        setTimeout(() => {
          processFile(fileIndex + 1);
        }, DELAY_MS);
      })
      .catch(err => console.log(err));
  };

  processFile(0);
});
