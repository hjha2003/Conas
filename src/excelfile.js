import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
const ExcelJS = require("exceljs");

export const ExportToExcel=()=>{

    const downloadExcelFile = async () => {
        // Create a new Excel workbook
        const workbook = new ExcelJS.Workbook();
    
        // Add multiple worksheets with data
        const worksheet1 = workbook.addWorksheet('Sheet 1');
        const worksheet2 = workbook.addWorksheet('Sheet 2');
    
        // Add data to the worksheets
        worksheet1.addRow(['Name', 'Age']);
        worksheet1.addRow(['John Doe', 30]);
        worksheet1.addRow(['Jane Smith', 25]);
    
        worksheet2.addRow(['City', 'Country']);
        worksheet2.addRow(['New York', 'USA']);
        worksheet2.addRow(['London', 'UK']);
    
        // Save the workbook as an Excel file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
    
        // Create a link and click it to trigger the download
        const a = document.createElement('a');
        a.href = url;
        a.download = 'example.xlsx';
        document.body.appendChild(a);
        a.click();
    
        // Clean up the URL object after the download
        URL.revokeObjectURL(url);
      };
    
      const copyDataBetweenWorksheets = async () => {
        const sourceWorkbook = new ExcelJS.Workbook();
        const targetWorkbook = new ExcelJS.Workbook();
    
        // Load the source and target workbooks
        await sourceWorkbook.xlsx.readFile('path/to/source.xlsx');
        await targetWorkbook.xlsx.readFile('path/to/target.xlsx');
    
        // Assuming you want to copy data from Sheet1 of the source workbook to Sheet1 of the target workbook
        const sourceSheet = sourceWorkbook.getWorksheet('Sheet1');
        const targetSheet = targetWorkbook.getWorksheet('Sheet1');
    
        // Get the maximum used row in the source sheet
        const maxRow = sourceSheet.rowCount;
    
        // Loop through each row and copy data to the target sheet
        for (let i = 1; i <= maxRow; i++) {
          const rowValues = sourceSheet.getRow(i).values;
          targetSheet.addRow(rowValues);
        }
    
        // Save the changes to the target workbook
        const buffer = await targetWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
    
        // Create a link and click it to trigger the download
        const a = document.createElement('a');
        a.href = url;
        a.download = 'target.xlsx';
        document.body.appendChild(a);
        a.click();
    
        // Clean up the URL object after the download
        URL.revokeObjectURL(url);
      };
    
      return (
        <div>
          <button onClick={downloadExcelFile}>Download Excel File</button>
          <button onClick={copyDataBetweenWorksheets}>Copy Data Between Worksheets</button>

        </div>
      );

}