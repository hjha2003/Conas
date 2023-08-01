import React from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

class ExcelCopy extends React.Component {
  handleFileUpload = async (event) => {
    const file = event.target.files[0];
    const sourceWorkbook = new ExcelJS.Workbook();
    const destinationWorkbook = new ExcelJS.Workbook();

    try {
      // Load the source workbook
      await sourceWorkbook.xlsx.load(file);

      // Create a new worksheet in the destination workbook
      const destinationWorksheet = destinationWorkbook.addWorksheet("Copied Data");

      // Get the source worksheet
      const sourceWorksheet = sourceWorkbook.getWorksheet("Sheet1");

      // Copy the data from the source worksheet to the destination worksheet
      sourceWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          destinationWorksheet.getCell(rowNumber, colNumber).value = cell.value;
        });
      });

      // Save the destination workbook
      const buffer = await destinationWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, "copied_workbook.xlsx");

      // Clear the input file element value
      event.target.value = "";
    } catch (error) {
      console.error("Error copying data:", error);
    }
  };

  render() {
    return (
      <div>
        <input type="file" onChange={this.handleFileUpload} />
      </div>
    );
  }
}

export default ExcelCopy;
