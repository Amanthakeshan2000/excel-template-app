import React from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const ExcelTemplate: React.FC = () => {
  const generateExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Template');

      // Setting column widths
      worksheet.columns = [
        { header: 'No.', key: 'no', width: 10 },
        { header: 'Company Name', key: 'company', width: 40 },
        { header: 'Awarded Daily/Monthly Quantity (MT)', key: 'quantity', width: 30 },
        { header: 'Allocated Monthly Quantity for October wrt Operation Plan (MT)', key: 'allocated', width: 30 },
        { header: 'Unit Selling Price Without Taxes (LKR/MT)', key: 'price', width: 30 },
        { header: 'PI Value With Taxes For October (LKR)', key: 'piValue', width: 30 },
      ];

      // Title and Header Section
      worksheet.mergeCells('A1:F1');
      worksheet.getCell('A1').value = 'LV/S/Fly Ash/004 November 2024 Pro-Forma';
      worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getCell('A1').font = { bold: true, size: 20 };

      worksheet.mergeCells('A2:F2');
      worksheet.getCell('A2').value = 'Annexure 01';
      worksheet.getCell('A2').alignment = { horizontal: 'right', vertical: 'middle' };
      worksheet.getCell('A2').font = { bold: true };

      worksheet.addRow({}); 

      // Header Row
      const headerRow = worksheet.addRow(['No.', 'Company Name', 'Awarded Daily/Monthly Quantity (MT)', 'Allocated Monthly Quantity for October wrt Operation Plan (MT)', 'Unit Selling Price Without Taxes (LKR/MT)', 'PI Value With Taxes For October (LKR)']);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9D9D9' }, 
        };
      });

      
      worksheet.mergeCells('A5:F5');
      const sectionHeader1 = worksheet.getCell('A5');
      sectionHeader1.value = 'Quality No.01 Buyers';
      sectionHeader1.font = { bold: true, size: 12 };
      sectionHeader1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      sectionHeader1.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' }, 
      };

      // Adding Data Rows
      const data1 = [
        { no: 1, company: 'Century Zone (Pvt) Ltd.', quantity: 250, allocated: 250, price: 8050, piValue: 2435641.03 },
        { no: 2, company: 'Fine Ash (Pvt) Ltd.', quantity: 2299.8, allocated: 1800, price: 6800, piValue: 14813538.46 },
        { no: 3, company: 'Betton Technology Lanka (Pvt.) Ltd.', quantity: 300, allocated: 300, price: 6551, piValue: 2378516.92 },
        { no: 4, company: 'Land Reclamation & Development Company Ltd.', quantity: 3600, allocated: 3600, price: 3117, piValue: 13580529.23 },
        { no: 5, company: 'Tokyo Eastern Cement Company (Pvt.) Limited.', quantity: 17410.6, allocated: 15680.15, price: 2950, piValue: 55982156.05 },
      ];
      data1.forEach((row) => {
        const newRow = worksheet.addRow(row);
        newRow.eachCell((cell) => {
          cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      // Section: Quality No.02 Buyers
      const lastRowNumber = worksheet.lastRow ? worksheet.lastRow.number : 0;
      worksheet.mergeCells(`A${lastRowNumber + 2}:F${lastRowNumber + 2}`);
      const sectionHeader2 = worksheet.getCell(`A${lastRowNumber + 2}`);
      sectionHeader2.value = 'Quality No.02 Buyers';
      sectionHeader2.font = { bold: true, size: 12 };
      sectionHeader2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      sectionHeader2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' }, // Yellow background
      };

      const data2 = [
        { no: 6, company: 'Adam Carbons Ltd.', quantity: 250, allocated: 250, price: 6052, piValue: 1831117.95 },
        { no: 7, company: 'Ceylon & Foreign Trades PLC.', quantity: 250, allocated: 250, price: 5786, piValue: 1750635.9 },
        { no: 8, company: 'Betton Technology Lanka (Pvt.) Ltd.', quantity: 250, allocated: 250, price: 5550, piValue: 1679230.77 },
      ];
      data2.forEach((row) => {
        const newRow = worksheet.addRow(row);
        newRow.eachCell((cell) => {
          cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      // Footer
      worksheet.addRow({});
      worksheet.addRow(['Prepared by', '', '', '', '', 'Approved by']);
      const footerRow = worksheet.lastRow;
      if (footerRow) {
        footerRow.eachCell((cell) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
      }

      // Generate and download the Excel file
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/octet-stream' });
      saveAs(blob, 'ExcelTemplate.xlsx');
    } catch (error) {
      console.error('Error generating Excel file:', error);
    }
  };

  return (
    <div>
      <h1>Excel Template Generator</h1>
      <button onClick={generateExcel}>Download Excel Template</button>
    </div>
  );
};

export default ExcelTemplate;
