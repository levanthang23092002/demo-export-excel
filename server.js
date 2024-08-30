const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();

app.use(cors());
app.use(bodyParser.json());

app.post('/api/export', async (req, res) => {
    const data = req.body;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    // Gộp các ô để tạo tiêu đề "Báo Cáo"
    worksheet.mergeCells('A1:Q1'); // Gộp từ cột A đến F cho tiêu đề báo cáo
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'Báo Cáo';
    titleCell.font = { color: { argb: 'FFFF0000' }, bold: true, size: 18 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }, // Màu nền vàng
    };
    worksheet.getRow(1).height = 40;
    

    // Thêm một hàng trống sau tiêu đề báo cáo để đẩy tiêu đề cột xuống
    worksheet.addRow([]);

    // Định nghĩa tiêu đề cột tại dòng thứ ba bắt đầu từ cột B
    worksheet.getCell('B3').value = 'Name';
    worksheet.getCell('C3').value = 'Product Model';
    worksheet.getCell('D3').value = 'Unit';
    worksheet.getCell('E3').value = 'Quantity';
    worksheet.getCell('F3').value = 'Specification';



    worksheet.getRow(3).eachCell({ includeEmpty: true }, (cell) => {
        if (cell.value) { // Kiểm tra xem ô có giá trị hay không
            cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 }; // Màu chữ trắng, đậm
            cell.alignment ={horizontal: 'center', vertical: 'middle'}
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFC107' }, // Màu nền vàng
            };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        }
    });
    worksheet.getRow(3).height = 20;

    // Thay đổi độ rộng cột để phù hợp với tiêu đề
    worksheet.getColumn(2).width = 20;  // 'Name'
    worksheet.getColumn(3).width = 30;  // 'Product Model'
    worksheet.getColumn(4).width = 10;  // 'Unit'
    worksheet.getColumn(5).width = 10;  // 'Quantity'
    worksheet.getColumn(6).width = 30;  // 'Specification'

    // Thêm hàng dữ liệu bắt đầu từ dòng thứ 4 và cột B
    data.forEach(row => {
        const newRow = worksheet.addRow([
            null, // Cột A để trống
            row.name,
            row.productModel,
            row.unit,
            row.quantity,
            row.specification
        ]);

        // Định dạng đường viền cho các ô dữ liệu
        newRow.eachCell({ includeEmpty: true }, (cell) => {
            if (cell.value) {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
            }
        });
    });

    // Lưu file tạm thời
    const filePath = path.join(__dirname, 'data.xlsx');
    await workbook.xlsx.writeFile(filePath);

    // Gửi file về phía client
    res.download(filePath, 'data.xlsx', (err) => {
        if (err) {
            console.log('Error during file download:', err);
            res.status(500).send('Error occurred during file download.');
        }

        // Xóa file sau khi đã tải
        fs.unlink(filePath, (err) => {
            if (err) console.error('Error deleting file:', err);
        });
    });
});

app.listen(5000, () => {
    console.log('Server is running on port 5000');
});
