package com.example.ExcelUploadDownload.service;

import com.example.ExcelUploadDownload.entity.Product;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ProductService {

    public ResponseEntity<?> generateExcel() throws IOException {
        try {
            List<Product> productList = List.of(
                    new Product(1, "Shampoo", 349.0),
                    new Product(2, "Shampoo2", 349.0)
            );

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Products");

            // Creating Header Row
            Row headerRow = sheet.createRow(0);
            String[] columns = {"ID", "Name", "Price"};
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);

                cell.setCellValue(columns[i]);
                cell.setCellStyle(getHeaderStyle(workbook));
            }

            // Filling Data Rows
            int rowIdx = 1;
            for (Product product : productList) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(product.getId());
                row.createCell(1).setCellValue(product.getName());
                row.createCell(2).setCellValue(product.getPrice());
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            workbook.close();

            byte[] byteArrayForExcel = out.toByteArray();

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDisposition(ContentDisposition.attachment().filename("Product List.xlsx").build());

            return new ResponseEntity<>(byteArrayForExcel, headers, HttpStatus.OK);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

    // For Styling the Header
    private CellStyle getHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }
}
