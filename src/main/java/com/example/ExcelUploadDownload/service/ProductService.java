package com.example.ExcelUploadDownload.service;

import com.example.ExcelUploadDownload.entity.Product;
import jakarta.servlet.http.HttpServletRequest;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.List;
import java.util.Map;

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

            Row headerRow = sheet.createRow(0);
            String[] columns = {"ID", "Name", "Price"};
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);

                cell.setCellValue(columns[i]);
                cell.setCellStyle(getHeaderStyle(workbook));
            }

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

    public ResponseEntity<?> generateLinkForExcel(HttpServletRequest request) throws IOException {

        String fileName = "data.xlsx";
        String filePath = System.getProperty("user.dir") + File.separator + fileName;

        System.out.println(filePath);
        List<Product> productList = List.of(
                new Product(1, "Shampoo", 349.0),
                new Product(2, "Shampoo2", 349.0)
        );

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        Row headerRow = sheet.createRow(0);
        String[] columns = {"ID", "Name", "Price"};
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);

            cell.setCellValue(columns[i]);
            cell.setCellStyle(getHeaderStyle(workbook));
        }

        int rowIdx = 1;
        for (Product product : productList) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(product.getId());
            row.createCell(1).setCellValue(product.getName());
            row.createCell(2).setCellValue(product.getPrice());
        }

        FileOutputStream fileOut = new FileOutputStream(filePath);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

        String serverUrl = request.getRequestURL().toString().replace("/generate-link-for-excel", "");
        String downloadUrl = serverUrl + "/files/" + fileName;
        System.out.println(downloadUrl);
        return ResponseEntity.ok(Map.of("Download link ", downloadUrl));

    }

    public ResponseEntity<?> saveExcelFile(MultipartFile file) {

        if (file.isEmpty()) {
            String message = "Please upload a valid Excel file.";
            return ResponseEntity.badRequest().body(Map.of("message", message, "status", HttpStatus.BAD_REQUEST));
        }

        try (
                InputStream inputStream = file.getInputStream();
                Workbook workbook = new XSSFWorkbook(inputStream)
        ) {
            boolean isHeader = true;
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (isHeader) {
                    isHeader = false;
                    continue;
                }

                Product pd = new Product(
                        row.getCell(0).getRowIndex(),
                        row.getCell(1).getStringCellValue(),
                        row.getCell(2).getNumericCellValue()
                );

                //as we don't have database currently that we just print here
                System.out.println("values :" + pd);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return ResponseEntity.ok().body(Map.of("message", "File Uploaded", "status", HttpStatus.OK));
    }
}