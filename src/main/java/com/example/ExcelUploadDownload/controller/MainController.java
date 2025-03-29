package com.example.ExcelUploadDownload.controller;

import com.example.ExcelUploadDownload.service.ProductService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class MainController {

    @Autowired
    ProductService productService;

    @GetMapping("/download-excel")
    public ResponseEntity<?> downloadExcel() throws IOException {
        return productService.generateExcel();

//        byte[] excelData = productService.generateExcel();
//        return ResponseEntity.ok()
//                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=products.xlsx")
//                .contentType(MediaType.APPLICATION_OCTET_STREAM)
//                .body(excelData);
    }
}
