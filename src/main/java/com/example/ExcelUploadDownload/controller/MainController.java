package com.example.ExcelUploadDownload.controller;

import com.example.ExcelUploadDownload.service.ProductService;
import jakarta.servlet.http.HttpServletRequest;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

@RestController
public class MainController {

    @Autowired
    ProductService productService;

    @GetMapping("/excel-download")
    public ResponseEntity<?> downloadExcel() throws IOException {
        return productService.generateExcel();
    }

    @GetMapping("/generate-link-for-excel")
    public ResponseEntity<?> generateLinkForExcel(HttpServletRequest request) throws IOException {
        return productService.generateLinkForExcel(request);
    }

    @PostMapping(value = "/excel-upload",
            consumes = {MediaType.MULTIPART_FORM_DATA_VALUE},
            produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<?> uploadExcelToSaveCoupons(@RequestParam("file") MultipartFile file) throws IOException {
        return productService.saveExcelFile(file);
    }


}
