package com.naik.uploade.controller;

import java.util.List;
import java.util.Map;

import com.naik.uploade.entity.Product;
import com.naik.uploade.helper.ExcelHelper;
import com.naik.uploade.helper.FileUploadHelper;
import com.naik.uploade.service.ExcelService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("api/v1")
public class FileUploadController {

    @Autowired
    private final FileUploadHelper fileUploadHelper;

    @Autowired
    private final ExcelService excelService;

    public FileUploadController(FileUploadHelper fileUploadHelper, ExcelService excelService) {
        this.fileUploadHelper = fileUploadHelper;
        this.excelService = excelService;
    }

    // This method to upload data locally .
    @PostMapping("/upload-file")
    public ResponseEntity<String> uploadFile(@RequestParam("file") MultipartFile file) {
        try {
            boolean f = fileUploadHelper.upload(file);
            if (f) {
                return ResponseEntity.ok("File is Uploaded Successfully...!");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Something went wrong ... Try again");
    }

    // Save data to data base from Excel sheet
    @PostMapping("/product/upload")
    public ResponseEntity<?> upload(@RequestParam("file") MultipartFile file) {
        if (ExcelHelper.checkExcelFormat(file)) {
            this.excelService.save(file);
            return ResponseEntity.ok(Map.of("message", "File is uploaded and data is saved to db"));
        }
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Please upload excel file ");
    }

    // Retrieve all data
    @GetMapping("/product")
    public List<Product> getAllProduct() {
        return this.excelService.getAllProducts();
    }
}
