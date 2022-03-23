package com.naik.uploade.service;

import java.util.List;

import com.naik.uploade.entity.Product;

import org.springframework.web.multipart.MultipartFile;

public interface ExcelService {

    public void save(MultipartFile file);

    public List<Product> getAllProducts();
    
}
