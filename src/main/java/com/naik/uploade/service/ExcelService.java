package com.naik.uploade.service;

import java.util.List;

import com.naik.uploade.entity.Product;
import com.naik.uploade.helper.ValidationExceptions;

import org.springframework.web.multipart.MultipartFile;

public interface ExcelService {

    public void save(MultipartFile file) throws ValidationExceptions;

    public List<Product> getAllProducts();
    
}
