package com.naik.uploade.service.impl;

import java.io.IOException;
import java.util.List;

import com.naik.uploade.entity.Product;
import com.naik.uploade.repository.ProductRepo;
import com.naik.uploade.service.ExcelService;
import com.naik.uploade.helper.ExcelHelper;
import com.naik.uploade.helper.ValidationExceptions;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private ProductRepo productrepo;

    public void save(MultipartFile file) throws ValidationExceptions {

        try {
            List<Product> product = ExcelHelper.convertExcelToListOfProduct(file.getInputStream());
            this.productrepo.saveAll(product);

        } catch (IOException e) {

            e.printStackTrace();
        }
    }

    public List<Product> getAllProducts() {

        return this.productrepo.findAll();
    }
}
