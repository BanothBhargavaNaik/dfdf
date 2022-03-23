package com.naik.uploade.service.impl;

import java.util.List;

import com.naik.uploade.entity.Product;
import com.naik.uploade.repository.ProductRepo;
import com.naik.uploade.service.ProductService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

@Service
public class ProductServiceImpl implements ProductService {

    @Autowired
    private final ProductRepo productRepo;

    public ProductServiceImpl(ProductRepo productRepo) {
        this.productRepo = productRepo;
    }

    @Override
    public List<Product> findProductswithSortingAsc(String field) {

        return productRepo.findAll(Sort.by(Sort.Direction.ASC, field));
    }

    @Override
    public List<Product> findProductswithSortingDes(String field) {

        return productRepo.findAll(Sort.by(Sort.Direction.DESC, field));
    }

    @Override
    public Page<Product> findProductWithPagination(int ofset, int pageSize) {
        Page<Product> pag =  productRepo.findAll(PageRequest.of(ofset, pageSize));
        return pag;
    }

    

}
