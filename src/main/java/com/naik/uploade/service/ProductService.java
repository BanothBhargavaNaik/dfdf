package com.naik.uploade.service;

import java.util.List;

import com.naik.uploade.entity.Product;

import org.springframework.data.domain.Page;

public interface ProductService {
        
    public List<Product> findProductswithSortingAsc(String field);
    public List<Product> findProductswithSortingDes(String field);
    public Page<Product> findProductWithPagination(int ofset, int pageSize);
}
