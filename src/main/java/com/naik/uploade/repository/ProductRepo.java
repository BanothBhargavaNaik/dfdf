package com.naik.uploade.repository;

import com.naik.uploade.entity.Product;

import org.springframework.data.jpa.repository.JpaRepository;

public interface ProductRepo extends JpaRepository<Product,Long> {
    
}
