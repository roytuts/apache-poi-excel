package com.roytuts.excel.report.generation.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.roytuts.excel.report.generation.entity.Product;

public interface ProductRepository extends JpaRepository<Product, Integer> {

}
