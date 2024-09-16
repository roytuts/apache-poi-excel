package com.roytuts.excel.report.generation;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.domain.EntityScan;
import org.springframework.data.jpa.repository.config.EnableJpaRepositories;

@SpringBootApplication
@EntityScan(basePackages = "com.roytuts.excel.report.generation.entity")
@EnableJpaRepositories(basePackages = "com.roytuts.excel.report.generation.repository")
public class SpringExcelReportApp {
	
	public static void main(String[] args) {
		SpringApplication.run(SpringExcelReportApp.class, args);
	}
	
}
