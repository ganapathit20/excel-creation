package com.exterro;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.exterro.excel.ExcelCreation;

@SpringBootApplication
public class SpringJpaApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(SpringJpaApplication.class, args);
		
		ExcelCreation.getExcel();
	
	}

}
