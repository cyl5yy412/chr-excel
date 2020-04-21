package com.chryl;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.chryl.mapper")
public class ChrExcelApplication {

    public static void main(String[] args) {
        SpringApplication.run(ChrExcelApplication.class, args);
    }

}
