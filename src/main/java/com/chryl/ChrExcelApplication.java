package com.chryl;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@MapperScan("com.chryl.**.mapper")
public class ChrExcelApplication {

    public static void main(String[] args) {
        SpringApplication.run(ChrExcelApplication.class, args);
    }

}
