package com.shenunicorn.template;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

@EnableAutoConfiguration
@SpringBootApplication 
public class Application extends SpringBootServletInitializer {
    private static final Logger LOGGER = LogManager.getLogger(Application.class);

    public static void main(String[] args) {
        LOGGER.info("################## Starting Spring Boot application.... ##################");
        SpringApplication.run(Application.class, args);
		LOGGER.info("################## Spring boot started. ##################");
    }
}
