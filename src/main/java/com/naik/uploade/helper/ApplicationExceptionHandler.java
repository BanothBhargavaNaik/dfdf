package com.naik.uploade.helper;

import com.naik.uploade.entity.AppError;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class ApplicationExceptionHandler {

    @ExceptionHandler(ValidationExceptions.class)
    public ResponseEntity<AppError> validationExceptionHandeler(ValidationExceptions ex) {
        AppError error = new AppError();
        error.setDescrpition(ex.getMessage());
        return ResponseEntity.badRequest().body(error);
    }

}
