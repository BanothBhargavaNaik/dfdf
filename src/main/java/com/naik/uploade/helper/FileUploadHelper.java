package com.naik.uploade.helper;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;



import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

@Component
public class FileUploadHelper  {

  public final String UPLOAD_DIR = "C:\\Users\\bhargava.banoth\\Downloads\\springMultipart-form-data-main\\springMultipart-form-data-main\\src\\main\\resources\\static\\images";

  public boolean upload(MultipartFile multipartFile) {
    boolean f = false;

    try {

      Files.copy(multipartFile.getInputStream(),
          Paths.get(UPLOAD_DIR + File.separator + multipartFile.getOriginalFilename()),
          StandardCopyOption.REPLACE_EXISTING);

      f = true;

    } catch (Exception e) {
      e.printStackTrace();
    }

    return f;
  }
}
