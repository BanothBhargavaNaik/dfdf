package com.naik.uploade.helper;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.ValidationException;

import com.naik.uploade.entity.Product;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

@Component
public class ExcelHelper {

    // To check That file is Excel formate or not
    public static boolean checkExcelFormat(MultipartFile file) {

        String contentType = file.getContentType();

        if (contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
            return true;
        } else {
            return false;
        }

    }

    // Email validation method
    public static class EmailValidatorSimple {

        private static final String EMAIL_PATTERN = "^(.+)@(\\S+)$";

        private static final Pattern pattern = Pattern.compile(EMAIL_PATTERN);

        public static boolean isValid(final String email) {
            Matcher matcher = pattern.matcher(email);
            return matcher.matches();
        }

    }



    // convert excel to list of products
    public static List<Product> convertExcelToListOfProduct(InputStream is) throws ValidationExceptions, IOException {
        List<Product> list = new ArrayList<>();

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(is);
            XSSFSheet sheet = workbook.getSheet("data");
            int rowNumber = 0;
            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row row = iterator.next();

                // validating excel header file

                if (rowNumber == 0) {
                    if (row.getCell(0).toString().equals("id") && row.getCell(1).toString().equals("productName")
                            && row.getCell(2).toString().equals("description")
                            && row.getCell(3).toString().equals("unitPrice")
                            && row.getCell(4).toString().equals("email")) {
                        rowNumber++;
                        continue;
                    } 
            } else{
                throw new ValidationException("Header cell is not matched.");
            }

                Iterator<Cell> cells = row.iterator();
                int cid = 0;
                Product p = new Product();

                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    

                    switch (cid) {

                        case 1:
                            p.setProductName(row.getCell(cid).toString());
                            break;

                        case 2:
                            p.setDescription(row.getCell(cid).toString());
                            break;

                        case 3:
                            p.setUnitPrice(Double.parseDouble(row.getCell(cid).toString()));
                            break;

                        case 4:
                            if (EmailValidatorSimple.isValid(row.getCell(cid).toString()))
                                p.setEmail(row.getCell(cid).toString());
                            break;
                        default:
                            break;
                    }
                    cid++;
                }
                list.add(p);
                workbook.close();
            }
        } catch (ValidationException e1) {
            try {
                throw new ValidationException("Header cell is not matched.");
            } catch (Exception e) {
                
                e1.printStackTrace();
            }

        }
        return list;
    }

}


// if (row == null || (((row.getCell(cid) == null || row.getCell(cid).toString().equals("") ||
//                                     row.getCell(cid).getCellType() == CellType.BLANK)))) 