package com.naik.uploade.helper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.naik.uploade.entity.Product;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
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
    public static List<Product> convertExcelToListOfProduct(InputStream is) {
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
                    } else {
                        break;
                    }
                }

                Iterator<Cell> cells = row.iterator();

                int cid = 0;

                Product p = new Product();

                while (cells.hasNext()) {
                    Cell cell = cells.next();

                    switch (cid) {
                        case 0:
                            if (cell.getNumericCellValue() != 0) {

                                p.setId((int) cell.getNumericCellValue());
                               
                            } else {
                                System.out.println("rowNumber : " + rowNumber + " " + "columnNumber : " + cid);
                            }
                            break;
                        case 1:
                            if (!cell.getStringCellValue().equals(null)) {

                                p.setProductName(cell.getStringCellValue());
                            } else {
                                System.out.println("rowNumber : " + rowNumber + " " + "columnNumber : " + cid);
                            }
                            break;
                        case 2:
                            if (!cell.getStringCellValue().equals(null)) {

                                p.setDescription(cell.getStringCellValue());
                            } else {
                                System.out.println("rowNumber : " + rowNumber + " " + "columnNumber : " + cid);
                            }
                            break;
                        case 3:
                            if (cell.getNumericCellValue() != 0) {

                                p.setUnitPrice(cell.getNumericCellValue());
                            } else {
                                System.out.println("rowNumber : " + rowNumber + " " + "columnNumber : " + cid);
                            }
                            break;
                        case 4:
                            if (EmailValidatorSimple.isValid(cell.getStringCellValue())) {
                                p.setEmail(cell.getStringCellValue());
                            } else {
                                System.out.println("rowNumber : " + rowNumber + " " + "columnNumber : " + cid);
                            }
                            break;
                        default:
                            break;
                    }
                    cid++;
                }
                list.add(p);

                workbook.close();

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;

    }

    // convert excel to list of products error sheet
    public void excelErrorSheet(String excelpath, String sheetName, int rowNumber, int columnNumber, String data) {

        try {
            File file = new File(excelpath);

            FileInputStream fis = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(fis);

            XSSFSheet sheet = wb.getSheet(sheetName);

            XSSFRow rowstart = sheet.getRow(0);

            XSSFCell celstart = rowstart.getCell(0);

            celstart.setCellValue("Issue Raised");

            XSSFRow row = sheet.getRow(rowNumber);

            XSSFCell cel = row.getCell(columnNumber);

            cel.setCellValue(data);

            FileOutputStream fio = new FileOutputStream(file);
            wb.write(fio);
            wb.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
