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

import com.naik.uploade.entity.Product;
import com.naik.uploade.validation.EmailValidatorSimple;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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

                        case 1:
                            if (row == null || (((row.getCell(cid) == null || row.getCell(cid).toString().equals("") ||
                                    row.getCell(cid).getCellType() == CellType.BLANK)))) {
                                System.out.println("rowNumber : " + cell.getRowIndex() + " " + "columnNumber : "
                                        + cell.getColumnIndex());

                            } else {
                                p.setProductName(cell.getStringCellValue());

                            }

                            break;
                        case 2:
                            if (row == null || (((row.getCell(cid) == null || row.getCell(cid).toString().equals("") ||
                                    row.getCell(cid).getCellType() == CellType.BLANK)))) {
                                System.out.println("rowNumber : " + cell.getRowIndex() + " " + "columnNumber : "
                                        + cell.getColumnIndex());

                            } else {

                                p.setDescription((row.getCell(2).toString()));
                            }
                            break;
                        case 3:
                            if (row == null || (((row.getCell(cid) == null || row.getCell(cid).toString().equals("") ||
                                    row.getCell(cid).getCellType() == CellType.BLANK)))) {
                                System.out.println("rowNumber : " + cell.getRowIndex() + " " + "columnNumber : "
                                        + cell.getColumnIndex());
                            } else {

                                p.setUnitPrice(Double.parseDouble(row.getCell(3).toString()));
                            }

                            break;
                        case 4:
                            if (row == null || (((row.getCell(cid) == null || row.getCell(cid).toString().equals("") ||
                                    row.getCell(cid).getCellType() == CellType.BLANK)))) {
                                System.out.println("rowNumber : " + cell.getRowIndex() + " " + "columnNumber : "
                                        + cell.getColumnIndex());
                            } else if (EmailValidatorSimple.isValid((row.getCell(cid)).toString())) {

                                p.setEmail((row.getCell(4).toString()));
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
        } catch (

        Exception e) {
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
