package com.teclan.poi.excel;

import java.io.*;

import java.math.BigDecimal;
import java.util.LinkedHashMap;
import java.util.List;

import com.teclan.poi.utils.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ExcelUtils {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 读取 Excel 所有内容,按 行 为单位
     *
     * @param filePath
     * @param handler
     */
    public static void readByRow(String filePath, ExcelHandler handler) {
        Workbook wb = null;
        InputStream input = null;
        File file = null;
        try {
            file = new File(filePath);
            input = new FileInputStream(file);
        } catch (IOException e1) {
            LOGGER.error(e1.getMessage(), e1);
        }
        try {
            wb = WorkbookFactory.create(input);
        } catch (Exception e1) {
            LOGGER.error(e1.getMessage(), e1);
        }

        try {
            for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = wb.getSheetAt(sheetNum);
                LOGGER.info("sheet {},name:{}", sheetNum, sheet.getSheetName());

                int LastRow = sheet.getLastRowNum();

                for (int i = 0; i <= LastRow; i++) {
                    Row row = sheet.getRow(i);
                    handler.handle(i, row);
                }
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }

    }

    /**
     * 读取Excel所有内容，按 sheet 页为单位
     *
     * @param filePath excel 文件路径
     * @param handler
     */
    public static void readBySheet(String filePath, ExcelHandler handler) {

        Workbook wb = null;
        InputStream input = null;
        File file = null;
        try {
            file = new File(filePath);
            input = new FileInputStream(file);
        } catch (IOException e1) {
            LOGGER.error(e1.getMessage(), e1);
        }
        try {
            wb = WorkbookFactory.create(input);
        } catch (Exception e1) {
            LOGGER.error(e1.getMessage(), e1);
        }

        try {
            for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = wb.getSheetAt(sheetNum);
                handler.handle(sheetNum, sheet);
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {

            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage(), e);
                }
            }

        }
    }

    /**
     * 读取指定sheet 页
     *
     * @param filePath excel 文件路径
     * @param sheetNo  指定的sheet页，从 0 开始计
     * @param handler
     */
    public static void readBySheet(String filePath, int sheetNo, ExcelHandler handler) {

        Workbook wb = null;
        InputStream input = null;
        File file = null;
        try {
            file = new File(filePath);
            input = new FileInputStream(file);
        } catch (IOException e1) {
            LOGGER.error(e1.getMessage(), e1);
        }
        try {
            wb = WorkbookFactory.create(input);
        } catch (Exception e1) {
            LOGGER.error(e1.getMessage(), e1);
        }

        try {
            Sheet sheet = wb.getSheetAt(sheetNo);
            handler.handle(sheetNo, sheet);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {

            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage(), e);
                }
            }

        }
    }


    /**
     * 写入 excel
     * @param filePath excel 文件路径
     * @param sheetName sheet 页名称
     * @param title 标题，即每个列的字段名称
     * @param rows 所有行的内容
     */
    public static void write(String filePath, String sheetName,List<String> title, List<LinkedHashMap<String, Object>> rows) {

        XSSFWorkbook wb = new XSSFWorkbook();

        XSSFSheet sheet = null;

        if (Objects.isNotNullString(sheetName)) {
            sheet = wb.createSheet(sheetName);
        } else {
            sheet = wb.createSheet();
        }

        XSSFRow titleRow = sheet.createRow(0);
        int index=0;
        for(String t:title){
            XSSFCell cell = titleRow.createCell(index++);
            cell.setCellValue(t);
        }

        for (int i = 0; i < rows.size(); i++) {
            createRow(sheet,i+1,rows.get(i));
        }
        FileOutputStream output = null;
        try {
            output = new FileOutputStream(filePath);
            wb.write(output);
            output.flush();

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            if (output != null) {
                try {
                    output.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage(), e);
                }
            }
        }
    }


    private static void createRow(XSSFSheet sheet,int rowIndex,LinkedHashMap<String, Object> row){
        XSSFRow r = sheet.createRow(rowIndex);
        int cellIndex=0;
        for (String key : row.keySet()) {
            XSSFCell cell = r.createCell(cellIndex++);
            cell.setCellValue(row.get(key) == null ? "" : row.get(key).toString());
        }
    }

    /**
     * Excel 追加内容
     * @param filePath
     * @param sheetName
     * @param rows
     */
    public static void append(String filePath,String sheetName, List<LinkedHashMap<String, Object>> rows) {

        FileInputStream is=null;
        XSSFWorkbook wb=null;

        try {
            is = new FileInputStream(filePath);

            wb = new XSSFWorkbook(is);

            if(wb==null){
                throw new Exception(String.format("Excel读取错误,文件:%s ",sheetName,filePath));
            }

            XSSFSheet sheet = wb.getSheet(sheetName);

            if(sheet==null){
                throw new Exception(String.format("名为 %s 的 sheet 页不存在,文件:%s ",sheetName,filePath));
            }

            int index = sheet.getLastRowNum();
            for (int i = 0; i < rows.size(); i++) {
                createRow(sheet,index+i+1,rows.get(i));
            }

        }catch (Exception e){
            LOGGER.error(e.getMessage(), e);
        }finally {
            if(is!=null){
                try {
                    is.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage(), e);
                }
            }
        }


        FileOutputStream output = null;
        try {
            output = new FileOutputStream(filePath);
            wb.write(output);
            output.flush();

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            if (output != null) {
                try {
                    output.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage(), e);
                }
            }
        }
    }

    public static String getValue(Cell cell) {

        if (Objects.isNull(cell)) {
            return "";
        }
        cell.setCellType(Cell.CELL_TYPE_STRING);
        if (Objects.isNullString(cell)) {
            return "";
        }
        return cell.getStringCellValue().trim();
    }

    public static String getPhoneValueString(Cell cell) {

        if (Objects.isNull(cell)) {
            return "";
        }
        if (Objects.isNullString(cell)) {
            return "";
        }

        try {
            BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
            return bd.toPlainString();
        } catch (Exception e) {
            return getValue(cell);
        }
    }

    public static String getNumberValue(Cell cell) {

        if (Objects.isNull(cell)) {
            return "";
        }
        if (Objects.isNullString(cell)) {
            return "";
        }

        try {
            BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
            return bd.toPlainString();
        }catch (Exception e){
            return getValue(cell);
        }
    }

}
