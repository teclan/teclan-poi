package com.teclan.excel;

import com.teclan.utils.Objects;
import com.teclan.word.DefaultWordTableHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DefaultExcelHandler implements  ExcelHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(DefaultExcelHandler.class);

    public void handle(int index,Sheet sheet) {

        int LastRow = sheet.getLastRowNum();

        for (int i = 0; i <= LastRow; i++) {
            Row row = sheet.getRow(i);
            this.handle(i,row);
        }

    }

    public void handle(int index,Row row) {

        StringBuffer sb = new StringBuffer();
        for(int i=0;i<row.getLastCellNum();i++) {
            Cell cell = row.getCell(i);
            String value = ExcelUtils.getNumberValue(cell);

            if(Objects.isNotNullString(value)) {
                sb.append(ExcelUtils.getValue(cell)).append("\t");
            }
        }

        LOGGER.info("第 {} 行，{}",index,sb.toString());

    }
}
