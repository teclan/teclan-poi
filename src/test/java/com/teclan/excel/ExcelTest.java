package com.teclan.excel;

import org.junit.Test;

public class ExcelTest {

    String filePath="excel/219182673.xlsx";


    @Test
    public void read(){
        ExcelUtils.readBySheet(filePath,new DefaultExcelHandler());
    }
}
