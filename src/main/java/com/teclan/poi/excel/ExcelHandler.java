package com.teclan.poi.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface ExcelHandler {

    public void handle(int index,Sheet sheet);

    public void handle(int index,Row row);
}
