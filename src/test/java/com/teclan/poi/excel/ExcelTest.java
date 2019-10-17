package com.teclan.poi.excel;

import org.junit.Test;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class ExcelTest {

    String filePath="excel/219182673.xlsx";


    @Test
    public void read(){
        ExcelUtils.readBySheet(filePath,new DefaultExcelHandler());
    }


    @Test
    public  void write(){

        ArrayList<String> title = new ArrayList<String>();
        title.add("序号");
        title.add("姓名");


        LinkedHashMap<String,Object> map = new LinkedHashMap<String,Object>();
        map.put("no",1);
        map.put("name","张三");

        List<LinkedHashMap<String,Object>> rows = new  ArrayList<LinkedHashMap<String,Object>>();
        rows.add(map);

        ExcelUtils.write("excel/test.xlsx","测试",title,rows);
    }


    @Test
    public  void append(){


        LinkedHashMap<String,Object> map = new LinkedHashMap<String,Object>();
        map.put("no",2);
        map.put("name","赵四");

        List<LinkedHashMap<String,Object>> rows = new  ArrayList<LinkedHashMap<String,Object>>();
        rows.add(map);

        ExcelUtils.append("excel/test.xlsx","测试",rows);
    }
}
