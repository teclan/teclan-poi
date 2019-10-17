package com.teclan.poi;

import com.teclan.poi.word.DefaultWordTableHandler;
import com.teclan.poi.word.WordTableHandler;
import com.teclan.poi.word.WorlUtils;

public class Main {

    public static void main(String[] args){

        String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（新）.doc";

        WordTableHandler handler = new DefaultWordTableHandler();

        WorlUtils.readTableByRow(filePath,handler);


    }
}
