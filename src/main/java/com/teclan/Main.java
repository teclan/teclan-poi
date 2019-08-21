package com.teclan;

import com.teclan.word.DefaultWordTableHandler;
import com.teclan.word.WordTableHandler;
import com.teclan.word.WorlUtils;

public class Main {

    public static void main(String[] args){

        String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（新）.doc";

        WordTableHandler handler = new DefaultWordTableHandler();

        WorlUtils.readTableByRow(filePath,handler);


    }
}
