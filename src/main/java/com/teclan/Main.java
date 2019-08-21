package com.teclan;

import com.teclan.word.DefaultWordTableRowHandler;
import com.teclan.word.WordTableRowHandler;
import com.teclan.word.WorlUtils;

public class Main {

    public static void main(String[] args){

        String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（新）.doc";

        WordTableRowHandler handler = new DefaultWordTableRowHandler();

        WorlUtils.readTable(filePath,handler);


    }
}
