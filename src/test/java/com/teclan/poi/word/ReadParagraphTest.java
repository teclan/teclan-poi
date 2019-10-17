package com.teclan.poi.word;

import org.junit.Test;

import java.io.IOException;

public class ReadParagraphTest {

    @Test
    public void read() throws IOException {

        String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（新）.doc";

        WorlUtils.readText(filePath,new DefaultWordTextHandler());
    }
}
