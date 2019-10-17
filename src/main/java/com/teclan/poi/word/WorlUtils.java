package com.teclan.poi.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

public class WorlUtils {
    private static final Logger LOGGER = LoggerFactory.getLogger(WorlUtils.class);

    public static void readTableByRow(String filePath, WordTableHandler handler){
        try{
            FileInputStream in = new FileInputStream(filePath);//载入文档
            // 处理docx格式 即office2007以后版本
            if(filePath.toLowerCase().endsWith("docx")){
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                XWPFDocument xwpf = new XWPFDocument(in);//得到word文档的信息
                Iterator<XWPFTable> it = xwpf.getTablesIterator();//得到word中的表格

                int index = 0;

                while(it.hasNext()){
                    LOGGER.info("\n正在处理第 {} 个表格 ...\n",++index);

                    XWPFTable table = it.next();

                    List<XWPFTableRow> rows = table.getRows();
                    //读取每一行数据
                    for (int i = 0; i < rows.size(); i++) {
                        XWPFTableRow  row = rows.get(i);
                        handler.handle(index,row);
                    }
                    LOGGER.info("\n第 {} 个表格处理结束 ...\n",++index);
                }
            }else{
                // 处理doc格式 即office2003版本
                POIFSFileSystem pfs = new POIFSFileSystem(in);
                HWPFDocument hwpf = new HWPFDocument(pfs);
                Range range = hwpf.getRange();//得到文档的读取范围
                TableIterator it = new TableIterator(range);

                int index = 0;
                while (it.hasNext()) {

                    LOGGER.info("\n正在处理第 {} 个表格 ...\n",++index);
                    Table tb = (Table) it.next();
                    for (int i = 0; i < tb.numRows(); i++) {
                        TableRow row = tb.getRow(i);
                        handler.handle(index,row);
                    }
                    LOGGER.info("\n第 {} 个表格处理结束 ...\n",++index);
                }
            }
        }catch(Exception e){
            LOGGER.error(e.getMessage(),e);
        }
    }


    public static void readTable(String filePath, WordTableHandler handler){
        try{
            FileInputStream in = new FileInputStream(filePath);//载入文档
            // 处理docx格式 即office2007以后版本
            if(filePath.toLowerCase().endsWith("docx")){
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                XWPFDocument xwpf = new XWPFDocument(in);//得到word文档的信息
                Iterator<XWPFTable> it = xwpf.getTablesIterator();//得到word中的表格

                int index = 0;

                while(it.hasNext()){
                    LOGGER.info("\n正在处理第 {} 个表格 ...\n",++index);

                    XWPFTable table = it.next();

                    handler.handle(index,table);
                    LOGGER.info("\n第 {} 个表格处理结束 ...\n",++index);
                }
            }else{
                // 处理doc格式 即office2003版本
                POIFSFileSystem pfs = new POIFSFileSystem(in);
                HWPFDocument hwpf = new HWPFDocument(pfs);
                Range range = hwpf.getRange();//得到文档的读取范围
                TableIterator it = new TableIterator(range);

                int index = 0;
                while (it.hasNext()) {

                    LOGGER.info("\n正在处理第 {} 个表格 ...\n",++index);
                    Table table = (Table) it.next();
                    handler.handle(index,table);
                    LOGGER.info("\n第 {} 个表格处理结束 ...\n",++index);
                }
            }
        }catch(Exception e){
            LOGGER.error(e.getMessage(),e);
        }
    }

    public static void readText(String filePath,WordTextHandler handler) throws IOException {

        InputStream is = new FileInputStream(filePath);
        WordExtractor ex = new WordExtractor(is);

        String text = ex.getText();

        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                LOGGER.error(e.getMessage(),e);
            }
        }

        handler.handler(text);
    }
}
