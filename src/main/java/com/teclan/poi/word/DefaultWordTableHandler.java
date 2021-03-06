package com.teclan.poi.word;

import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class DefaultWordTableHandler implements WordTableHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(DefaultWordTableHandler.class);


    public void handle(int index, XWPFTableRow row) {
        List<XWPFTableCell> cells = row.getTableCells();
        StringBuffer sb = new StringBuffer();
        for (int j = 0; j < cells.size(); j++) {
            XWPFTableCell cell = cells.get(j);
            sb.append(cell.getText());
        }

        LOGGER.info("行号：{}，内容：{}",index,sb.toString());
    }

    public void handle(int index, TableRow row) {

        StringBuffer sb = new StringBuffer();

        for (int j = 0; j < row.numCells(); j++) {
            TableCell td = row.getCell(j);//取得单元格
            //取得单元格的内容
            for (int k = 0; k < td.numParagraphs(); k++) {
                Paragraph para = td.getParagraph(k);
                String text = para.text();
                //去除后面的特殊符号
                if (null != text && !"".equals(text)) {
                    text = text.substring(0, text.length() - 1);
                }
                sb.append(text);
            }
        }

        LOGGER.info("行号：{}，内容：{}",index,sb.toString());

    }

    public void handle(int index, XWPFTable table) {

        List<XWPFTableRow> rows = table.getRows();
        //读取每一行数据
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow  row = rows.get(i);
            this.handle(index,row);
        }
    }

    public void handle(int index, Table table) {

        for (int i = 0; i < table.numRows(); i++) {
            TableRow row = table.getRow(i);
           this.handle(index,row);
        }

    }
}
