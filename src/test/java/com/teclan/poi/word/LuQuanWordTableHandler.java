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

import java.util.ArrayList;
import java.util.List;

public class LuQuanWordTableHandler implements WordTableHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(LuQuanWordTableHandler.class);

    List<Person> people = new ArrayList<Person>();

    public List<Person> getPeople(){
        return people;
    }

    public void handle(int index, XWPFTableRow row) {
        List<XWPFTableCell> cells = row.getTableCells();
        StringBuffer sb = new StringBuffer();
        for (int j = 0; j < cells.size(); j++) {
            XWPFTableCell cell = cells.get(j);
            sb.append(cell.getText());
        }
    }

    public void handle(int index, TableRow row) {

        StringBuffer sb = new StringBuffer();
        Person person = new Person();

        person.setName(getCell(row,1));

        if("".equals(person.getName())||"姓名".equals(person.getName())){
            return;
        }

        person.setSex(getCell(row,2));
        person.setType(getCell(row,3));
        person.setIdCard(getCell(row,4));
        person.setLocation(getCell(row,5));
        person.setEnterTime(getCell(row,6));
        person.setOutTime(getCell(row,7));
        person.setContNo(getCell(row,8));
        person.setWage(getCell(row,9));
        person.setCardNo(getCell(row,10));
        person.setBank(getCell(row,11));
        person.setPhone(getCell(row,12));
        person.setMemo(getCell(row,13));

        people.add(person);

        LOGGER.info("表{},姓名:{},身份证:{}",index,person.getName(),person.getIdCard());

    }

    private String getCell(TableRow row, int index) {

        TableCell cell = row.getCell(index);

        Paragraph para = cell.getParagraph(0);
        String text = para.text();

        if (null != text && !"".equals(text)) {
            text = text.substring(0, text.length() - 1);
        }

        return text.trim();
    }

    public void handle(int index, XWPFTable table) {

        List<XWPFTableRow> rows = table.getRows();
        //读取每一行数据
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            this.handle(index, row);
        }
    }

    public void handle(int index, Table table) {
        for (int i = 0; i < table.numRows(); i++) {
            TableRow row = table.getRow(i);
            this.handle(index, row);
        }
    }
}