package com.teclan.poi.word;

import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public interface WordTableHandler {

    public void handle(int index, XWPFTableRow row);

    public void handle(int index, TableRow row);

    public void handle(int index, XWPFTable table);

    public void handle(int index, Table table);
}
