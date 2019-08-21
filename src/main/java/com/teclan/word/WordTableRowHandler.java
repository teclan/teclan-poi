package com.teclan.word;

import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public interface WordTableRowHandler {

    public void handler(int index,XWPFTableRow row);

    public void handler(int index,TableRow row);
}
