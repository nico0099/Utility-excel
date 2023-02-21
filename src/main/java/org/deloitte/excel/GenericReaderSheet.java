package org.deloitte.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;
import java.util.List;

public interface GenericReaderSheet<T> {

    int getSkipHeaderRows();

    int getIndex();

    T readContent(Workbook workbook, Sheet sheet, Iterator<Row> iterator);
}
