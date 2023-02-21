package org.deloitte.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface GenericWriterSheet {
    String getSheetName();

    List<String> getHeader();

    void generateContent(Sheet sheet);
}
