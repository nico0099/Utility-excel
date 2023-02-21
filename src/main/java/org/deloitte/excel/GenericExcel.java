package org.deloitte.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public abstract class GenericExcel<T> {

    private static final int COLUMN_WIDTH = 30;
    private static final String DATE_FORMAT = "dd-MM-yyyy";
    private static final String DATE_TIME_FORMAT = "dd-MM-yyyy HH:mm";
    private static final DateTimeFormatter localDateFormatter = DateTimeFormatter.ofPattern(DATE_FORMAT);
    private static final DateTimeFormatter localDateTimeFormatter = DateTimeFormatter.ofPattern(DATE_TIME_FORMAT);

    protected List<T> sheets = new ArrayList<>();

    public abstract void initializeSheet();

    Sheet generateSheet(Workbook workbook, String sheetName) {
        return workbook.createSheet(sheetName);
    }

    void generateHeader(Workbook workbook, Sheet sheet, List<String> headers) {
        Row row = sheet.createRow(0);
        int colNum = 0;
        for(String header: headers) {
            Cell cell = row.createCell(colNum++);
            Font font = workbook.createFont();
            font.setBold(true);
            CellStyle style = workbook.createCellStyle();
            style.setFont(font);
            cell.setCellStyle(style);
            cell.setCellValue(header);
        }
    }

    /**
     * Read list of all rows in excel file
     *
     * @param workbook: workbook created
     * @return list of all rows in excel file
     */
    public List<Object> readFile(XSSFWorkbook workbook) {
        List<Object> objList = new ArrayList<>();
        this.sheets.clear();
        this.initializeSheet();

        for (T t : this.sheets) {
            GenericReaderSheet genericReaderSheet = (GenericReaderSheet) t;
            Sheet sheet = workbook.getSheetAt(genericReaderSheet.getIndex());
            Iterator<Row> iteratorSheet = sheet.iterator();
            for (int i = 0; i < genericReaderSheet.getSkipHeaderRows(); i++) iteratorSheet.next();
            if (sheet.getLastRowNum() > 100000) {
                throw new RuntimeException("Il file excel caricato contiene un numero di righe superiore a 100.000");
            }
            Object obj = genericReaderSheet.readContent(workbook, sheet, iteratorSheet);
            objList.add(obj);
        }

        return objList;
    }

    public List<Object> readInputStream(InputStream inputStream) {
        List<Object> objList = new ArrayList<>();
        this.sheets.clear();
        this.initializeSheet();

        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            for (T t : this.sheets) {
                GenericReaderSheet genericReaderSheet = (GenericReaderSheet) t;
                Sheet sheet = workbook.getSheetAt(genericReaderSheet.getIndex());
                Iterator<Row> iteratorSheet = sheet.iterator();
                for (int i = 0; i < genericReaderSheet.getSkipHeaderRows(); i++) iteratorSheet.next();
                Object obj = genericReaderSheet.readContent(workbook, sheet, iteratorSheet);
                objList.add(obj);
            }

            return objList;
        } catch (IOException ioException) {
            ioException.printStackTrace();
            throw new RuntimeException("Excel non conforme");
        }
    }

    public byte[] generateBody() {
        this.sheets.clear();
        this.initializeSheet();

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(10000); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            for (T t : this.sheets) {
                GenericWriterSheet genericWriterSheet = (GenericWriterSheet) t;
                Sheet sheet = this.generateSheet(workbook, genericWriterSheet.getSheetName());
                sheet.setDefaultColumnWidth(COLUMN_WIDTH);
                this.generateHeader(workbook, sheet, genericWriterSheet.getHeader());
                genericWriterSheet.generateContent(sheet);
            }

            workbook.write(out);
            return out.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        }
    }


    public static boolean isRowEmpty(Row row, int lastCellNum) {
        if (row == null) {
            return true;
        }
        if (row.getLastCellNum() <= 0) {
            return true;
        }
        for (int cellNum = row.getFirstCellNum(); cellNum <= lastCellNum; cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != CellType.BLANK && !cell.toString().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    public static <T> T readCellValue(Row row, int cellIndex, Class<T> tClass) {
        if (row.getCell(cellIndex) == null) return null;
        if (String.class.equals(tClass)) {
            return tClass.cast(row.getCell(cellIndex).getStringCellValue().trim());
        } else if (Long.class.equals(tClass)) {
            return tClass.cast(Double.valueOf(row.getCell(cellIndex).getNumericCellValue()).longValue());
        } else if (LocalDate.class.equals(tClass)) {
            return tClass.cast(row.getCell(cellIndex).getLocalDateTimeCellValue().toLocalDate());
        } else if (LocalDateTime.class.equals(tClass)) {
            return tClass.cast(row.getCell(cellIndex).getLocalDateTimeCellValue());
        } else if (Integer.class.equals(tClass)) {
            //Necessario il try catch perché la libreria trasforma la stringa vuota in 0.0
            try {
                String stringValue = row.getCell(cellIndex).getStringCellValue();
                if (!"".equals(stringValue)) {
                    return tClass.cast(Double.valueOf(row.getCell(cellIndex).getNumericCellValue()).intValue());
                } else {
                    return null;
                }
            } catch (Exception e) {
                return tClass.cast(Double.valueOf(row.getCell(cellIndex).getNumericCellValue()).intValue());
            }
        }
        return tClass.cast(row.getCell(cellIndex).getStringCellValue());
    }

    public static void setCellValue(Row row, int cellIndex, Object value) {
        try {
            if (value == null) {
                row.createCell(cellIndex).setCellValue("");
            } else if(value instanceof Integer) {
                row.createCell(cellIndex).setCellValue((int) value);
            } else if (value instanceof Double) {
                row.createCell(cellIndex).setCellValue((double) value);
            } else if (value instanceof Float) {
                row.createCell(cellIndex).setCellValue((float) value);
            } else if (value instanceof LocalDate) {
                row.createCell(cellIndex).setCellValue(((LocalDate) value).format(localDateFormatter));
            } else if (value instanceof LocalDateTime) {
                row.createCell(cellIndex).setCellValue(((LocalDateTime) value).format(localDateTimeFormatter));
            } else {
                row.createCell(cellIndex).setCellValue(String.valueOf(value));
            }
        } catch (NullPointerException nullPointerWithEmptyString) {
            row.createCell(cellIndex).setCellValue("");
        }
    }

    /**
     * for given sheet finds row number of input string
     * @param sheet       work sheet
     * @param cellContent string that contains the term to search
     * @return row index of the first occurrence of the given string
     */
    public static int findCellRowIndex(Sheet sheet, String cellContent) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING &&
                        cell.getRichStringCellValue().getString().trim().equals(cellContent)) {

                    return row.getRowNum();

                }
            }
        }
        return 0;
    }

}
