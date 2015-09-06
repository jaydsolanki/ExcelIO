/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.jaydsolanki.excelio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author jaysolanki
 */
public class ExcelIO {

    public enum WorkbookType {

        OFFICE_1997_2007, OFFICE_2013_AND_ABOVE
    }

    private Workbook workbook;
    private String excelFilePath;
    private int currentSheet = 0;

    public static void main(String[] args) {
        try {

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public ExcelIO(String excelFilePath) {
        this.excelFilePath = excelFilePath;
    }

    public ExcelIO(InputStream is, WorkbookType workbookType) throws IOException {
        switch (workbookType) {
            case OFFICE_1997_2007:
                workbook = new HSSFWorkbook(is);
                break;
            case OFFICE_2013_AND_ABOVE:
                workbook = new XSSFWorkbook(is);
        }
    }

    public void loadExcelFile(WorkbookType workbookType) throws FileNotFoundException, IOException, IllegalArgumentException {
        if (workbookType == null) {
            throw new IllegalArgumentException("The file does not seem to have a valid extension! please pass argument of file type in function loadExcelFile(WorkbookType workbookType)");
        }
        switch (workbookType) {
            case OFFICE_1997_2007:
                workbook = new HSSFWorkbook(new FileInputStream(new File(excelFilePath)));
                break;
            case OFFICE_2013_AND_ABOVE:
                workbook = new XSSFWorkbook(new FileInputStream(new File(excelFilePath)));
        }
    }

    public void loadExcelFile() throws FileNotFoundException, IllegalArgumentException, IOException {
        loadExcelFile(getDefaultWorkbook());
    }

    private WorkbookType getDefaultWorkbook() {
        if (excelFilePath == null) {
            return null;
        } else if (excelFilePath.endsWith("xls")) {
            return WorkbookType.OFFICE_1997_2007;
        } else {
            return WorkbookType.OFFICE_2013_AND_ABOVE;
        }
    }

    public List<List<String>> readSheet() {
        if (workbook.getSheetAt(currentSheet) == null) {
            return null;
        }
        return readSheet(workbook.getSheetAt(currentSheet));
    }

    public List<List<String>> readSheet(int sheetNo) throws IllegalArgumentException {
        if (sheetNo > workbook.getNumberOfSheets()) {
            throw new IllegalArgumentException("The sheet number " + sheetNo + " specified is out of bounds. Total Sheets: " + workbook.getNumberOfSheets());
        }
        return readSheet(workbook.getSheetAt(sheetNo));
    }

    public List<List<String>> readSheet(String sheetName) throws IllegalArgumentException {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            String sheetNames[] = new String[workbook.getNumberOfSheets()];
            for (int i = 0; i < sheetNames.length; i++) {
                sheetNames[i] = workbook.getSheetName(i);
            }
            throw new IllegalArgumentException("Sheetname not found:\"" + sheetName + "\". Available sheets: " + Arrays.toString(sheetNames));
        }
        return readSheet(sheet);
    }

    public boolean insertCell(Object Obj, int sheetNo, int rowNo, int cellNo) {
        return insertCell(Obj, workbook.getSheetAt(sheetNo), rowNo, cellNo);
    }

    public boolean insertCell(Object Obj, String sheetName, int rowNo, int cellNo) {
        return insertCell(Obj, workbook.getSheet(sheetName), rowNo, cellNo);
    }

    public boolean insertCell(Object Obj, int rowNo, int cellNo) {
        return insertCell(Obj, workbook.getSheetAt(getCurrentSheet()), rowNo, cellNo);
    }

    public boolean insertRow(List rowObj, int sheetNo, int rowNo) {
        return insertRow(rowObj, workbook.getSheetAt(sheetNo), rowNo);
    }

    public boolean insertRow(List rowObj, String sheetName, int rowNo) {
        return insertRow(rowObj, workbook.getSheet(sheetName), rowNo);
    }

    public boolean insertRow(List rowObj, int rowNo) {
        return insertRow(rowObj, workbook.getSheetAt(getCurrentSheet()), rowNo);
    }

    public boolean insertData(List<List<Object>> data, int sheetNo, int startIndex) {
        return insertData(data, workbook.getSheetAt(sheetNo), startIndex);
    }

    public boolean insertData(List<List<Object>> data, String sheetName, int startIndex) {
        return insertData(data, workbook.getSheet(sheetName), startIndex);
    }

    public boolean insertData(List<List<Object>> data, int startIndex) {
        return insertData(data, workbook.getSheetAt(getCurrentSheet()), startIndex);
    }

    /**
     * @return the currentSheet
     */
    public int getCurrentSheet() {
        return currentSheet;
    }

    /**
     * @param currentSheet the currentSheet to set
     */
    public void setCurrentSheet(int currentSheet) {
        this.currentSheet = currentSheet;
    }

    public void setCurrentSheet(String sheetName) {
        this.currentSheet = workbook.getSheetIndex(sheetName);
    }

    private List<List<String>> readSheet(Sheet sheet) {
        List<List<String>> data = new ArrayList<>();
        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            Row row = sheet.getRow(i);
            List<String> rowList = new ArrayList<String>();
            for (int j = 0; j < row.getLastCellNum(); j++) {
                rowList.add(row.getCell(j) + "");
            }
            data.add(rowList);
        }
        return data;
    }

    private boolean insertCell(Object obj, Sheet sheet, int rowNo, int cellNo) {

        if (sheet == null) {
            return false;
        }

        Row row = sheet.getRow(rowNo);
        if (row == null) {
            row = sheet.createRow(rowNo);
        }
        Cell cell = row.getCell(cellNo);
        if (cell == null) {
            cell = row.createCell(cellNo);
        }
        cell.setCellValue(obj.toString());
        return true;
    }

    public void createNewSheet() {
        workbook.createSheet();
    }

    public void createNewSheet(String sheetName) {
        workbook.createSheet(sheetName);
    }

    private boolean insertRow(List<Object> rowObjs, Sheet sheet, int rowNo) {
        if (sheet == null) {
            return false;
        }
        for (int i = 0; i < rowObjs.size(); i++) {
            insertCell(rowObjs.get(i).toString(), sheet, rowNo, i);
        }
        return true;
    }

    private boolean insertData(List<List<Object>> data, Sheet sheet, int startIndex) throws IllegalArgumentException {
        if (sheet == null) {
            return false;
        }
        if (startIndex < 0) {
            throw new IllegalArgumentException("Start Position cannot be smaller than 0. Got : " + startIndex);
        }
        for (int i = startIndex; i < startIndex + data.size(); i++) {
            List<Object> cells = data.get(i);
            for (int j = 0; j < cells.size(); j++) {
                insertCell(cells.get(j), sheet, i, j);
            }
        }
        return true;
    }

}
