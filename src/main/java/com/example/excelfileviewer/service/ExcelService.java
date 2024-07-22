package com.example.excelfileviewer.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

@Service
public class ExcelService {

    private Map<String, List<Map<String, Object>>> excelData = new HashMap<>();

    @PostConstruct
    public void init() throws IOException {
        String folderPath = "excel"; // 更新为您的文件夹路径
        loadExcelFiles(folderPath);
    }

    public void loadExcelFiles(String folderPath) throws IOException {
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls"));

        if (files != null) {
            for (File file : files) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {
                    for (Sheet sheet : workbook) {
                        List<Map<String, Object>> sheetData = new ArrayList<>();
                        int rows = sheet.getPhysicalNumberOfRows();
                        Row headerRow = sheet.getRow(0);
                        if (headerRow == null) {
                            continue; // Skip sheets without a header row
                        }
                        for (int i = 1; i < rows; i++) {
                            Row row = sheet.getRow(i);
                            if (row != null) {
                                Map<String, Object> rowData = new HashMap<>();
                                for (Cell cell : row) {
                                    Cell headerCell = headerRow.getCell(cell.getColumnIndex());
                                    String header = headerCell != null ? headerCell.toString() : "Column " + cell.getColumnIndex();
                                    rowData.put(header, cell.toString());
                                }
                                sheetData.add(rowData);
                            }
                        }
                        excelData.put(file.getName() + " - " + sheet.getSheetName(), sheetData);
                    }
                }
            }
        }
    }

    public List<Map<String, Object>> search(String keyword) {
        List<Map<String, Object>> results = new ArrayList<>();
        Pattern pattern = Pattern.compile(Pattern.quote(keyword), Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);

        for (Map.Entry<String, List<Map<String, Object>>> entry : excelData.entrySet()) {
            for (Map<String, Object> row : entry.getValue()) {
                boolean matchFound = false;
                Map<String, Object> result = new HashMap<>();
                for (Map.Entry<String, Object> cell : row.entrySet()) {
                    String value = cell.getValue().toString();
                    if (pattern.matcher(value).find()) {
                        matchFound = true;
                        result.put("matchedCellHeader", cell.getKey());
                        result.put("matchedCell", value);

                        // Get right cell
                        Object rightCell = getAdjacentCell(entry.getValue(), row, cell.getKey(), "right");
                        result.put("rightCell", rightCell != null ? rightCell.toString() : "none");

                        // Get bottom cell
                        Object bottomCell = getAdjacentCell(entry.getValue(), row, cell.getKey(), "bottom");
                        result.put("bottomCell", bottomCell != null ? bottomCell.toString() : "none");
                        break;
                    }
                }
                if (matchFound) {
                    result.put("source", entry.getKey());
                    results.add(result);
                }
            }
        }
        return results;
    }

    private Object getAdjacentCell(List<Map<String, Object>> sheetData, Map<String, Object> rowData, String header, String direction) {
        int rowIndex = sheetData.indexOf(rowData);
        int columnIndex = new ArrayList<>(rowData.keySet()).indexOf(header);

        if (direction.equals("right")) {
            if (columnIndex + 1 < rowData.size()) {
                String nextHeader = new ArrayList<>(rowData.keySet()).get(columnIndex + 1);
                return rowData.get(nextHeader);
            }
        } else if (direction.equals("bottom")) {
            if (rowIndex + 1 < sheetData.size()) {
                Map<String, Object> nextRowData = sheetData.get(rowIndex + 1);
                return nextRowData.get(header);
            }
        }
        return null;
    }
}

