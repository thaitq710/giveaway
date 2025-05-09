package vn.goldenboy.giveawayserver.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import vn.goldenboy.giveawayserver.exception.BusinessException;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

    public List<Map<String, Object>> processExcelFile(MultipartFile file) {
        if (file.isEmpty()) {
            throw new BusinessException("FILE_EMPTY", "File is empty");
        }

        // Map to store number frequency and row information
        Map<Integer, List<Integer>> numberRowsMap = new HashMap<>();

        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {

            // Iterate through all sheets
            for (Sheet sheet : workbook) {
                for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null)
                        continue;

                    // Iterate through cells in the row
                    for (int cellIndex = 0; cellIndex < row.getPhysicalNumberOfCells(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        int num = extractValidNumber(cell);
                        if (num >= 1 && num <= 1000) {
                            // If number is valid, store the row where it appears
                            numberRowsMap
                                    .computeIfAbsent(num, k -> new ArrayList<>())
                                    .add(rowIndex + 1); // rowIndex + 1 to make row numbers start from 1
                        }
                    }
                }
            }

            // Find minimum frequency
            int minFreq = numberRowsMap.values().stream()
                    .mapToInt(List::size)
                    .min()
                    .orElse(0);

            // Return result: number and its row appearances
            List<Map<String, Object>> result = new ArrayList<>();
            for (Map.Entry<Integer, List<Integer>> entry : numberRowsMap.entrySet()) {
                if (entry.getValue().size() == minFreq) {
                    Map<String, Object> numberInfo = new HashMap<>();
                    numberInfo.put("number", entry.getKey());
                    numberInfo.put("count", entry.getValue().size());
                    numberInfo.put("rows", entry.getValue()); // Rows where this number appears
                    result.add(numberInfo);
                }
            }

            return result;
        } catch (BusinessException e) {
            throw e;
        } catch (Exception e) {
            throw new BusinessException("FILE_PROCESSING_ERROR", "Error processing Excel file: " + e.getMessage());
        }
    }

    private int extractValidNumber(Cell cell) {
        try {
            if (cell == null)
                return -1;

            return switch (cell.getCellType()) {
                case NUMERIC -> (int) cell.getNumericCellValue();
                case STRING -> Integer.parseInt(cell.getStringCellValue().trim());
                default -> -1;
            };
        } catch (Exception e) {
            return -1;
        }
    }
}