package task;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class ExcelAnalyzer {

    public static void main(String[] args) throws InvalidFormatException {
        String filePath = "C:\\Users\\starl\\Desktop\\Assignment_Timecard.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm");

            for (Row row : sheet) {
                Cell nameCell = row.getCell(getColumnIndex(sheet, "Employee Name"));
                Cell positionCell = row.getCell(getColumnIndex(sheet, "Position ID"));
                Cell timeInCell = row.getCell(getColumnIndex(sheet, "Time"));
                Cell timeOutCell = row.getCell(getColumnIndex(sheet, "Time Out"));

                String employeeName = nameCell.getStringCellValue();
                String positionID = positionCell.getStringCellValue();

                Date timeIn = null;
                Date timeOut = null;

                if (timeInCell != null && timeInCell.getCellType() == CellType.NUMERIC) {
                    timeIn = timeInCell.getDateCellValue();
                }

                if (timeOutCell != null && timeOutCell.getCellType() == CellType.NUMERIC) {
                    timeOut = timeOutCell.getDateCellValue();
                }

                // Check conditions
                if (checkConsecutiveDays(sheet, row.getRowNum(), "Employee Name", timeIn, 7, dateFormat)) {
                    System.out.println("Employee: " + employeeName + ", Position: " + positionID +
                            " has worked for 7 consecutive days.");
                }

                if (checkTimeBetweenShifts(sheet, row.getRowNum(), "Employee Name", timeIn, 10, 1, dateFormat)) {
                    System.out.println("Employee: " + employeeName + ", Position: " + positionID +
                            " has less than 10 hours between shifts but greater than 1 hour.");
                }

                if (checkSingleShiftDuration(timeIn, timeOut, 14)) {
                    System.out.println("Employee: " + employeeName + ", Position: " + positionID +
                            " has worked for more than 14 hours in a single shift.");
                }
            }

        } catch (IOException | ParseException e) {
            e.printStackTrace();
        }
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1; // Column not found
    }

    private static boolean checkConsecutiveDays(Sheet sheet, int rowIndex, String columnName, Date currentDate, int consecutiveDays, SimpleDateFormat dateFormat) throws ParseException {
        int consecutiveCount = 0;

        for (int i = rowIndex; i >= 0; i--) {
            Row currentRow = sheet.getRow(i);

            if (currentRow != null) {
                Cell currentCell = currentRow.getCell(getColumnIndex(sheet, columnName));

                if (currentCell != null && currentCell.getCellType() == CellType.NUMERIC) {
                    Date currentDateForRow = currentCell.getDateCellValue();

                    if (dateDifference(currentDate, currentDateForRow, dateFormat) == 1) {
                        consecutiveCount++;
                        if (consecutiveCount == consecutiveDays) {
                            return true;
                        }
                    } else {
                        break; // Break the loop if consecutive days are not found
                    }
                }
            }
        }
        return false;
    }

    private static boolean checkTimeBetweenShifts(Sheet sheet, int rowIndex, String columnName, Date currentDate, int maxHours, int minHours, SimpleDateFormat dateFormat) throws ParseException {
        for (int i = rowIndex - 1; i >= 0; i--) {
            Row previousRow = sheet.getRow(i);

            if (previousRow != null) {
                Cell previousCell = previousRow.getCell(getColumnIndex(sheet, columnName));

                if (previousCell != null && previousCell.getCellType() == CellType.NUMERIC) {
                    Date previousDate = previousCell.getDateCellValue();

                    long hoursBetween = hoursBetween(previousDate, currentDate);

                    if (hoursBetween < maxHours && hoursBetween > minHours) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private static boolean checkSingleShiftDuration(Date timeIn, Date timeOut, int maxHours) {
        if (timeIn != null && timeOut != null) {
            long duration = timeOut.getTime() - timeIn.getTime();
            long hours = duration / (60 * 60 * 1000);
            return hours > maxHours;
        }
        return false;
    }

    private static long dateDifference(Date date1, Date date2, SimpleDateFormat dateFormat) throws ParseException {
        long diff = date1.getTime() - date2.getTime();
        return TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
    }

    private static long hoursBetween(Date date1, Date date2) {
        return Math.abs(date1.getTime() - date2.getTime()) / (60 * 60 * 1000);
    }
}
