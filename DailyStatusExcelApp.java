package com.kunal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.*;
import java.time.format.TextStyle;
import java.util.*;

public class DailyStatusExcelApp {

    static final String BASE_PATH = "D:\\Status"; // change this path

    public static void main(String[] args) {
        try {
            Scanner sc = new Scanner(System.in);

            System.out.print("Enter username: ");
            String user = sc.nextLine().trim();

            File userDir = new File(BASE_PATH + user);
            if (!userDir.exists()) userDir.mkdirs();

            YearMonth ym = YearMonth.now();
            String fileName = ym + ".xlsx";
            File excelFile = new File(userDir, fileName);

            Workbook wb;
            Sheet sheet;

            if (!excelFile.exists()) {
                // CREATE NEW FILE
                wb = new XSSFWorkbook();
                sheet = wb.createSheet("Status");
                createHeader(wb, sheet);
                createMonthTemplate(wb, sheet, ym);
            } else {
                // OPEN EXISTING FILE
                FileInputStream fis = new FileInputStream(excelFile);
                wb = new XSSFWorkbook(fis);
                sheet = wb.getSheetAt(0);
                fis.close();
            }

            // DAILY UPDATE
            updateToday(sheet, wb, sc);

            // SAVE
            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }
            wb.close();

            System.out.println("Daily status updated successfully.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void createHeader(Workbook wb, Sheet sheet) {
        CellStyle headerStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        Row header = sheet.createRow(0);
        String[] cols = {"Date", "Day", "Login", "Logout", "Task"};

        for (int i = 0; i < cols.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(cols[i]);
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(i, 5000);
        }
    }

    static void createMonthTemplate(Workbook wb, Sheet sheet, YearMonth ym) {

        CellStyle weekendStyle = wb.createCellStyle();
        weekendStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        weekendStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle holidayStyle = wb.createCellStyle();
        holidayStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        holidayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // HOLIDAYS (CUSTOMIZE)
        Set<LocalDate> holidays = new HashSet<>();
        holidays.add(LocalDate.of(2026, 1, 26)); // example
        holidays.add(LocalDate.of(2026, 8, 15)); // example

        LocalDate start = ym.atDay(1);
        int days = ym.lengthOfMonth();

        for (int i = 0; i < days; i++) {
            LocalDate date = start.plusDays(i);
            Row row = sheet.createRow(i + 1);

            row.createCell(0).setCellValue(date.toString());
            row.createCell(1).setCellValue(
                    date.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH)
            );
            row.createCell(2).setCellValue("");
            row.createCell(3).setCellValue("");
            row.createCell(4).setCellValue("");

            boolean isWeekend = date.getDayOfWeek() == DayOfWeek.SATURDAY ||
                    date.getDayOfWeek() == DayOfWeek.SUNDAY;

            boolean isHoliday = holidays.contains(date);

            if (isWeekend || isHoliday) {
                CellStyle style = isHoliday ? holidayStyle : weekendStyle;
                for (int c = 0; c < 5; c++) {
                    row.getCell(c).setCellStyle(style);
                }
            }
        }
    }

    static void updateToday(Sheet sheet, Workbook wb, Scanner sc) {

        LocalDate today = LocalDate.now();
        String todayStr = today.toString();

        System.out.print("Login time (HH:mm): ");
        String login = sc.nextLine();

        System.out.print("Logout time (HH:mm): ");
        String logout = sc.nextLine();

        System.out.print("Task: ");
        String task = sc.nextLine();

        int rows = sheet.getLastRowNum();

        for (int i = 1; i <= rows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell dateCell = row.getCell(0);
            if (dateCell == null) continue;

            if (todayStr.equals(dateCell.getStringCellValue())) {
                row.getCell(2).setCellValue(login);
                row.getCell(3).setCellValue(logout);
                row.getCell(4).setCellValue(task);
                break;
            }
        }
    }
}

