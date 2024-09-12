//Задание: разработать программу, рассчитывающую график платежей по кредиту.
// Входные данные: сумма кредита, срок, процентная ставка, дата выдачи (1-31), тип графика (аннуитетный или дифференцированный). 
// Выходные данные: таблица, содержащая даты и суммы платежей.
// Условия: платеж по кредиту осуществляется ежемесячно,
// при попадании на воскресенье или праздники платёж переносится на первый рабочий день.
// Что ожидаем в качестве решения: код.

package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.Scanner;

public class CreditCalculator {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Введите сумму кредита: ");
        double creditSum = scanner.nextDouble();
        System.out.println("Введите срок кредита в месяцах: ");
        double term = scanner.nextDouble();
        System.out.println("Введите процентную ставку кредита: ");
        double interestRate = scanner.nextDouble();
        System.out.println("Введите дату выдачи кредита (1-31): ");
        double issueDate = scanner.nextDouble();
        System.out.println("График платежей аннуитентный? Введите 'true' или 'false': ");
        boolean graphTypeAnn = scanner.nextBoolean();
        int scale = 2;
        double r = (interestRate / 100) / 12;

        LocalDate newYear = LocalDate.of(0, 1, 1);
        LocalDate victoryDay = LocalDate.of(0, 5, 9);
        LocalDate cristmasDay = LocalDate.of(0, 1, 7);
        LocalDate localDate = LocalDate.now();
        LocalDate targetDate = LocalDate.of(localDate.getYear(), localDate.getMonth(), (int) issueDate);
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Задание 1");
        Row row = sheet.createRow(0);
        Cell cell0 = row.createCell(0);
        cell0.setCellValue("Дата");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue("Сумма платежа");

        if (graphTypeAnn) {
            calculateAnn(r, creditSum, term, scale, targetDate, newYear, victoryDay, cristmasDay, sheet);
        } else {
            calculateDiff(term, targetDate, newYear, victoryDay, cristmasDay, creditSum, term, interestRate, scale, sheet);
        }
        try (FileOutputStream outputStream = new FileOutputStream("Task1.xlsx")) {
            workbook.write(outputStream);
            System.out.println("Data saved to Excel file successfully.");
        } catch (IOException e) {
            System.out.println("Error saving data to Excel file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing the workbook: " + e.getMessage());
            }
        }
    }

    private static void calculateAnn(double r, double creditSum, double term, int scale, LocalDate targetDate, LocalDate newYear, LocalDate victoryDay, LocalDate cristmasDay, Sheet sheet) {
        double result;
        double roundResult;
        result = (r * creditSum) / (1 - Math.pow(1 + r, -term));
        BigDecimal bigDecimal = new BigDecimal(result);
        bigDecimal = bigDecimal.setScale(scale, RoundingMode.HALF_UP);
        roundResult = bigDecimal.doubleValue();
        for (int i = 0; i < term; i++) {
            LocalDate nextDate = targetDate.plusMonths(i);
            LocalDate holidayCheckDate = LocalDate.of(0, nextDate.getMonthValue(), nextDate.getDayOfMonth());
            if (nextDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                nextDate = nextDate.plusDays(1);
            }
            if (holidayCheckDate.equals(newYear) || holidayCheckDate.equals(victoryDay) || holidayCheckDate.equals(cristmasDay)) {
                nextDate = nextDate.plusDays(1);
            }
            Row row = sheet.createRow(i + 1);
            Cell cellDate = row.createCell(0);
            cellDate.setCellValue(nextDate.toString());
            Cell cellResult = row.createCell(1);
            cellResult.setCellValue(roundResult);
        }
    }

    private static void calculateDiff(double term, LocalDate targetDate, LocalDate newYear, LocalDate victoryDay, LocalDate cristmasDay, double creditSum, double temp, double interestRate, int scale, Sheet sheet) {
        double roundResult;
        double result;
        for (int i = 0; i < term; i++) {
            LocalDate nextDate = targetDate.plusMonths(i);
            LocalDate holidayCheckDate = LocalDate.of(0, nextDate.getMonthValue(), nextDate.getDayOfMonth());
            if (nextDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                nextDate = nextDate.plusDays(1);
            }
            if (holidayCheckDate.equals(newYear) || holidayCheckDate.equals(victoryDay) || holidayCheckDate.equals(cristmasDay)) {
                nextDate = nextDate.plusDays(1);
            }
            boolean isLeapYear = nextDate.getYear() % 4 == 0;
            int daysYear;
            if (isLeapYear) {
                daysYear = 366;
            } else {
                daysYear = 365;
            }
            result = creditSum / temp + ((creditSum * (interestRate / 100) * nextDate.getMonthValue()) / daysYear);
            BigDecimal bigDecimal = new BigDecimal(result);
            bigDecimal = bigDecimal.setScale(scale, RoundingMode.HALF_UP);
            roundResult = bigDecimal.doubleValue();
            Row row = sheet.createRow(i + 1);
            Cell cellDate = row.createCell(0);
            cellDate.setCellValue(nextDate.toString());
            Cell cellResult = row.createCell(1);
            cellResult.setCellValue(roundResult);
            temp = temp - 1;
            creditSum = creditSum - result;
        }
    }
}
