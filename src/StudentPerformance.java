import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class StudentPerformance {

    public static void main(String[] args) {
        String excelFilePath = JOptionPane.showInputDialog("Введите путь к исходному Excel-файлу(через двойной слеш \\):");

        File inputFile = new File(excelFilePath);
        String outputFilePath = inputFile.getParent() + "/output.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = firstSheet.iterator();

            int excellent = 0;
            int good = 0;
            int average = 0;
            int notAdmitted = 0;
            double totalScore = 0;
            int totalStudents = 0;

            List<String> excellentStudents = new ArrayList<>();
            List<String> goodStudents = new ArrayList<>();
            List<String> averageStudents = new ArrayList<>();
            List<String> notAdmittedStudents = new ArrayList<>();


            if (iterator.hasNext()) {
                iterator.next();
            }

            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Cell nameCell = nextRow.getCell(0);
                Cell scoreCell = nextRow.getCell(1);

                String name = nameCell.getStringCellValue();
                int score = (int) scoreCell.getNumericCellValue();

                if (score == 5) {
                    excellent++;
                    excellentStudents.add(name);
                } else if (score == 4) {
                    good++;
                    goodStudents.add(name);
                } else if (score == 3) {
                    average++;
                    averageStudents.add(name);
                } else {
                    notAdmitted++;
                    notAdmittedStudents.add(name);
                }

                totalScore += score;
                totalStudents++;
            }

            double averageScore = totalScore / totalStudents;

            Workbook outWorkbook = new XSSFWorkbook();
            Sheet outSheet = outWorkbook.createSheet("Results");

            Row row = outSheet.createRow(0);
            row.createCell(0).setCellValue("Отличники");
            row.createCell(1).setCellValue("Хорошисты");
            row.createCell(2).setCellValue("Троешники");
            row.createCell(3).setCellValue("Нет допуска");
            row.createCell(4).setCellValue("Средний балл");

            row = outSheet.createRow(1);
            row.createCell(0).setCellValue(excellent);
            row.createCell(1).setCellValue(good);
            row.createCell(2).setCellValue(average);
            row.createCell(3).setCellValue(notAdmitted);
            row.createCell(4).setCellValue(averageScore);

            int rowIndex = 3;
            for (String student : excellentStudents) {
                row = outSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(student);
            }

            rowIndex = 3;
            for (String student : goodStudents) {
                row = outSheet.getRow(rowIndex++);
                if (row == null) {
                    row = outSheet.createRow(rowIndex - 1);
                }
                row.createCell(1).setCellValue(student);
            }

            rowIndex = 3;
            for (String student : averageStudents) {
                row = outSheet.getRow(rowIndex++);
                if (row == null) {
                    row = outSheet.createRow(rowIndex - 1);
                }
                row.createCell(2).setCellValue(student);
            }

            rowIndex = 3;
            for (String student : notAdmittedStudents) {
                row = outSheet.getRow(rowIndex++);
                if (row == null) {
                    row = outSheet.createRow(rowIndex - 1);
                }
                row.createCell(3).setCellValue(student);
            }

            FileOutputStream outputStream = new FileOutputStream(outputFilePath);
            outWorkbook.write(outputStream);
            outWorkbook.close();

            workbook.close();
            inputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}