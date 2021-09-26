package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//-----------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей form_control //
//-----------------------------------------------------------------//
public class FormControl {

    // Метод для добавления новых данных в таблицу form_control
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        System.out.println();
        System.out.println("--------------------");
        System.out.println("Таблица form_control");
        System.out.println("--------------------");

        try {
            int rowCount = 5;
            int disciplineCount = 0;

            // Определяем значение последнего id_teach_plan в таблице discipline_plan_ed
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id_teach_plan FROM discipline_plan_ed ORDER BY id_teach_plan DESC LIMIT 1;");
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем количество видов форм контроля (экзамен, зачёт, зачёт с оценкой etc)
            row = sheet.getRow(0);
            int k = 3;
            int formControlCount = 3;
            while (formControlCount != 80) {
                if (row.getCell(k).getStringCellValue().contains("з.е.") ||
                    row.getCell(k).getStringCellValue().contains("Итого акад.часов")) {
                    break;
                }

                formControlCount++;
                k++;
            }

            // Определяем нужный id из таблицы discipline_plan_ed
            ppstatement = connection.prepareStatement("SELECT id FROM discipline_plan_ed WHERE id_teach_plan=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int idDPE = result.getInt("id");

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("ФТД")) {
                    break;
                }

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(1).getStringCellValue()).equals("")||
                    (row.getCell(2).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                    (row.getCell(2).getStringCellValue()).contains("элективные дисциплины") ||
                    (row.getCell(2).getStringCellValue()).contains("специализации")) {
                    rowCount++;
                    continue;
                }

                for (int i = 3; i < formControlCount; i++) {
                    String semester = row.getCell(i).getStringCellValue();
                    int semesterCount = 0;
                    int semesterLength = semester.length();

                    for (int j = 0; j < semesterLength; j++) {
                        if (semester.charAt(j) >= '1' && semester.charAt(j) <='9') {
                            semesterCount = Integer.parseInt(String.valueOf(semester.charAt(j)));

                            ppstatement = connection.prepareStatement("INSERT INTO " +
                                "form_control(id_discipline_plan_ed, semester, id_type_control) " +
                                "VALUES(?, ?, ?);");
                            ppstatement.setInt(1, idDPE);
                            ppstatement.setInt(2, semesterCount);
                            ppstatement.setInt(3, i - 2);
                            ppstatement.executeUpdate();
                        } else if (semester.charAt(j) >= 'A' && semester.charAt(j) <= 'C') {
                            String semesterLetter = String.valueOf(semester.charAt(j));
                            if (semesterLetter.contains("A")) semesterCount = 10;
                            else if (semesterLetter.contains("B")) semesterCount = 11;
                            else if (semesterLetter.contains("C")) semesterCount = 12;

                            ppstatement = connection.prepareStatement("INSERT INTO " +
                                "form_control(id_discipline_plan_ed, semester, id_type_control) " +
                                "VALUES(?, ?, ?);");
                            ppstatement.setInt(1, idDPE);
                            ppstatement.setInt(2, semesterCount);
                            ppstatement.setInt(3, i - 2);
                            ppstatement.executeUpdate();
                        }
                    }
                }

                idDPE++;
                disciplineCount++;
                rowCount++;

                String discipline = row.getCell(2).getStringCellValue();
                System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + discipline);
            }

            System.out.println("--------------------");
            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("Всего было добавлено записей: " + disciplineCount);
            System.out.println("--------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу form_control " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("--------------------");
        }
    }
}