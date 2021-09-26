package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//---------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей discipline //
//---------------------------------------------------------------//
public class Disciplines {

    // Метод для добавления новых данных в таблицу discipline
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        System.out.println();
        System.out.println("------------------");
        System.out.println("Таблица discipline");
        System.out.println("------------------");

        try {
            int rowCount = 5;
            int disciplineCount = 0;

            // В переменную id добавляем значение, которое отсутствует в таблице discipline
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT MAX(ID) FROM discipline;");
            result.next();
            int disciplineId = result.getInt("max") + 1;

            while ((row = sheet.getRow(rowCount)) != null) {
                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("Дисциплины")) {
                    rowCount++;
                    continue;
                }

                // Добавляем новую запись в таблице discipline
                String disciplineName = (row.getCell(2).getStringCellValue());
                ppstatement = connection.prepareStatement("SELECT COUNT(id) AS count FROM discipline " +
                    "WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                ppstatement.setString(1, disciplineName.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();

                // Если дисциплина отсутствует в таблице - добавить новую запись
                if (result.getInt("count") == 0) {
                    ppstatement = connection.prepareStatement("INSERT INTO discipline(id, name) VALUES(?, ?);");
                    ppstatement.setInt(1, disciplineId);
                    ppstatement.setString(2, disciplineName);
                    ppstatement.executeUpdate();

                    disciplineId++;
                    disciplineCount++;

                    System.out.println(disciplineCount + ". В таблицу discipline была добавлена дисциплина: " + disciplineName);
                }

                rowCount++;
            }

            if (disciplineCount != 0) {
                System.out.println("------------------");
                System.out.println("Добавление новых записей в таблицу прошло успешно!");
                System.out.println("Всего было добавлено записей: " + disciplineCount);
                System.out.println("------------------");
            } else {
                System.out.println("Недостающих дисциплин не обнаружено!");
                System.out.println("------------------");
            }
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу discipline " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("------------------");
        }
    }
}