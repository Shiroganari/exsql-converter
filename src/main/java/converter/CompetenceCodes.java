package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Locale;

//---------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей competence_codes //
//---------------------------------------------------------------------//
public class CompetenceCodes {

    // Метод для добавления новых данных в таблицу competence_codes
    public static void update(Connection connection, String fileName) throws IOException {
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Компетенции");
        HSSFRow row;

        System.out.println();
        System.out.println("------------------------");
        System.out.println("Таблица competence_codes");
        System.out.println("------------------------");

        try {
            int rowCount = 0;

            while ((row = sheet.getRow(rowCount)) != null) {
                String competenceCode = row.getCell(1).getStringCellValue();

                if (competenceCode.equals("") || competenceCode.contains("ПКр")) {
                    rowCount++;
                    continue;
                }

                ppstatement = connection.prepareStatement("SELECT id FROM competence_kodes WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                ppstatement.setString(1, competenceCode.toLowerCase(Locale.ROOT));
                result = ppstatement.executeQuery();
                result.next();

                // Если такая компетенция отсутствует в таблице, добавляем её
                if (result.getRow() == 0) {
                    ppstatement = connection.prepareStatement("INSERT INTO competence_kodes(name) VALUES(?);");
                    ppstatement.setString(1, competenceCode);
                    ppstatement.executeUpdate();
                }

                rowCount++;
            }

            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу " +
                "competence_codes не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("------------------------");
        }
    }
}