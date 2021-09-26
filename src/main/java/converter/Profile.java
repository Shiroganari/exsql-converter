package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Locale;

//------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей profile //
//------------------------------------------------------------//
public class Profile {

    // Метод для добавления новых данных в таблицу profile
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        System.out.println();
        System.out.println("----------------");
        System.out.println("Таблица profiles");
        System.out.println("----------------");

        try {
            int rowCount = 15;
            int cellCount = 1;

            // Определение кода специальности
            row = sheet.getRow(rowCount);
            String codeSpecialty = row.getCell(cellCount).getStringCellValue();

            // Определение названия специальности
            row = sheet.getRow(rowCount + 2);
            String fullName = row.getCell(cellCount).getStringCellValue();
            String[] fullNameSpecialty = fullName.split("\n");
            String nameSpecialtys = fullNameSpecialty[0].replaceAll("[^a-яА-Я ]", "").trim();
            String nameProfile;

            if (!fullName.contains("Профиль")) {
                nameProfile = fullName.replaceAll("[^a-яА-Я ]", "");
            } else {
                nameProfile = fullNameSpecialty[1].replaceAll("[^a-яА-Я ]", "");
                nameProfile = nameProfile.replace("Профиль", "").trim();
            }

            // Определяем, присутствует ли в таблице specialtys такая специальность
            int idSpecialtys;
            ppstatement = connection.prepareStatement("SELECT * FROM specialtys WHERE REPLACE(code, '.', '')=?;");
            ppstatement.setString(1, codeSpecialty.replaceAll("[^0-9]", ""));
            result = ppstatement.executeQuery();
            result.next();

            /* Если такая специальность отсутствует, добавляем её в таблицу и извлекаем её id,
            иначе извлекаем id уже существующей записи */
            if (result.getRow() == 0) {
                ppstatement = connection.prepareStatement("INSERT INTO specialtys(name, code) VALUES(?, ?);");
                ppstatement.setString(1, nameSpecialtys);
                ppstatement.setString(2, codeSpecialty);
                ppstatement.executeUpdate();

                ppstatement = connection.prepareStatement("SELECT * FROM specialtys WHERE REPLACE(code, '.', '')=?;");
                ppstatement.setString(1, codeSpecialty.replaceAll("[^0-9]", ""));
                result = ppstatement.executeQuery();
                result.next();
                idSpecialtys = result.getInt("id");
            } else {
                idSpecialtys = result.getInt("id");
            }

            // Если такой профиль уже существует в таблице, его повторного добавления в таблицу не происходит
            ppstatement = connection.prepareStatement("SELECT id FROM profile WHERE LOWER(REPLACE(name, ' ', ''))=?;");
            ppstatement.setString(1, nameProfile.toLowerCase(Locale.ROOT).replaceAll("[^а-яА-Я]", ""));
            result = ppstatement.executeQuery();
            result.next();

            if (result.getRow() != 0) {
                System.out.println("Новых записей в таблицу добавлено не было.");
                System.out.println("----------------");
                return;
            }

            // Определяем id таблицы profile
            int idProfile;
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM profile ORDER BY id DESC LIMIT 1;");
            result.next();

            if (result.getRow() == 0) {
                idProfile = 1;
            } else {
                idProfile = result.getInt("id") + 1;
            }

            // Отправляем все собранные данные в таблицу profile
            ppstatement = connection.prepareStatement("INSERT INTO " +
                "profile(id, code, specialty, profile, id_specialtys, name) VALUES(?, ?, ?, ?, ?, ?);");
            ppstatement.setInt(1, idProfile);
            ppstatement.setString(2, codeSpecialty);
            ppstatement.setString(3, nameSpecialtys);
            ppstatement.setString(4, nameProfile);
            ppstatement.setInt(5, idSpecialtys);
            ppstatement.setString(6, nameProfile);
            ppstatement.executeUpdate();

            System.out.println("В таблицу был добавлен профиль: " + nameProfile);
            System.out.println("----------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу profile " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("----------------");
        }
    }
}