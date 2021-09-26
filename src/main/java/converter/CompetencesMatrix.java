package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Locale;

//-----------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей competences_matrix //
//-----------------------------------------------------------------------//
public class CompetencesMatrix {

    // Метод для добавления новых данных в таблицу competences_matrix
    public static void update(Connection connection, String fileName) throws IOException, SQLException {
        PreparedStatement ppstatement;
        Statement statement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        System.out.println();
        System.out.println("--------------------------");
        System.out.println("Таблица competences_matrix");
        System.out.println("--------------------------");

        try {
            row = sheet.getRow(28);
            String[] qualification = row.getCell(0).getStringCellValue().split(" ");
            String qualificationName = qualification[1].toLowerCase(Locale.ROOT);

            if (qualificationName.contains("бакалавр")) qualificationName = "Бакалавриат";
            else if (qualificationName.contains("магистр")) qualificationName = "Магистратура";
            else if (qualificationName.contains("специалист")) qualificationName = "Специалитет";

            // Поиск id уровня образования в таблице level
            ppstatement = connection.prepareStatement("SELECT id FROM level WHERE name=?;");
            ppstatement.setString(1, qualificationName);
            result = ppstatement.executeQuery();
            result.next();
            int idLevel = result.getInt("id");

            // Поиск id статуса компетенции "Ожидает проверки" в таблицу competence_status
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM competence_status WHERE name='Ожидает проверки';");
            result.next();
            int idStatus = result.getInt("id");

            ppstatement = connection.prepareStatement("INSERT INTO competences_matrix(id_level, id_competence_status) VALUES(?, ?);");
            ppstatement.setInt(1, idLevel);
            ppstatement.setInt(2, idStatus);
            ppstatement.executeUpdate();

            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("--------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу " +
                "competences_matrix не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("--------------------------");
        }
    }
}