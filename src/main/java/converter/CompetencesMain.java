package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Locale;

//---------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей competences_main //
//---------------------------------------------------------------------//
public class CompetencesMain {

    // Метод для добавления новых данных в таблицу competences_main
    public static void update(Connection connection, String fileName) throws IOException {
        PreparedStatement ppstatement;
        Statement statement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        System.out.println();
        System.out.println("------------------------");
        System.out.println("Таблица competences_main");
        System.out.println("------------------------");

        try {
            int rowCount = 0;
            int competenceCount = 0;

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

            // Поиск id ФГОС в таблице fgos_beta
            row = sheet.getRow(30);
            String[] fullFgosName = row.getCell(19).getStringCellValue().split(" ");;
            String fgosName = fullFgosName[2];
            String fgosDate = fullFgosName[4];

            String[] fullFgosDate = fgosDate.split("\\.");
            String fgostDay = fullFgosDate[0];
            String fgostMonth = fullFgosDate[1];
            String fgostYear = fullFgosDate[2];
            fgosDate = (fgostYear) + "-" + fgostMonth + "-" + fgostDay;
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            java.util.Date utilDate = format.parse(fgosDate);
            java.sql.Date sqlDateStart = new java.sql.Date(utilDate.getTime());

            ppstatement = connection.prepareStatement("SELECT id FROM fgos_beta WHERE order_num=? AND fgos_date=? AND id_edu_lvl=?");
            ppstatement.setInt(1, Integer.parseInt(fgosName));
            ppstatement.setDate(2, sqlDateStart);
            ppstatement.setInt(3, idLevel);
            result = ppstatement.executeQuery();
            result.next();
            int idGost = result.getInt("id");

            // Определяем последнюю запись таблицы teach_plan
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id_profile FROM teach_plan ORDER BY id DESC;");
            result.next();
            int idProfile = result.getInt("id_profile");

            // Определяем необходимый id в таблице specialtys
            ppstatement = connection.prepareStatement("SELECT * FROM profile WHERE id=?;");
            ppstatement.setInt(1, idProfile);
            result = ppstatement.executeQuery();
            result.next();
            int idSpecialtys = result.getInt("id_specialtys");

            sheet = workbook.getSheet("Компетенции");

            while ((row = sheet.getRow(rowCount)) != null) {
                String competenceCode = row.getCell(1).getStringCellValue();

                if (competenceCode.equals("") || competenceCode.contains("ПКр")) {
                    rowCount++;
                    continue;
                }

                String competenceContent = row.getCell(3).getStringCellValue().trim();

                // Если такое содержание компетенции отсутствует в таблице competence_content, добавляем его
                ppstatement = connection.prepareStatement("SELECT * FROM competence_content WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                ppstatement.setString(1, competenceContent.toLowerCase(Locale.ROOT).replaceAll(" ", ""));
                result = ppstatement.executeQuery();
                result.next();

                if (result.getRow() == 0) {
                    ppstatement = connection.prepareStatement("INSERT INTO competence_content(name) VALUES(?);");
                    ppstatement.setString(1, competenceContent);
                    ppstatement.executeUpdate();

                    ppstatement = connection.prepareStatement("SELECT * FROM competence_content WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                    ppstatement.setString(1, competenceContent.toLowerCase(Locale.ROOT).replaceAll(" ", ""));
                    result = ppstatement.executeQuery();
                    result.next();
                }

                int idCompContent = result.getInt("id");

                // Поиск id кода компетенции competenceCode
                ppstatement = connection.prepareStatement("SELECT id FROM competence_kodes WHERE LOWER(name)=?;");
                ppstatement.setString(1, competenceCode.toLowerCase(Locale.ROOT));
                result = ppstatement.executeQuery();
                result.next();
                int idCompCode = result.getInt("id");

                // Поиск id типа компетенции competenceCode
                ppstatement = connection.prepareStatement("SELECT id FROM competence_types WHERE short=?;");
                ppstatement.setString(1, competenceCode.replaceAll("[^а-яА-Я]", "").trim());
                result = ppstatement.executeQuery();
                result.next();
                int idCompType = result.getInt("id");

                // Если в таблице отсутсвует компетенция competenceCode, имеющая idGost, тогда добавляем её
                ppstatement = connection.prepareStatement("SELECT * FROM competences_main WHERE id_competence_kodes=? AND id_gost=?;");
                ppstatement.setInt(1, idCompCode);
                ppstatement.setInt(2, idGost);
                result = ppstatement.executeQuery();
                result.next();

                if (result.getRow() == 0) {
                    ppstatement = connection.prepareStatement("INSERT INTO competences_main(id_competence_kodes, id_competence_types, id_competence_content, is_active, id_gost, id_specialtys, id_profile) " +
                        "VALUES(?, ?, ?, 1, ?, ?, ?);");

                    ppstatement.setInt(1, idCompCode);
                    ppstatement.setInt(2, idCompType);
                    ppstatement.setInt(3, idCompContent);
                    ppstatement.setInt(4, idGost);
                    ppstatement.setInt(5, idSpecialtys);
                    ppstatement.setInt(6, idProfile);
                    ppstatement.executeUpdate();

                    competenceCount++;

                    System.out.println(competenceCount + ". В таблицу была добавлена компетенция: " + competenceCode + " - " + qualificationName);
                }

                rowCount++;
            }

            System.out.println("------------------------");
            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("Всего было добавлено записей: " + competenceCount);
            System.out.println("------------------------");
        } catch (SQLException | ParseException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу " +
                "competences_main не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("------------------------");
        }
    }
}