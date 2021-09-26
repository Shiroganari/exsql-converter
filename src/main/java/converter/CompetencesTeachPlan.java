package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Locale;

//---------------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей competences_teach_plan //
//---------------------------------------------------------------------------//
public class CompetencesTeachPlan {

    // Метод для добавления новых данных в таблицу competences_teach_plan
    public static void update(Connection connection, String fileName) throws IOException, SQLException {
        PreparedStatement ppstatement;
        Statement statement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Компетенции");
        HSSFRow row;

        System.out.println();
        System.out.println("------------------------------");
        System.out.println("Таблица competences_teach_plan");
        System.out.println("------------------------------");

        try {
            int rowCount = 0;
            int disciplineCount = 0;

            String competenceCode = "";
            String disciplineName;

            while ((row = sheet.getRow(rowCount)) != null) {
                if (!row.getCell(1).getStringCellValue().equals("") && row.getCell(0).getStringCellValue().equals("")) {
                    competenceCode = row.getCell(1).getStringCellValue();
                    rowCount++;
                    continue;
                }

                if (!row.getCell(2).getStringCellValue().equals("") && row.getCell(1).getStringCellValue().equals("")) {
                    disciplineName = row.getCell(3).getStringCellValue();

                    ppstatement = connection.prepareStatement("SELECT id FROM discipline WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                    ppstatement.setString(1, disciplineName.toLowerCase(Locale.ROOT).replaceAll("\\s+",""));
                    result = ppstatement.executeQuery();
                    result.next();
                    if (result.getRow() == 0)
                    {
                        rowCount++;
                        continue;
                    }
                    int idDiscipline = result.getInt("id");

                    ppstatement = connection.prepareStatement("SELECT id_course_disc_plan_ed FROM discipline_plan_ed WHERE id_discipline=? ORDER BY id DESC;");
                    ppstatement.setInt(1, idDiscipline);
                    result = ppstatement.executeQuery();
                    result.next();
                    if (result.getRow() == 0)
                    {
                        rowCount++;
                        continue;
                    }
                    int idCDPE = result.getInt("id_course_disc_plan_ed");

                    ppstatement = connection.prepareStatement("SELECT semester FROM course_disc_plan_ed WHERE id_course_disc_plan_ed=?;");
                    ppstatement.setInt(1, idCDPE);
                    result = ppstatement.executeQuery();
                    result.next();
                    int semester = result.getInt("semester");

                    ppstatement = connection.prepareStatement("SELECT id FROM competence_kodes WHERE name=?;");
                    ppstatement.setString(1, competenceCode);
                    result = ppstatement.executeQuery();
                    result.next();
                    int idCompCode = result.getInt("id");

                    ppstatement = connection.prepareStatement("SELECT * FROM competences_main WHERE id_competence_kodes=?;");
                    ppstatement.setInt(1, idCompCode);
                    result = ppstatement.executeQuery();
                    result.next();
                    int lastCompId = result.getInt("id");

                    statement = connection.createStatement();
                    result = statement.executeQuery("SELECT id FROM competences_matrix ORDER BY id DESC;");
                    result.next();
                    int level = result.getInt("id");

                    statement = connection.createStatement();
                    result = statement.executeQuery("SELECT * FROM teach_plan ORDER BY id DESC;");
                    result.next();
                    int idTeachPlan = result.getInt("id");

                    ppstatement = connection.prepareStatement("INSERT INTO competences_teach_plan(id_teach_plan, id_competences, id_course_disc_plan_ed, semester, " +
                        "id_competences_matrix) VALUES(?, ?, ?, ?, ?);");

                    ppstatement.setInt(1, idTeachPlan);
                    ppstatement.setInt(2, lastCompId);
                    ppstatement.setInt(3, idCDPE);
                    ppstatement.setInt(4, semester);
                    ppstatement.setInt(5, level);
                    ppstatement.executeUpdate();
                    disciplineCount++;

                    System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + disciplineName + "(Сем. " + semester + ")");

                    rowCount++;
                    continue;
                }

                rowCount++;
            }

            System.out.println("------------------------------");
            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("Всего было добавлено записей: " + disciplineCount);
            System.out.println("------------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу " +
                "competences_teach_plan не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("------------------------------");
        }
    }
}