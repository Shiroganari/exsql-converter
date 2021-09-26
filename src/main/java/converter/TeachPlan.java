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

//---------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей teach_plan //
//---------------------------------------------------------------//
public class TeachPlan {

    // Метод для добавления новых данных в таблицу teach_plan
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        System.out.println();
        System.out.println("------------------");
        System.out.println("Таблица teach_plan");
        System.out.println("------------------");

        try {
            /* Если таблица teach_plan пуста, присваиваем переменной id значение 1,
            иначе берём последнее значение id из таблицы и инкрементируем его */
            int id;
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;");
            result.next();
            if (result.getRow() == 0)
                id = 1;
            else
                id = result.getInt("id") + 1;

            // Год начала 4-ёх или 6-ти годичного обучения
            int fistStudyYear = Integer.parseInt(sheet.getRow(28).getCell(19).getStringCellValue());

            // Год начала и год конца определенного курса
            String academicYear = sheet.getRow(29).getCell(19).getStringCellValue();

            // Тип обучения
            String trainingForm = sheet.getRow(30).getCell(0).getStringCellValue();

            int fullType = academicYear.indexOf('-');
            String date = academicYear.substring(0, fullType);

            int yearStart = Integer.parseInt(date); // Год начала определенного курса
            int monthStart = 9;
            int dayStart = 1;
            int monthEnd = 8;
            int dayEnd = 31;

            // Столбик date_start
            String dateStart = (yearStart) + "-" + monthStart + "-" + dayStart;
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            java.util.Date utilDate = format.parse(dateStart);
            java.sql.Date sqlDateStart = new java.sql.Date(utilDate.getTime());

            // Столбик course
            int course = yearStart - fistStudyYear + 1;

            // Столбик date_end
            String dateEnd = (yearStart + 1) + "-" + monthEnd + "-" + dayEnd;
            format = new SimpleDateFormat("yyyy-MM-dd");
            utilDate = format.parse(dateEnd);
            java.sql.Date sqlDateEnd = new java.sql.Date(utilDate.getTime());

            // Столбик id_profile
            row = sheet.getRow(17);
            String fullName = row.getCell(1).getStringCellValue();
            String[] fullNameSpecialty = fullName.split("\n");
            String nameProfile;
            if (!fullName.contains("Профиль")) {
                nameProfile = fullName.replaceAll("[^a-яА-Я ]", "");
            } else {
                nameProfile = fullNameSpecialty[1].replaceAll("[^a-яА-Я ]", "").trim();
                nameProfile = nameProfile.replace("Профиль", "").trim();
            }

            ppstatement = connection.prepareStatement("SELECT id FROM profile WHERE LOWER(REPLACE(name, ' ', ''))=? ORDER BY id DESC LIMIT 1;");
            ppstatement.setString(1, nameProfile.toLowerCase(Locale.ROOT).replaceAll("[^а-яА-Я]", ""));
            result = ppstatement.executeQuery();
            result.next();
            int idProfile;
            if (result.getRow() != 0)
                idProfile = result.getInt("id");
            else
                idProfile = 0;

            // Столбик id_form
            String[] formName = trainingForm.split((" "));
            String fullFormName;

            if (trainingForm.contains("сокращенная"))
                fullFormName = formName[2] + " сокращенная";
            else if (trainingForm.contains("ускоренная"))
                fullFormName = formName[2] + " ускоренная";
            else
                fullFormName = formName[2];

            ppstatement = connection.prepareStatement("SELECT id FROM form_of_training WHERE name=?;");
            ppstatement.setString(1, fullFormName);
            result = ppstatement.executeQuery();
            result.next();
            int idForm = result.getInt("id");

            ppstatement = connection.prepareStatement("INSERT INTO " +
                "teach_plan(id, course, id_profile, id_form, date_start, date_end) VALUES(?, ?, ?, ?, ?, ?);");
            ppstatement.setInt(1, id);
            ppstatement.setInt(2, course);
            ppstatement.setInt(3, idProfile);
            ppstatement.setInt(4, idForm);
            ppstatement.setDate(5, sqlDateStart);
            ppstatement.setDate(6, sqlDateEnd);
            ppstatement.executeUpdate();

            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("В таблицу был добавлен учебный год: " + academicYear);
            System.out.println("------------------");
        } catch (SQLException | ParseException | NumberFormatException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу teach_plan " +
                "не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("------------------");
        }
    }
}