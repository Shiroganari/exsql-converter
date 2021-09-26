package converter;

import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

//---------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей grafik_education //
//---------------------------------------------------------------------//
public class GrafikEducation {

    // Метод для добавления новых данных в таблицу grafik_education
    public static void update(Connection connection) {
        PreparedStatement ppstatement;
        Statement statement;
        ResultSet result;

        System.out.println();
        System.out.println("------------------------");
        System.out.println("Таблица grafik_education");
        System.out.println("------------------------");

        try {
            // Определяем последний id в таблице teach_plan
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;");
            result.next();
            int idTeachPlan = result.getInt("id");

            // Определяем год начала курса в таблице teach_plan
            ppstatement = connection.prepareStatement("SELECT date_start FROM teach_plan WHERE id=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();

            while (result.next()) {
                String pattern = "yyyy-MM-dd";
                DateFormat df = new SimpleDateFormat(pattern);
                Date year = result.getDate("date_start");
                String yearStart = df.format(year);
                int fullType = yearStart.indexOf('-');
                String date = yearStart.substring(0, fullType);
                int dateStart = Integer.parseInt(date);

                ppstatement = connection.prepareStatement("INSERT INTO grafik_education(id_teach_plan, year_start) VALUES(?, ?);");
                ppstatement.setInt(1, idTeachPlan);
                ppstatement.setInt(2, dateStart);
                ppstatement.executeUpdate();

                System.out.println("Добавление новых записей в таблицу прошло успешно!");
                System.out.println("В таблицу был добавлен год начала курса: " + dateStart);
                System.out.println("------------------------");
            }
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу grafik_education " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("------------------------");
        }
    }
}