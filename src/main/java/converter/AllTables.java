package converter;

import java.io.IOException;
import java.sql.*;

//-------------------------------------------------------------------------------//
// Данный класс предназначен для работы со всеми основными таблицами базы данных //
//-------------------------------------------------------------------------------//
public class AllTables {

    // Метод для добавления новых данных во все основные таблицы
    public static void updateAllTables(Connection connection, String fileName) throws IOException, SQLException {
        try {
            connection.setAutoCommit(false);

            Disciplines.update(connection, fileName); // Проверка недостающих дисциплин
            Subdivisions.update(connection, fileName); // Добавление данных в таблицу podrazdelenies
            Profile.update(connection, fileName); // Добавление данных в таблицу profile
            ModuleChoose.update(connection, fileName); // Добавление данных в таблицу module_choose
            TeachPlan.update(connection, fileName); // Добавление данных в таблицу teach_plan
            GrafikEducation.update(connection); // Добавление данных в таблицу grafik_education
            String grafikFormat = GrafikEducationDays.checkType(connection, fileName);
            if (grafikFormat.contains("Old table")) {
                GrafikEducationDays.updateOld(connection, fileName); // Добавление данных в таблицу grafik_education_days
            } else {
                GrafikEducationDays.updateNew(connection, fileName); // Добавление данных в таблицу grafik_education_days
            }
            DisciplinePlanEd.update(connection, fileName); // Добавление данных в таблицу discipline_plan_ed
            CourseDiscPlanEd.update(connection, fileName); // Добавление данных в таблицу course_disc_plan_ed
            FormControl.update(connection, fileName); // Добавление данных в таблицу form_control
            CompetenceCodes.update(connection, fileName); // Добавление данных в таблицу competence_codes
            CompetencesMain.update(connection, fileName); // Добавление данных в таблицу competences_main_temp
            CompetencesMatrix.update(connection, fileName); // Добавление данных в таблицу competences_matrix
            CompetencesTeachPlan.update(connection, fileName); // Добавление данных в таблицу competences_teach_plan

            connection.commit();
        } catch (SQLException e) {
            e.printStackTrace();
            connection.rollback();
        } finally {
            connection.setAutoCommit(true);
        }
    }

    // Метод для удаления данных из всех основных таблиц
    public static void clearAll(Connection connection) {
        Statement statement;

        try {
            statement = connection.createStatement();

            int podrazCount = statement.executeUpdate("DELETE FROM podrazdelenies; ALTER SEQUENCE podrazdelenies_id_seq RESTART WITH 1;");
            int moduleChooseCount = statement.executeUpdate("DELETE FROM module_choose; ALTER SEQUENCE module_choose_id_seq RESTART WITH 1");
            int formControlCount = statement.executeUpdate("DELETE FROM form_control; ALTER SEQUENCE form_control_id_seq RESTART WITH 1");
            int teachPlanCount = statement.executeUpdate("DELETE FROM teach_plan;");
            int grafikEduCount = statement.executeUpdate("DELETE FROM grafik_education; ALTER SEQUENCE grafik_education_id_seq RESTART WITH 1");
            int grafukEduDaysCount = statement.executeUpdate("DELETE FROM grafik_education_days; ALTER SEQUENCE grafik_education_days_id_seq RESTART WITH 1");
            int dpeCount = statement.executeUpdate("DELETE FROM discipline_plan_ed; ALTER SEQUENCE discipline_plan_ed_id_seq RESTART WITH 1");
            int cdpeCount = statement.executeUpdate("DELETE FROM course_disc_plan_ed; ALTER SEQUENCE course_disc_plan_ed_id_seq RESTART WITH 1");
            int profileCount = statement.executeUpdate("DELETE FROM profile;");
            int compContentCount = statement.executeUpdate("DELETE FROM competence_content; ALTER SEQUENCE competence_content_id_seq RESTART WITH 1");
            int compKodesCount = statement.executeUpdate("DELETE FROM competence_kodes; ALTER SEQUENCE competence_codes_id_seq RESTART WITH 1");
            int compMainCount = statement.executeUpdate("DELETE FROM competences_main; ALTER SEQUENCE competences_id_seq RESTART WITH 1");
            int compTeachPlan = statement.executeUpdate("DELETE FROM competences_teach_plan;");

            System.out.println("Количество удалённых записей в таблице podrazdelenies: " + podrazCount);
            System.out.println("Количество удалённых записей в таблице module_choose: " + moduleChooseCount);
            System.out.println("Количество удалённых записей в таблице form_control: " + formControlCount);
            System.out.println("Количество удалённых записей в таблице teach_plan: " + teachPlanCount);
            System.out.println("Количество удалённых записей в таблице grafik_education: " + grafikEduCount);
            System.out.println("Количество удалённых записей в таблице grafik_education_days: " + grafukEduDaysCount);
            System.out.println("Количество удалённых записей в таблице discipline_plan_ed: " + dpeCount);
            System.out.println("Количество удалённых записей в таблице course_disc_plan_ed: " + cdpeCount);
            System.out.println("Количество удалённых записей в таблице profile: " + profileCount);
            System.out.println("Количество удалённых записей в таблице competence_content: " + compContentCount);
            System.out.println("Количество удалённых записей в таблице competence_kodes: " + compKodesCount);
            System.out.println("Количество удалённых записей в таблице competences_main_temp: " + compMainCount);
            System.out.println("Количество удалённых записей в таблице competences_teach_plan: " + compTeachPlan);
        } catch (SQLException e) {
            System.out.println("----------------------------------------------------------------");
            System.out.println("Что-то пошло не так. Удаление записей не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("----------------------------------------------------------------");
        }
    }
}