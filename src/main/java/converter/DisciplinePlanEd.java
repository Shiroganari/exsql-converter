package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//-----------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей discipline_plan_ed //
//-----------------------------------------------------------------------//
public class DisciplinePlanEd {

    // Метод для добавления новых данных в таблицу discipline_plan_ed
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        System.out.println();
        System.out.println("--------------------------");
        System.out.println("Таблица discipline_plan_ed");
        System.out.println("--------------------------");

        try {
            // Определение последнего id в таблице teach_plan
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;");
            result.next();
            int idTeachPlan = result.getInt("id");

            int number = 0;
            int rowCount = 5;
            int idModule = 0;
            int disciplineCount = 0;
            boolean isModuleChoose = false;

            /* Если таблица discipline_plan_ed пуста, присваиваем переменной idCDPE значение 1,
            иначе берём последнее значение id_course_dis_plan_ed из таблицы discipline_plan_ed и инкрементируем его */
            int idCDPE;
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id_course_disc_plan_ed FROM discipline_plan_ed " +
                "ORDER BY id_course_disc_plan_ed DESC LIMIT 1;");
            result.next();

            if (result.getRow() == 0) {
                idCDPE = 1;
            } else {
                idCDPE = result.getInt("id_course_disc_plan_ed") + 1;
            }

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("ФТД")) {
                    break;
                }

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("специализации")) {
                    isModuleChoose = false;
                    number = 0;
                    rowCount++;
                    continue;
                }

                String index = row.getCell(1).getStringCellValue();
                int fullType = index.indexOf('.');
                String type = index.substring(0, fullType);

                String typePart;
                if (index.contains("В.")) {
                    typePart = "В";
                } else {
                    typePart = "Б";
                }

                // Определяем название дисциплины. Если в названии дисциплины находятся определённые слова, пропускаем строчку
                String discipline = (row.getCell(2).getStringCellValue());
                if (discipline.contains("Дисциплины") ||
                    discipline.contains("элективные")) {
                    rowCount++;
                    isModuleChoose = true;
                    idModule += 2;
                    number = 0;
                    continue;
                }

                if (discipline.contains("Производственная практика")) {
                    number = 0;
                    isModuleChoose = false;
                }

                // Ищем нужную дисциплину в таблице discipline и определяем её id
                ppstatement = connection.prepareStatement("SELECT id FROM discipline " +
                    "WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                ppstatement.setString(1, discipline.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();
                int idDiscipline = result.getInt("id");

                // Ищем нужный тип блока в таблице type и определяем его id
                ppstatement = connection.prepareStatement("SELECT id FROM type WHERE value=?;");
                ppstatement.setString(1,type);
                result = ppstatement.executeQuery();
                result.next();
                int idType = result.getInt("id");

                // Ищем к какой части относится дисциплина и определяем id этой части
                ppstatement = connection.prepareStatement("SELECT id FROM type_part WHERE value=?;");
                ppstatement.setString(1, typePart);
                result = ppstatement.executeQuery();
                result.next();
                int idTypePart = result.getInt("id");

                int idModuleChoose = 0;
                if (isModuleChoose) {
                    idModuleChoose = idModule;
                    number++;
                }

                // Определяем название и id кафедры
                sheet = workbook.getSheet("ПланСвод");
                row = sheet.getRow(rowCount);
                int lastColumn = row.getLastCellNum() - 1;
                String namePodraz = row.getCell(lastColumn).getStringCellValue();
                ppstatement = connection.prepareStatement("SELECT id FROM podrazdelenies WHERE LOWER(REPLACE(name, ' ', ''))=?;");
                ppstatement.setString(1, namePodraz.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();
                int idPodrazdelenie = result.getInt("id");

                // Отправляем все собранные данные в таблицу discipline_plan_ed
                sheet = workbook.getSheet("План");
                ppstatement = connection.prepareStatement("INSERT INTO discipline_plan_ed" +
                    "(id_discipline, id_course_disc_plan_ed, id_type, id_type_part, id_module_choose," +
                    "number, id_teach_plan, id_podrazdelenie) " + "VALUES(?, ?, ?, ?, ?, ?, ?, ?);");
                ppstatement.setInt(1, idDiscipline);
                ppstatement.setInt(2, idCDPE);
                ppstatement.setInt(3, idType);
                ppstatement.setInt(4, idTypePart);
                ppstatement.setInt(5, idModuleChoose);
                ppstatement.setInt(6, number);
                ppstatement.setInt(7, idTeachPlan);
                ppstatement.setInt(8, idPodrazdelenie);
                ppstatement.executeUpdate();

                disciplineCount++;
                idCDPE++;
                rowCount++;
                System.out.println(disciplineCount + (". В таблицу была добавлена дисциплина: ") + discipline);
            }

            System.out.println("--------------------------");
            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("Всего было добавлено записей: " + disciplineCount);
            System.out.println("--------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу " +
                "discipline_plan_ed не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("--------------------------");
        }
    }
}