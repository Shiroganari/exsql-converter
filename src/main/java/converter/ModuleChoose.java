package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей module_choose //
//------------------------------------------------------------------//
public class ModuleChoose {

    // Метод для добавления новых данных в таблицу module_choose
    public static void update(Connection connection, String fileName) throws IOException {
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        System.out.println();
        System.out.println("---------------------");
        System.out.println("Таблица module_choose");
        System.out.println("---------------------");

        try {
            int rowCount = 5;
            int moduleCount = 0;
            int disciplineCount = 0;

            boolean physicalCulture = false;
            String name = "";

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово Блок 2, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("Блок 2"))
                    break;

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("")) {
                    rowCount++;
                    continue;
                }

                if ((row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                    row.getCell(2).getStringCellValue().contains("элективные дисциплины")) {
                    if (row.getCell(2).getStringCellValue().contains("элективные дисциплины")) {
                        physicalCulture = true;
                    } else {
                        physicalCulture = false;
                    }

                    moduleCount = 0;
                    rowCount++;
                    name = row.getCell(2).getStringCellValue();
                    continue;
                }

                if (name.equals("")) {
                    rowCount++;
                    continue;
                }

                // Название дисциплины а также номер модуля по выбору
                String disciplineName = (row.getCell(2).getStringCellValue());
                String index = row.getCell(1).getStringCellValue();
                int fullType = index.indexOf('.');
                String firstIndex = index.substring(0, fullType);

                // Тип модуля по выбору
                ppstatement = connection.prepareStatement("SELECT id FROM type WHERE value=?;");
                ppstatement.setString(1, firstIndex);
                result = ppstatement.executeQuery();
                result.next();
                int idType = result.getInt("id");

                // Отправляем все собранные данные в таблицу module_choose
                ppstatement = connection.prepareStatement("INSERT INTO module_choose(id_type, name) VALUES(?, ?);");
                ppstatement.setInt(1, idType);
                ppstatement.setString(2, name);
                ppstatement.executeUpdate();

                moduleCount++;
                if (moduleCount / 2 == 1 && !physicalCulture) name = "";
                disciplineCount++;
                rowCount++;

                System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + disciplineName);
            }

            if (disciplineCount != 0) {
                System.out.println("---------------------");
                System.out.println("Добавление новых записей в таблицу прошло успешно!");
                System.out.println("Всего было добавлено записей: " + disciplineCount);
                System.out.println("---------------------");
            } else {
                System.out.println("Дисциплин по выбору обнаружено не было.");
                System.out.println("---------------------");
            }
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу module_choose " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("---------------------");
        }
    }
}