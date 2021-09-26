package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

//-------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей podrazdelenies //
//-------------------------------------------------------------------//
public class Subdivisions {

    // Метод для добавления новых данных в таблицу podrazdelenies
    public static void update(Connection connection, String fileName) throws IOException {
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("ПланСвод");
        HSSFRow row;

        System.out.println();
        System.out.println("----------------------");
        System.out.println("Таблица podrazdelenies");
        System.out.println("----------------------");

        try {
            int rowCount = 5;
            int unitCount = 0;

            while ((row = sheet.getRow(rowCount)) != null) {
                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("Дисциплины")) {
                    rowCount++;
                    continue;
                }

                int lastCell = row.getLastCellNum() - 1; // Номер ячейки, в которой находятся нужные данные
                String unitName = row.getCell(lastCell).getStringCellValue(); // Название подразделения
                int idPodrazType = 1; // Тип подразделения | 1 - кафедра

                /* Если запись с таким именем уже присутствует в таблице,
                то пропускаем её и переходим к следующей строчке */
                ppstatement = connection.prepareStatement("SELECT id FROM podrazdelenies WHERE name=?;");
                ppstatement.setString(1, unitName);
                result = ppstatement.executeQuery();
                result.next();

                if (result.getRow() != 0)
                {
                    rowCount++;
                    continue;
                }

                // Отправляем все собранные данные в таблицу podrazdelenies
                ppstatement = connection.prepareStatement("INSERT INTO " +
                    "podrazdelenies(name, id_type_podrazdelenie) VALUES(?, ?);");
                ppstatement.setString(1, unitName);
                ppstatement.setInt(2, idPodrazType);
                ppstatement.executeUpdate();

                unitCount++;
                rowCount++;

                System.out.println(unitCount + ". В таблицу podrazdelenies был добавлен элемент: " + unitName);
            }

            if (unitCount != 0) {
                System.out.println("----------------------");
                System.out.println("Добавление новых записей в таблицу прошло успешно!");
                System.out.println("Всего было добавлено записей: " + unitCount);
                System.out.println("----------------------");
            } else {
                System.out.println("Недостающих подразделение не обнаружено!");
                System.out.println("----------------------");
            }
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу podrazdelenies " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("----------------------");
        }
    }
}