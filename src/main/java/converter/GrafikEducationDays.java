package converter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//--------------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей grafik_education_days //
//--------------------------------------------------------------------------//
public class GrafikEducationDays {

    // Методы для добавления новых данных в таблицу grafik_education_days
    public static void updateOld(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row;
        HSSFCell cell;

        System.out.println();
        System.out.println("-----------------------------");
        System.out.println("Таблица grafik_education_days");
        System.out.println("-----------------------------");

        try {
            // Определяем последний id в таблице grafik_education
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM grafik_education ORDER BY id DESC LIMIT 1;");
            result.next();
            int idGrafikEduId = result.getInt("id");

            // Определяем нужный id_teach_plan в таблице grafik_education
            ppstatement = connection.prepareStatement("SELECT id_teach_plan FROM grafik_education WHERE id=?;");
            ppstatement.setInt(1, idGrafikEduId);
            result = ppstatement.executeQuery();
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем нужный учебный год в таблице grafik_education
            ppstatement = connection.prepareStatement("SELECT year_start FROM grafik_education WHERE id_teach_plan=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int yearStart = result.getInt("year_start");

            // Определяем номер курса в таблице teach_plan
            ppstatement = connection.prepareStatement("SELECT course FROM teach_plan WHERE id=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int course = result.getInt("course") - 1;

            int monthCount = 1; // 1 - Сентябрь | 12 - Август
            int weekCount = 1;
            int lastWeek = 53;

            while (weekCount < lastWeek) {
                ppstatement = connection.prepareStatement("INSERT INTO " +
                    "grafik_education_days(id_grafik_education, day, mounth, year, id_vid_activ) " +
                    "VALUES(?, ?, ?, ?, ?);");

                boolean newMonth = false;
                int tableRow = 12 + (7 * course);
                int firstDay;
                int lastDay;
                int whichMonth = 0;

                row = sheet.getRow(2); // 2
                String[] days;
                String weekDays;
                weekDays = row.getCell(weekCount).getStringCellValue();

                if (weekDays.equals("")) {
                    row = sheet.getRow(1);
                    weekDays = row.getCell(weekCount).getStringCellValue();
                }
                days = weekDays.split(" ");
                firstDay = Integer.parseInt(days[0]);
                lastDay = Integer.parseInt(days[days.length - 1]);

                /* Если первый день недели начинается в конце одного месяца,
                а последний день недели в начале следующего месяца */
                if (firstDay > lastDay) {
                    if (firstDay + 6 - 30 == lastDay)
                        whichMonth = 30;
                    else if (firstDay + 6 - 31 == lastDay)
                        whichMonth = 31;
                    else
                        whichMonth = 28;
                }

                boolean cellsMerged = false;
                int vidActiv = 0;

                while (tableRow < 18 + (course * 7)) {
                    row = sheet.getRow(tableRow);

                    if (firstDay > lastDay || newMonth) {
                        ppstatement.setInt(3, monthCount * 10 + (monthCount + 1));
                        newMonth = true;
                    } else {
                        ppstatement.setInt(3, monthCount);
                    }

                    ppstatement.setInt(1, idGrafikEduId);
                    ppstatement.setInt(2, firstDay);
                    ppstatement.setInt(4, yearStart);

                    String activ = row.getCell(weekCount).getStringCellValue();

                    if (!cellsMerged) {
                        switch (activ) {
                            case "":
                                vidActiv = 1;
                                break;
                            case "Э":
                                vidActiv = 2;
                                break;
                            case "У":
                                vidActiv = 3;
                                break;
                            case "П":
                                vidActiv = 4;
                                break;
                            case "Д":
                                vidActiv = 5;
                                break;
                            case "К":
                                vidActiv = 6;
                                break;
                            case "*":
                                vidActiv = 7;
                                break;
                            default:
                                break;
                        }
                    }

                    cell = row.getCell(weekCount);
                    int rowIndex = cell.getRowIndex();
                    int tempColumnIndex = cell.getColumnIndex();
                    String columnIndex = CellReference.convertNumToColString(tempColumnIndex);
                    String firstFullIndex = columnIndex + "" + rowIndex;
                    String secondFullIndex = columnIndex + "" + (rowIndex + 5);

                    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                        String firstIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getFirstColumn()) + sheet.getMergedRegion(i).getFirstRow();
                        String secondIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getLastColumn()) + sheet.getMergedRegion(i).getLastRow();

                        if (firstIndex.equals(firstFullIndex) && secondIndex.equals(secondFullIndex)) {
                            cellsMerged = true;
                        }
                    }

                    if (firstDay == whichMonth && monthCount != 3) {
                        firstDay = 0;
                    }

                    if (firstDay == whichMonth && monthCount == 3) {
                        newMonth = true;
                    }

                    ppstatement.setInt(5, vidActiv);
                    ppstatement.executeUpdate();

                    firstDay++;
                    tableRow++;
                }

                if (!newMonth && lastDay == 30 || lastDay == 31) {
                    monthCount++;
                }

                if (newMonth) {
                    monthCount++;
                }

                weekCount++;
            }

            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("-----------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу grafik_education_days " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("-----------------------------");
        }

    }
    public static void updateNew(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row;
        HSSFCell cell;

        System.out.println();
        System.out.println("-----------------------------");
        System.out.println("Таблица grafik_education_days");
        System.out.println("-----------------------------");

        try {
            // Определяем последний id в таблице grafik_education
            statement = connection.createStatement();
            result = statement.executeQuery("SELECT id FROM grafik_education ORDER BY id DESC LIMIT 1;");
            result.next();
            int idGrafikEduId = result.getInt("id");

            // Определяем нужный id_teach_plan в таблице grafik_education
            ppstatement = connection.prepareStatement("SELECT id_teach_plan FROM grafik_education WHERE id=?;");
            ppstatement.setInt(1, idGrafikEduId);
            result = ppstatement.executeQuery();
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем нужный учебный год в таблице grafik_education
            ppstatement = connection.prepareStatement("SELECT year_start FROM grafik_education WHERE id_teach_plan=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int yearStart = result.getInt("year_start");

            // Определяем номер курса в таблице teach_plan
            ppstatement = connection.prepareStatement("SELECT course FROM teach_plan WHERE id=?;");
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int course = result.getInt("course") - 1;

            int monthCount = 0; // Столбик mounth
            int weekCount = 1;
            int lastWeek = 53;
            int daysCount;

            String[] months = {"Сентябрь", "Октябрь", "Ноябрь", "Декабрь", "Январь", "Февраль",
                "Март", "Апрель", "Май", "Июнь", "Июль", "Август"};

            while (weekCount <= lastWeek) {
                row = sheet.getRow((course * 17) + 2);
                String month = row.getCell(weekCount).getStringCellValue();

                for (int i = 0; i < 12; i++) {
                    if (month.equals(months[i]))
                        monthCount = i + 1;
                }

                int firstWeekDay = 0;
                int fd = (course * 17) + 3;
                row = sheet.getRow(fd);
                while (row.getCell(weekCount).getStringCellValue().equals("")) {
                    fd++;
                    firstWeekDay++;
                    row = sheet.getRow(fd);
                }
                String firstDay = row.getCell(weekCount).getStringCellValue();

                int ld = (course * 17) + 8;
                row = sheet.getRow(ld);
                while (row.getCell(weekCount).getStringCellValue().equals("")) {
                    ld--;
                    row = sheet.getRow(ld);
                }
                String lastDay = row.getCell(weekCount).getStringCellValue();

                int firstDayDate = Integer.parseInt(firstDay);
                int lastDayDate = Integer.parseInt(lastDay);

                if (firstDayDate > lastDayDate) {
                    if (monthCount < 9)
                        monthCount = monthCount * 10 + (monthCount + 1);
                    else
                        monthCount = monthCount * 100 + (monthCount + 1);
                }

                int tableRow = 12 + firstWeekDay + (course * 17);
                int vidActiv = 0; // Столбик id_vid_activ
                boolean cellsMerged = false;

                while (tableRow < (18 + (course * 17))) {
                    ppstatement = connection.prepareStatement("INSERT INTO " +
                        "grafik_education_days(id_grafik_education, day, mounth, year, id_vid_activ) " +
                        "VALUES(?, ?, ?, ?, ?);");
                    row = sheet.getRow(tableRow  - 9);
                    daysCount = Integer.parseInt(row.getCell(weekCount).getStringCellValue());
                    row = sheet.getRow(tableRow);
                    String activ = row.getCell(weekCount).getStringCellValue();

                    if (!cellsMerged) {
                        switch (activ) {
                            case "":
                                vidActiv = 1;
                                break;
                            case "Э":
                                vidActiv = 2;
                                break;
                            case "У":
                                vidActiv = 3;
                                break;
                            case "П":
                                vidActiv = 4;
                                break;
                            case "Д":
                                vidActiv = 5;
                                break;
                            case "К":
                                vidActiv = 6;
                                break;
                            case "*":
                                vidActiv = 7;
                                break;
                            default:
                                break;
                        }
                    }

                    cell = row.getCell(weekCount);
                    int rowIndex = cell.getRowIndex();
                    int tempColumnIndex = cell.getColumnIndex();
                    String columnIndex = CellReference.convertNumToColString(tempColumnIndex);
                    String firstFullIndex = columnIndex + "" + rowIndex;
                    String secondFullIndex = columnIndex + "" + (rowIndex + 5);

                    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                        String firstIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getFirstColumn()) + sheet.getMergedRegion(i).getFirstRow();
                        String secondIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getLastColumn()) + sheet.getMergedRegion(i).getLastRow();

                        if (firstIndex.equals(firstFullIndex) && secondIndex.equals(secondFullIndex)) {
                            cellsMerged = true;
                        }
                    }

                    ppstatement.setInt(1, idGrafikEduId);
                    ppstatement.setInt(2, daysCount);
                    ppstatement.setInt(3, monthCount);
                    ppstatement.setInt(4, yearStart);
                    ppstatement.setInt(5, vidActiv);
                    ppstatement.executeUpdate();

                    if (weekCount == 53 && daysCount == lastDayDate) {
                        break;
                    }

                    tableRow++;
                }

                weekCount++;
            }

            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("-----------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу grafik_education_days " +
                "не было завершено на 100%");
            e.printStackTrace();
            System.out.println("-----------------------------");
        }
    }

    // Метод для определения типа графика (old / new)
    public  static String checkType(Connection connection, String fileName) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row = sheet.getRow(2);
        HSSFCell cell = row.getCell(0);

        if (cell.getStringCellValue().contains("Числа")) {
            return "Old table";
        } else {
            return "New table";
        }
    }

}