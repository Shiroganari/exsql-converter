package converter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

//------------------------------------------------------------------------//
// Данный класс содержит методы для работы с таблицей course_disc_plan_ed //
//------------------------------------------------------------------------//
public class CourseDiscPlanEd {

    // Метод для добавления новых данных в таблицу course_disc_plan_ed
    public static void update(Connection connection, String fileName) throws IOException {
        Statement statement;
        PreparedStatement ppstatement;
        ResultSet result;

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        System.out.println();
        System.out.println("---------------------------");
        System.out.println("Таблица course_disk_plan_ed");
        System.out.println("---------------------------");

        try {
            int currentSemester = 1;
            int semesterCount = 0;
            int disciplineCount = 0;
            String disciplineName;

            // Определение количества семестров
            String term = "";
            String studyForm = "";

            row = sheet.getRow(30);
            studyForm = row.getCell(0).getStringCellValue();
            row = sheet.getRow(31);
            term = row.getCell(0).getStringCellValue();

            if (term.contains("2г")) {
                semesterCount = 4;
            }
            else if (term.contains("4г")) {
                semesterCount = 8;
            }
            else if (term.contains("5л")) {
                semesterCount = 10;
            }
            else if (term.contains("2г 6м")) {
                semesterCount = 6;
            }
            else if (term.contains("4г 6м")) {
                semesterCount = 10;
            }
            else if (term.contains("5л 6м")) {
                semesterCount = 12;
            }
            else if (term.contains("3г 6м")) {
                semesterCount = 7;
            }

            /* Расчёт количества столбцов для одного семестра,
            а также получение номера столбца, на котором начинается 1-ый семестр */
            sheet = workbook.getSheet("План");
            row = sheet.getRow(0);
            int cellCourseOne = 0;
            int cellCourseTwo = cellCourseOne;

            while (!row.getCell(cellCourseOne).getStringCellValue().contains("Курс 1")) {
                cellCourseOne++;
            }

            while (!row.getCell(cellCourseTwo).getStringCellValue().contains("Курс 2")) {
                cellCourseTwo++;
            }

            int totalCell;
            if (studyForm.contains("Заочная") && !studyForm.contains("Очно-заочная")) {
                totalCell = cellCourseTwo - cellCourseOne;
                semesterCount /= 2;
            } else {
                totalCell =  (cellCourseTwo - cellCourseOne) / 2;
            }

            while (currentSemester <= semesterCount) {
                int rowCount = 5;

                // Определение id_course_dis_plan_ed в таблице discipline_plan_ed
                statement = connection.createStatement();
                result = statement.executeQuery("SELECT * FROM discipline_plan_ed ORDER BY id_teach_plan DESC LIMIT 1;");
                result.next();
                int idCDPE = result.getInt("id_course_disc_plan_ed");

                while ((row = sheet.getRow(rowCount)) != null) {
                    // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                    if ((row.getCell(0).getStringCellValue()).contains("ФТД")) {
                        break;
                    }

                    // Пропускаем ненужные строки и переходим к следующим
                    if ((row.getCell(0).getStringCellValue()).equals("") ||
                        (row.getCell(1).getStringCellValue()).equals("") ||
                        (row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                        (row.getCell(2).getStringCellValue()).contains("элективные дисциплины") ||
                        (row.getCell(2).getStringCellValue()).contains("специализации")) {
                        rowCount++;
                        continue;
                    }

                    int idZachEd = 1;
                    int zachEd = 0;
                    int lec = 0;
                    int lab = 0;
                    int prac = 0;
                    int sr = 0;
                    int control = 0;

                    if (totalCell == 9) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 9)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 9)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 9)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 9)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 9)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 7 + ((currentSemester - 1) * 9)).getStringCellValue());
                    }
                    else if (totalCell == 8) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 8)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 8)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 8)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 8)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 8)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 8)).getStringCellValue());
                    }
                    else if (totalCell == 7) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 7)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 7)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 7)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 7)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 7)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 7)).getStringCellValue());
                    } else if (totalCell == 6) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 6)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 1 + ((currentSemester - 1) * 6)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 6)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 6)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 6)).getStringCellValue());

                        HSSFRow tempRow = sheet.getRow(2);
                        if (!tempRow.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 6)).getStringCellValue().contains("Формы"))
                            control = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 6)).getStringCellValue());
                        else
                            control = 0;
                    } else if (totalCell == 5) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 5)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 1 + ((currentSemester - 1) * 5)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 5)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 5)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 5)).getStringCellValue());
                        control = 0;
                    } else if (totalCell == 4) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 4)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 1 + ((currentSemester - 1) * 4)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 4)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 4)).getStringCellValue());
                    }

                    if (zachEd == 0 && lec == 0 && lab == 0 && prac == 0 && sr == 0 && control == 0) {
                        rowCount++;
                        idCDPE++;
                        continue;
                    }

                    // Отправляем все собранные данные в таблицу course_disc_plan_ed
                    ppstatement = connection.prepareStatement("INSERT INTO course_disc_plan_ed"
                        + "(id_zach_ed, zach_ed, semester, lec, lab, prac, sr, control, id_course_disc_plan_ed) "
                        + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?);");

                    ppstatement.setInt(1, idZachEd);
                    ppstatement.setInt(2, zachEd);
                    ppstatement.setInt(3, currentSemester);
                    ppstatement.setInt(4, lec);
                    ppstatement.setInt(5, lab);
                    ppstatement.setInt(6, prac);
                    ppstatement.setInt(7, sr);
                    ppstatement.setInt(8, control);
                    ppstatement.setInt(9, idCDPE);
                    ppstatement.executeUpdate();

                    disciplineName = row.getCell(2).getStringCellValue();
                    disciplineCount++;

                    System.out.println(disciplineCount + ". Семестр " + currentSemester + ". В таблицу была добавлена дисциплина: " + disciplineName);

                    rowCount++;
                    idCDPE++;
                }

                currentSemester++;
            }

            System.out.println("---------------------------");
            System.out.println("Добавление новых записей в таблицу прошло успешно!");
            System.out.println("---------------------------");
        } catch (SQLException e) {
            System.out.println("Что-то пошло не так. Добавление новых записей в таблицу  " +
                "course_disk_plan_ed не было завершено на 100%.");
            e.printStackTrace();
            System.out.println("---------------------------");
        }
    }
}