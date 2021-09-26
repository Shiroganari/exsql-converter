import converter.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

public class Main {
    public static void main(String[] args) throws IOException {
        Connection connection;

        //---------------------------//
        // Подключение к базе данных //
        //---------------------------//
        Properties props = new Properties();
        try (InputStream in = Files.newInputStream(Paths.get("database.properties"))) {
            props.load(in);
        } catch (NoSuchFileException e) {
            System.out.println("Отсутствует файл database.properties");
            System.out.println();
            e.printStackTrace();
            System.out.println("----------------------------------------");
        }

        String DATABASE = props.getProperty("DATABASE");
        String PASS = props.getProperty("PASSWORD");
        String USER = props.getProperty("LOGIN");
        String PORT = props.getProperty("PORT");
        String IP = props.getProperty("IP");

        String DB_URL = "jdbc:postgresql://" + IP + ":" + PORT + "/" + DATABASE;

        System.out.println("Проверка подключения к PostgreSQL JDBC...");

        try {
            Class.forName("org.postgresql.Driver");
        } catch (ClassNotFoundException e) {
            System.out.println("PostgreSQL JDBC Driver не найден.");
            e.printStackTrace();
            return;
        }

        System.out.println("PostgreSQL JDBC Driver был успешно подключен!");

        try {
            connection = DriverManager.getConnection(DB_URL, USER, PASS);
        } catch (SQLException e) {
            System.out.println("Ошибка подключения.");
            e.printStackTrace();
            return;
        }

        if (connection != null) {
            System.out.println("Вы успешно подключились к базе данных!");
        } else {
            System.out.println("Не удалось подключиться к базе данных.");
        }

        //--------------------------------//
        // Работа с данными в базе данных //
        //--------------------------------//
        try {
            String actionType = args[0]; // Тип действия, который хочет выполнить пользователь

            switch (actionType) {
                //---------------------------------------//
                // Очистка основных таблиц в базе данных //
                //---------------------------------------//
                case "clear":
                    AllTables.clearAll(connection);
                    break;

                //---------------------------------------------------//
                // Добавление данных одного .xls файла в базу данных //
                //---------------------------------------------------//
                case "update": {
                    String fileName = args[1]; // Название .xls файла

                    //--------------------------------//
                    // Проверка недостающих дисциплин //
                    //--------------------------------//
                    Disciplines.update(connection, fileName);

                    //------------------------//
                    // Таблица podrazdelenies //
                    //------------------------//
                    Subdivisions.update(connection, fileName);

                    //-----------------//
                    // Таблица profile //
                    //-----------------//
                    Profile.update(connection, fileName);

                    //-----------------------//
                    // Таблица module_choose //
                    //-----------------------//
                    ModuleChoose.update(connection, fileName);

                    //--------------------//
                    // Таблица teach_plan //
                    //--------------------//
                    TeachPlan.update(connection, fileName);

                    //--------------------------//
                    // Таблица grafik_education //
                    //--------------------------//
                    GrafikEducation.update(connection);

                    //-------------------------------//
                    // Таблица grafik_education_days //
                    //-------------------------------//
                    String grafikFormat = GrafikEducationDays.checkType(connection, fileName);
                    if (grafikFormat.contains("Old table")) {
                        GrafikEducationDays.updateOld(connection, fileName);
                    } else {
                        GrafikEducationDays.updateNew(connection, fileName);
                    }

                    //----------------------------//
                    // Таблица discipline_plan_ed //
                    //----------------------------//
                    DisciplinePlanEd.update(connection, fileName);

                    //-----------------------------//
                    // Таблица course_disk_plan_ed //
                    //-----------------------------//
                    CourseDiscPlanEd.update(connection, fileName);

                    //----------------------//
                    // Таблица form_control //
                    //----------------------//
                    FormControl.update(connection, fileName);

                    //--------------------------//
                    // Таблица competence_codes //
                    //--------------------------//
                    CompetenceCodes.update(connection, fileName);

                    //--------------------------//
                    // Таблица competences_main //
                    //--------------------------//
                    CompetencesMain.update(connection, fileName);

                    //----------------------------//
                    // Таблица competences_matrix //
                    //----------------------------//
                    CompetencesMatrix.update(connection, fileName);

                    //--------------------------------//
                    // Таблица competences_teach_plan //
                    //--------------------------------//
                    CompetencesTeachPlan.update(connection, fileName);

                    break;
                }

                //--------------------------------------------------//
                // Добавление данных всех .xls файлов в базу данных //
                //--------------------------------------------------//
                case "update_all": {
                    String tableFormat = args[1];
                    File folder = new File("files");
                    String path = folder.getAbsolutePath();
                    File[] listOfFiles = folder.listFiles();
                    assert listOfFiles != null;

                    for (File listOfFile : listOfFiles) {
                        String fileName = path + "/" + listOfFile.getName();
                        System.out.println("\n");
                        System.out.println("Извлечение данных из файла: " + listOfFile.getName());

                        AllTables.updateAllTables(connection, fileName);
                    }

                    break;
                }
            }
        } catch (ArrayIndexOutOfBoundsException | SQLException e) {
            e.printStackTrace();
        }
    }
}