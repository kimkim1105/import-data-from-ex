import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Date;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) throws ClassNotFoundException {
        Class.forName ("oracle.jdbc.OracleDriver");
        String jdbcURL = "jdbc:oracle:thin:";
        String username = "";
        String password = "";

        String excelFilePath = "file/...xlsx";

        int batchSize = 20;
        int countrow = 0;

        Connection connection = null;

        try {
            long start = System.currentTimeMillis();

            FileInputStream inputStream = new FileInputStream(excelFilePath);

            Workbook workbook = new XSSFWorkbook(inputStream);

            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();

            try {
                connection = DriverManager.getConnection(jdbcURL, username, password);
            } catch (SQLException e) {
                e.printStackTrace();
            }
            connection.setAutoCommit(false);

            String sql = "INSERT INTO ... () VALUES (?)";
            PreparedStatement statement = connection.prepareStatement(sql);

            int count = 0;

            rowIterator.next(); // skip the header row

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();

                    int columnIndex = nextCell.getColumnIndex();
                    System.out.println("index==========="+columnIndex);
                    switch (columnIndex) {
                        case 0:
                            Date ... = nextCell.getDateCellValue();
                            statement.setTimestamp(1, new Timestamp(....getTime()));
                            break;
                        case 1:
                            int ... = (int) nextCell.getNumericCellValue();
                            statement.setInt(2, ...);
                            break;
                    }
                }
                countrow = countrow + 1;
                System.out.println("count======"+countrow);
                statement.addBatch();

                if (count % batchSize == 0) {
                    statement.executeBatch();
                }

            }

            workbook.close();

            // execute the remaining queries
            statement.executeBatch();

            connection.commit();
            connection.close();

            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));

        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }

    }
}
