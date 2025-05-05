import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class UpdateCsv {

    public static void main(String[] args) {
        String filePath = "data.csv"; // Replace with your CSV file path
        int rowToUpdate = 2;          // Row number to update (0-based index)
        int columnToUpdate = 1;       // Column number to update (0-based index)
        String newValue = "Updated Value";

        try {
            List<String[]> rows = new ArrayList<>();
            BufferedReader reader = new BufferedReader(new FileReader(filePath));
            String line;
            int rowCount = 0;

            while ((line = reader.readLine()) != null) {
                String[] fields = line.split(",");
                if (rowCount == rowToUpdate) {
                    if (columnToUpdate < fields.length) {
                        fields[columnToUpdate] = newValue;
                    }
                }
                rows.add(fields);
                rowCount++;
            }
            reader.close();

            BufferedWriter writer = new BufferedWriter(new FileWriter(filePath));
            for (String[] rowData : rows) {
                writer.write(String.join(",", rowData));
                writer.newLine();
            }
            writer.close();

            System.out.println("Successfully updated the CSV file.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
