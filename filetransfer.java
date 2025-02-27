import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.stream.Stream;

public class FileArchiver {

    public static void main(String[] args) {

        String sourceDirectory = "path/to/your/source/directory"; // Replace with your source directory path
        String archivedDirectory = "path/to/your/archived/directory"; // Replace with your archived directory path
        String inputDirectory = "C:/Users/yourUserName/Desktop/input"; // Replace with your input directory path

        Path sourcePath = Paths.get(sourceDirectory);
        Path archivedPath = Paths.get(archivedDirectory);
        Path inputPath = Paths.get(inputDirectory);

        try {
            // Create directories if they don't exist
            Files.createDirectories(archivedPath);
            Files.createDirectories(inputPath);

            // Move files to archived
            archiveFiles(sourcePath, archivedPath);

            //Move newly created files to input folder
            moveNewFilesToInput(sourcePath, inputPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void archiveFiles(Path source, Path archive) throws IOException {
        try (Stream<Path> files = Files.list(source)) {
            files.filter(Files::isRegularFile)
                    .forEach(file -> {
                        try {
                            Path target = archive.resolve(file.getFileName());
                            Files.move(file, target, StandardCopyOption.REPLACE_EXISTING);
                            System.out.println("Moved file: " + file.getFileName() + " to archive.");
                        } catch (IOException e) {
                            System.err.println("Error moving file: " + file.getFileName() + " - " + e.getMessage());
                        }
                    });
        }
    }

    private static void moveNewFilesToInput(Path source, Path input) throws IOException{
        try (Stream<Path> files = Files.list(source)) {
            files.filter(Files::isRegularFile)
                    .forEach(file -> {
                        try {
                            BasicFileAttributes attributes = Files.readAttributes(file, BasicFileAttributes.class);
                            if(attributes.creationTime().toMillis() > System.currentTimeMillis() - 60000) { //Check if created within the last minute. Adjust time as needed.
                                Path target = input.resolve(file.getFileName());
                                Files.move(file, target, StandardCopyOption.REPLACE_EXISTING);
                                System.out.println("Moved new file: " + file.getFileName() + " to input.");
                            }

                        } catch (IOException e) {
                            System.err.println("Error moving file: " + file.getFileName() + " - " + e.getMessage());
                        }
                    });
        }
    }
}
