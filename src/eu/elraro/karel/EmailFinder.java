package eu.elraro.karel;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

public class EmailFinder {

	private static Matcher matcher = null;

	private static List<String> emails = new ArrayList<String>();

	public static void main(String[] args) {

		if (args.length < 2 || args.length > 2) {
			System.out.println("Incorrect use.");
			System.out.println("java -jar EmailFinder.jar <Directory> <File where save emails>");
			System.exit(-1);
		}

		Path dir = null;

		try {
			dir = Paths.get(args[0]);
		} catch (InvalidPathException e) {
			System.out.println("Invalid directory.");
			System.exit(-1);
		}

		List<String> docs = getFileNames(new ArrayList<String>(), dir);

		System.out.println(".doc files found: " + docs.size());

		for (String doc : docs) {
			try {
				WordExtractor word = new WordExtractor(new HWPFDocument(new FileInputStream(doc)));
				String text = word.getText();

				matcher = Pattern.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+").matcher(text);
				while (matcher.find()) {
					emails.add(matcher.group());
				}

				word.close();
			} catch (IOException | OfficeXmlFileException e) {
				System.out.println(
						"There is a problem with the next file: "
								+ doc);
			}
		}

		FileWriter writer;
		try {
			writer = new FileWriter(args[1]);
			for (String email : emails) {
				writer.write(email);
				writer.write("\n");
			}
			writer.close();
		} catch (IOException e) {
			System.out.println("Invalid file or cannot open file.");
			System.exit(-1);
		}
		
		System.exit(0);
	}

	private static List<String> getFileNames(List<String> fileNames, Path dir) {
		try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir)) {
			for (Path path : stream) {
				if (path.toFile().isDirectory()) {
					getFileNames(fileNames, path);
				} else {
					if (path.toAbsolutePath().toString().endsWith(".doc")
							&& !path.getFileName().toString().startsWith("~$")) {
						fileNames.add(path.toAbsolutePath().toString());
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return fileNames;
	}

}
