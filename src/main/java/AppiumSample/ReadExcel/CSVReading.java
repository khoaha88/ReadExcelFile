package AppiumSample.ReadExcel;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.Scanner;

public class CSVReading {
	public static void readCSVFile() throws FileNotFoundException{
		Scanner scanner = new Scanner(new File("setup.csv"));
        scanner.useDelimiter("=");
        while(scanner.hasNext()){
            System.out.print(scanner.next()+"|");
        }
        scanner.close();
	}
}
