import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // Problem 1:
        System.out.println("Please input a word to be reversed: ");
        String word = scanner.nextLine();
        System.out.println("The reversed word is: " + reverse(word));

        // Problem 2:
        System.out.println("Please input words separated by a white space to be counted: ");
        String[] countWords = scanner.nextLine().split(" ");
        HashMap<String, Integer> countedWords = countWords(countWords);
        System.out.println("The number of occurrences for each word");
        for (Map.Entry<String, Integer> entry : countedWords.entrySet()) {
            System.out.println(entry.getKey() + ":" + entry.getValue());
        }

        // Problem 3:
        System.out.println("Please input strings, separated by a white space to be converted into Arraylist," +
                " traversed and printed: ");
        String[] wordsList = scanner.nextLine().split(" ");
        List<String> list = Arrays.asList(wordsList);
        traverseArrayList(list);

        // Problem 4: Very similar to the second problem
        System.out.println("Please input a word string, in which the characters are to be counted: ");
        String characters = scanner.nextLine();
        Map<Character, Integer> countedChars = countCharacters(characters);
        // For how to make the algorithm more efficient check out the comments on the helper method countCharacters
        // This way it is more correctly structured as the I/O is in the main method only
        System.out.println("The number of occurrences for the duplicate characters in the word");
        for (Map.Entry<Character, Integer> entry : countedChars.entrySet()) {
            if (entry.getValue() > 1) {
                System.out.println(entry.getKey() + ":" + entry.getValue());
            }
        }

        //Problem 5:
        System.out.println("Please input the path to your Excel file: ");
        String filePath = scanner.nextLine();
        // Determine whether it is a before 2003 (binary excel) or OOXML (after 2003) format
        if (filePath.charAt(filePath.length() - 1) == 'x') {
            try (FileInputStream file = new FileInputStream(filePath);
                 XSSFWorkbook workbook = new XSSFWorkbook(file)){

                // for each sheet
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    XSSFSheet workSheet = workbook.getSheetAt(i);
                    // For each row
                    readRows(workSheet);
                }
            } catch (IOException e) {
                System.out.println("The file could not be read or is not in the correct format");
                System.out.println("Additional information: " + e.getMessage());
                e.printStackTrace();
            }
        }
         else {
            try (FileInputStream file = new FileInputStream(filePath);
                 HSSFWorkbook workbook = new HSSFWorkbook(file)){

                // get each sheet and read it
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    HSSFSheet workSheet = workbook.getSheetAt(i);
                    readRows(workSheet);
                }
            } catch (IOException e) {
                System.out.println("The file could not be read or is not in the correct format");
                System.out.println("Additional information: " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    private static String reverse(String word) {
        // Stringbuilder is more efficient to my knowledge,
        // otherwise simple String concatenation in the loop is also an option
        StringBuilder reversed = new StringBuilder();
        for (int i = word.length() - 1; i >= 0; i--) {
            reversed.append(word.charAt(i));
        }
        return reversed.toString();
    }

    // I am using List instead of ArrayList, but it is it's generic type,
    // so for this example the implementation does not differ
    private static void traverseArrayList(List<String> list) {
        // First using an enhanced for loop
        for (String word : list) {
            System.out.println(word);
        }
        // Then traversing using an iterator and a while loop
        Iterator<String> iterator = list.iterator();
        while (iterator.hasNext()) {
            System.out.println(iterator.next());
        }
    }

    private static HashMap<String, Integer> countWords(String[] words) {
        HashMap<String, Integer> countWords = new HashMap<>();
        // Check if each word is already present in the hashmap, if not add it, else increment
        for (String word : words) {
            Integer value = countWords.get(word);
            if (value == null) {
                // putting the word in the hashmap with value 1 for occurrence = 1
                countWords.put(word, 1);
            } else {
                // incrementing the amount of occurrences for the already present word
                countWords.put(word, ++value);
            }
        }
        return countWords;
    }

    private static HashMap<Character, Integer> countCharacters(String word) {
        // The same as countWords but with characters
        HashMap<Character, Integer> countChars = new HashMap<>();
        for (int i = 0; i < word.length(); i++) {
            Integer value = countChars.get(word.charAt(i));
            if (value == null) {
                countChars.put(word.charAt(i), 1);
            } else {
                // As we do not need the number of occurrences of the searched duplicate characters,
                // instead of putting, we could simply print out the duplicate character here, so we do not have to loop
                // through the hashmap in the main method again, and we also potentially save up to n put operations
                countChars.put(word.charAt(i), ++value);
            }
        }
        return countChars;
    }

    private static void readRows(Sheet workSheet) {
        // For each row
        for (Row row : workSheet) {
            // Iterate through the cells
            Iterator<Cell> cellIterator = row.cellIterator();
            // Print out the cell values separated by |
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.print(cell.getStringCellValue() + " | ");
            }
            System.out.println();
        }
    }

}

