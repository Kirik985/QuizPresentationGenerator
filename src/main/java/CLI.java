import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class CLI {
    public static void main(String[] args) {
        new CLI().start();
    }
    public static Scanner scanner;
    public static String ask(String query) {
        System.out.print(query);
        return scanner.nextLine();
    }
    public static XSLFTextRun changeText(XSLFTextShape textShape, Object text) {
        XSLFTextRun run = textShape.getTextParagraphs().get(0).addNewTextRun();
        run.setText(String.valueOf(text));
        return run;
    }
    public static boolean checkPresentationFile(String presentationName) {
        File file = new File(presentationName);
        if (!file.exists()) {
            try {
                if (!file.createNewFile()) {
                    System.out.println("Error creating file!");
                    return true;
                }
                try (
                        XMLSlideShow slideShow = new XMLSlideShow();
                        FileOutputStream outputStream = new FileOutputStream(presentationName)
                ) {
                    System.out.println("Presentation not found, was created. Add needed slides and start again");
                    slideShow.write(outputStream);
                    Desktop.getDesktop().open(file);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return true;
        }
        try {
            Files.copy(Path.of(presentationName), Path.of("new-" + presentationName), StandardCopyOption.REPLACE_EXISTING);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }
    public void start() {
        scanner = new Scanner(System.in, StandardCharsets.UTF_8);
        System.out.print("Do you want to load config: ");
        String useConfigStr = scanner.nextLine();
        boolean useConfig = switch (useConfigStr) {
            case "y", "Y", "1", "yes", "YES", "true" -> true;
            default -> false;
        };
        String presentationName = "invalidName";
        int numCategories = -1;
        int numQuestions = -1;
        List<Integer> prices = List.of();
        List<Category> categories = List.of();
        if (useConfig) {
            Gson gson = new GsonBuilder().setPrettyPrinting().create();
            File configFile = new File("presentationConfig.json");
            if (!configFile.exists()) {
                try {
                    if (!configFile.createNewFile()) {
                        System.out.println("Error creating config file!");
                        return;
                    }
                    Config config = new Config("example.pptx", 2, 2, List.of(100, 200), List.of(
                            new Category("Category 1", List.of(
                                    new Question("Example Question 1", "Example Answer 1"),
                                    new Question("Example Question 2", "Example Answer 2")
                            )),
                            new Category("Category 2", List.of(
                                    new Question("Example Question 2.1", "Example Answer 2.1"),
                                    new Question("Example Question 2.2", "Example Answer 2.2")
                            ))
                    ));
                    String jsonString = gson.toJson(config);
                    FileOutputStream stream = new FileOutputStream(configFile);
                    stream.write(jsonString.getBytes(StandardCharsets.UTF_8));
                    stream.close();
                    System.out.println("Created config.json, edit it to customize questions and answers");
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return;
            } else {
                try {
                    FileInputStream stream = new FileInputStream(configFile);
                    String jsonString = new String(stream.readAllBytes(), StandardCharsets.UTF_8);
                    stream.close();
                    Config config = gson.fromJson(jsonString, Config.class);
                    presentationName = config.presentationName;
                    if (!presentationName.endsWith(".pptx")) {
                        presentationName += ".pptx";
                    }
                    if (checkPresentationFile(presentationName)) return;
                    presentationName = "new-" + presentationName;
                    numCategories = config.numCategories;
                    numQuestions = config.numQuestions;
                    prices = config.prices;
                    categories = config.categories;
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        } else {
            presentationName = ask("What is the name of your presentation: ");
            if (!presentationName.endsWith(".pptx")) {
                presentationName += ".pptx";
            }
            if (checkPresentationFile(presentationName)) return;
            presentationName = "new-" + presentationName;
            numCategories = Integer.parseInt(ask("How many categories do you want: "));
            numQuestions = Integer.parseInt(ask("How many questions do you want in each category: "));
            prices = new ArrayList<>();
            for (int i = 0; i < numQuestions; i++) {
                prices.add(i, Integer.parseInt(ask("Price for question number " + (i + 1) + ": ")));
            }
            categories = new ArrayList<>();
            for (int i = 0; i < numCategories; i++) {
                categories.add(i, new Category(ask("Name of category number " + (i + 1) + ": "), new ArrayList<>()));
                for (int j = 0; j < numQuestions; j++) {
                    categories.get(i).questions.add(new Question(ask("Question number " + (j + 1) + " in category " + categories.get(0).name + ": "), ask("Answer for question number " + (j + 1) + " in category " + categories.get(0).name + ": ")));
                }
            }
        }
        try (
             FileInputStream inputStream = new FileInputStream(presentationName);
             XMLSlideShow slideShow = new XMLSlideShow(inputStream);
             FileOutputStream outputStream = new FileOutputStream(presentationName)
        ) {
            XSLFSlideLayout questionLayout = slideShow.getSlideMasters().get(1).getLayout("Question");
            XSLFSlideLayout answerLayout = slideShow.getSlideMasters().get(1).getLayout("Answer");
            XSLFTable table = (XSLFTable) slideShow.getSlides().get(0).getShapes().get(0);
            for (int i = 0; i < numCategories; i++) {
                changeText(table.getCell(i, 0),categories.get(i).name);
                for (int j = 0; j < numQuestions; j++) {
                    XSLFTextRun priceRun = changeText(table.getCell(i, j + 1), prices.get(j));
                    XSLFSlide questionSlide = slideShow.createSlide(questionLayout);
                    questionSlide.getPlaceholder(0).appendText(categories.get(i).questions.get(j).question, false);
                    priceRun.createHyperlink().linkToSlide(questionSlide);
                    XSLFSlide answerSlide = slideShow.createSlide(answerLayout);
                    answerSlide.getPlaceholder(0).appendText(categories.get(i).questions.get(j).answer, false);
                }
            }
            slideShow.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public record Config(String presentationName, int numCategories, int numQuestions, List<Integer> prices, List<Category> categories) { }
    public record Question(String question, String answer) { }
    public record Category(String name, List<Question> questions) { }
}