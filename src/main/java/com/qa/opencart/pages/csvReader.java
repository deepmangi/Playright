package com.qa.opencart.pages;

import java.io.FileReader;
import java.io.IOException;
import java.net.http.HttpResponse;
import java.util.List;

import com.microsoft.playwright.Browser;
import com.microsoft.playwright.BrowserContext;
import com.microsoft.playwright.Playwright;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;

public class csvReader {

    public static void main(String[] args) throws IOException, CsvException {
        try (Playwright playwright = Playwright.create()) {
            Browser browser = playwright.chromium().launch();
            BrowserContext context = browser.newContext();

            // Define the API endpoint URL
            String apiUrl = "https://api.example.com/data";

            // Read test data from a CSV file
            CSVReader csvReader = new CSVReader(new FileReader("testdata.csv"));
            List<String[]> testData = csvReader.readAll();

            for (String[] data : testData) {
                // Extract data from the CSV
                String param1 = data[0];
                String param2 = data[1];

               System.out.println(param1);
               System.out.println(param2);
            }

            // Close the browser
            browser.close();
        }
    }
}