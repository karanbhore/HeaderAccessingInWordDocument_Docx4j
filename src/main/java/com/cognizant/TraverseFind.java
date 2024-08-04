package com.cognizant;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.ClassFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;

public class TraverseFind {

    /**
     * Example of how to find and replace text in a Word document
     * using traversal.
     */
    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("Automatic Date Change.docx"));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // Custom finder for Text elements
        ClassFinder textFinder = new ClassFinder(Text.class);
        new TraversalUtil(documentPart.getContent(), textFinder);

        // Process found Text elements
        for (Object o : textFinder.results) {
            Object unwrapped = XmlUtils.unwrap(o);

            if (unwrapped instanceof Text) {
                Text text = (Text) unwrapped;
                String textValue = text.getValue();

                if (textValue.trim().contains("Date")) {
                    System.out.println("Found 'Date' placeholder: " + textValue);

                    // Replace 'Date' placeholder with current date
                    SimpleDateFormat sdf = new SimpleDateFormat("dd MMMM yyyy", Locale.GERMAN);
                    String replacedValue = textValue.trim().replaceAll("Date", sdf.format(new Date()));
                    text.setValue(replacedValue);
                    System.out.println("Replaced with: " + replacedValue);
                }
            }
        }

        // Save the updated document
        wordMLPackage.save(new java.io.File("Updated_Automatic_Date_Change.docx"));
        System.out.println("Document saved successfully.");
    }
}
