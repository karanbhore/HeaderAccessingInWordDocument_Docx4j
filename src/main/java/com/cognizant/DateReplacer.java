package com.cognizant;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.ClassFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;

public class DateReplacer {

    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("Automatic Date Change.docx"));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // Process headers and footers
        processHeadersAndFooters(documentPart);

        // Process main document
        processContent(documentPart.getContent());

        wordMLPackage.save(new java.io.File("Automatic Date Change.docx"));
        System.out.println("Document saved successfully.");
    }

    private static void processHeadersAndFooters(MainDocumentPart documentPart) throws Exception {
        RelationshipsPart relationshipsPart = documentPart.getRelationshipsPart();
        List<org.docx4j.relationships.Relationship> relationships = relationshipsPart.getRelationships().getRelationship();

        for (org.docx4j.relationships.Relationship relationship : relationships) {
            if (relationship.getType().equals(org.docx4j.openpackaging.parts.relationships.Namespaces.HEADER)) {
                HeaderPart headerPart = (HeaderPart) relationshipsPart.getPart(relationship);
                processContent(headerPart.getContent());
            } else if (relationship.getType().equals(org.docx4j.openpackaging.parts.relationships.Namespaces.FOOTER)) {
                FooterPart footerPart = (FooterPart) relationshipsPart.getPart(relationship);
                processContent(footerPart.getContent());
            }
        }
    }

    private static void processContent(List<Object> content) {
        ClassFinder finder = new ClassFinder(Text.class);
        new TraversalUtil(content, finder);

        for (Object o : finder.results) {
            Object o2 = XmlUtils.unwrap(o);

            if (o2 instanceof Text) {
                Text txt = (Text) o2;
                String value = txt.getValue();

                // Replace date placeholders
                if (value != null && value.trim().contains("Date")) {
                    SimpleDateFormat sdf = new SimpleDateFormat("dd MMMM yyyy", Locale.GERMAN);
                    String replacedValue = value.trim().replaceAll("Date", sdf.format(new Date()));
                    txt.setValue(replacedValue);
                    System.out.println("Replaced 'Date' with '" + replacedValue + "'");
                }
            }
        }
    }
}
