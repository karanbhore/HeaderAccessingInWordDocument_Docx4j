package com.cognizant;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.vml.CTTextbox;
import org.docx4j.wml.CTTxbxContent;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DocxReplace {

    private static final Logger logger = LoggerFactory.getLogger(DocxReplace.class);

    public static void replaceElementInTextBoxInHeader(WordprocessingMLPackage template) throws Docx4JException {
        logger.info("Starting the text box replacement process in the header.");

        RelationshipsPart relationshipPart = template.getMainDocumentPart().getRelationshipsPart();
        List<Relationship> relationships = relationshipPart.getRelationships().getRelationship();

        for (Relationship r : relationships) {
            if (r.getType().equals(Namespaces.HEADER)) {
                JaxbXmlPart part = (JaxbXmlPart) relationshipPart.getPart(r);
                List<Object> textBoxContents = getAllElementFromObject(part.getContents(), CTTextbox.class);

                for (Object element : textBoxContents) {
                    if (element instanceof CTTextbox) {
                        CTTxbxContent content = ((CTTextbox) element).getTxbxContent();
                        List<Object> textBoxTexts = getAllElementFromObject(content, Text.class);
                        for (Object textElement : textBoxTexts) {
                            Text text = (Text) textElement;
                            logger.info("Found text in text box: '{}'", text.getValue());
                            if (text.getValue().trim().contains("Date")) {
                                logger.info("Found a 'Date' placeholder in a text box. Replacing it with the current date.");
                                SimpleDateFormat sdf = new SimpleDateFormat("dd MMMM yyyy", Locale.GERMAN);
                                String replacedValue = text.getValue().trim().replaceAll("Date", sdf.format(new Date()));
                                text.setValue(replacedValue);
                                logger.info("Replaced 'Date' with '{}'", replacedValue);
                            } else {
                                logger.info("No 'Date' placeholder found in this text: '{}'", text.getValue());
                            }
                        }
                    }
                }
            }
        }
        logger.info("Text box replacement process in the header completed.");
    }

    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
            if (obj.getClass().equals(toSearch)) {
                result.add(obj);
            }
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    public static void main(String[] args) {
        try {
            logger.info("Loading the document.");
            WordprocessingMLPackage template = WordprocessingMLPackage.load(new java.io.File("Automatic Date Change.docx"));
            replaceElementInTextBoxInHeader(template);
            template.save(new java.io.File("Automatic Date Change.docx"));
            logger.info("Document saved successfully.");
        } catch (Exception e) {
            logger.error("An error occurred while processing the document.", e);
        }
    }
}
