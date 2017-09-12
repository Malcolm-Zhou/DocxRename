package com.docx.test.main;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.POIXMLDocument;
import org.apache.xmlbeans.XmlException;

import java.io.*;

/**
 * @author Malcolm on 12/09/2017.
 */
public class Main {
    public static void main(String[] args) {
        File folder = new File("files");
        String[] files = folder.list();

        for (String file : files) {
            File f = new File(folder, file);
            String fullFileName = f.getName();
            String fileName = fullFileName.substring(0, fullFileName.lastIndexOf("."));
            String fullPath = folder.getAbsolutePath() + "//" + fullFileName;
            if (fullFileName.contains(".doc")) {
                if (fullFileName.endsWith(".docx")) {
                    XWPFWordExtractor extractor = null;
                    try {
                        extractor = new XWPFWordExtractor(POIXMLDocument.openPackage(fullPath));
                    } catch (XmlException e) {
                        e.printStackTrace();
                    } catch (OpenXML4JException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    String text = extractor.getText();
                    String title = (text.split("\n"))[0];
                    System.out.println(title);

                    File toFile = new File(folder.getAbsolutePath().replace("files", "renamed") + "//" + fullFileName.replace(fileName, title));
                    if (!toFile.exists()) {
                        f.renameTo(toFile);
                    } else {
                        title = title + " copy";
                        File toFileCopy = new File(folder.getAbsolutePath().replace("files", "renamed") + "//" + fullFileName.replace(fileName, title));
                        f.renameTo(toFileCopy);
                    }
                } else if (fullFileName.endsWith(".doc")) {
                    try {
                        InputStream is = new FileInputStream(new File(fullPath));
                        WordExtractor extractor = new WordExtractor(is);
                        String text = extractor.getText();
                        String title = (text.split("\n"))[0];
                        System.out.println(title);

                        File toFile = new File(folder.getAbsolutePath().replace("files", "renamed") + "//" + fullFileName.replace(fileName, title));
                        if (!toFile.exists()) {
                            f.renameTo(toFile);
                        } else {
                            title = title + " copy";
                            File toFileCopy = new File(folder.getAbsolutePath().replace("files", "renamed") + "//" + fullFileName.replace(fileName, title));
                            f.renameTo(toFileCopy);
                        }
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }

            }

        }


    }
}
