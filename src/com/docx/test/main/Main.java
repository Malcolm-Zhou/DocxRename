package com.docx.test.main;

import org.apache.commons.io.FileUtils;
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
    public static void main(String[] args) throws IOException {
        File folder = new File("files");
        String[] files = folder.list();

        String slash = "\\";
        //  Mac use "/"
        //  Windows use "\\"

        for (String file : files) {
            File f = new File(folder, file);
            String fullFileName = f.getName();
            String fileName = fullFileName.substring(0, fullFileName.lastIndexOf("."));
            String fullPath = folder.getAbsolutePath() + slash + fullFileName;
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


                    File toFile = new File(folder.getAbsolutePath().replace("files", "renamed") + slash + fullFileName.replace(fileName, title));
                    if (!toFile.exists()) {
                        FileUtils.copyFile(f, toFile);
                    } else {
                        File renamedFolder = new File("renamed");
                        String[] renameds = renamedFolder.list();
                        Integer copyCount = 0;
                        for (String renamed : renameds) {
                            System.out.println(renamed);
                            if (renamed.contains(title)) {
                                copyCount++;
                            }
                        }
                        title = title + " (" + copyCount + ")";
                        File toFileCopy = new File(folder.getAbsolutePath().replace("files", "renamed") + slash + fullFileName.replace(fileName, title));
                        FileUtils.copyFile(f, toFileCopy);
                    }
                } else if (fullFileName.endsWith(".doc")) {
                    try {
                        InputStream is = new FileInputStream(new File(fullPath));
                        WordExtractor extractor = new WordExtractor(is);
                        String text = extractor.getText();
                        String title = (text.split("\n"))[0];
                        System.out.println(title);

                        File toFile = new File(folder.getAbsolutePath().replace("files", "renamed") + slash + fullFileName.replace(fileName, title));
                        if (!toFile.exists()) {
                            FileUtils.copyFile(f, toFile);
                        } else {
                            File renamedFolder = new File("renamed");
                            String[] renameds = renamedFolder.list();
                            Integer copyCount = 0;
                            for (String renamed : renameds) {
                                System.out.println(renamed);
                                if (renamed.contains(title)) {
                                    copyCount++;
                                }
                            }
                            title = title + " (" + copyCount + ")";
                            File toFileCopy = new File(folder.getAbsolutePath().replace("files", "renamed") + slash + fullFileName.replace(fileName, title));
                            System.out.println("Copy");
                            FileUtils.copyFile(f, toFileCopy);
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
