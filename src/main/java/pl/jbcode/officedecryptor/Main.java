package pl.jbcode.officedecryptor;

import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.ZipParameters;
import org.apache.commons.io.FileUtils;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.swing.*;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;

public class Main {
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        MainForm frm = new MainForm();
        frm.btnDoIt.addActionListener(evt -> {
            try {
                //get text fields
                String path = frm.txtFileIn.getText().trim();
                String readPassword = frm.txtReadPW.getText().trim();
                String save = frm.txtFileOut.getText().trim();

                //input validation
                if (path.equals("") || readPassword.equals("") || save.equals("")) {
                    JOptionPane.showMessageDialog(frm, "Please fill in all text fields.", "Error", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                if (path.equals(save)) {
                    JOptionPane.showMessageDialog(frm, "Save file needs to be different from input file.", "Error", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                //read file from specified path
                POIFSFileSystem fs = new POIFSFileSystem(new File(path), true);
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor d = Decryptor.getInstance(info);

                //try decrypting document
                if (!d.verifyPassword(readPassword)) {
                    JOptionPane.showMessageDialog(frm, "Invalid password", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                //get decrypted file content
                InputStream dataStream = d.getDataStream(fs);

                //write file to disk
                File targetFile = new File(save);
                FileUtils.copyInputStreamToFile(dataStream, targetFile);

                //if we only want to remove read-protection, we are done here
                if (!frm.cbRemoveWrite.isSelected()) {
                    JOptionPane.showMessageDialog(frm, "Success! Removed read password.\n\nWritten file to:\n\n" + save, "Success", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }

                //if we want to remove write-protection we need to edit files inside the Office file container
                ZipFile zip = new ZipFile(targetFile);

                //find the document descriptor basing on file type
                String target = null;
                if (path.endsWith(".pptx")) {
                    target = "ppt/presentation.xml";
                } else if (path.endsWith(".docx")) {
                    target = "word/document.xml";
                } else if (path.endsWith(".xlsx")) {
                    target = "xl/workbook.xml";
                }

                if (target != null) {
                    //get stream for the file inside container
                    InputStream is = zip.getInputStream(zip.getFileHeader(target));

                    //init xml reader
                    DocumentBuilderFactory fac = DocumentBuilderFactory.newInstance();
                    Document doc = fac.newDocumentBuilder().parse(is);

                    //this node specifies the write-protection flag
                    NodeList list = doc.getElementsByTagName("p:modifyVerifier");

                    //if it exists, remove all occurences
                    if (list.getLength() > 0) {
                        for (int i = 0; i < list.getLength(); i++) {
                            Node self = list.item(i);
                            self.getParentNode().removeChild(self);
                        }
                    } else
                    //if not, this means file is not write-protected and we abort here
                    {
                        JOptionPane.showMessageDialog(frm, "File is not write protected. Removed only read password.\n\nWritten file to:\n\n" + save, "Success", JOptionPane.INFORMATION_MESSAGE);
                        return;
                    }

                    //init xml writer
                    TransformerFactory tf = TransformerFactory.newInstance();
                    Transformer trans = tf.newTransformer();
                    trans.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
                    trans.setOutputProperty(OutputKeys.INDENT, "yes");
                    StringWriter sw = new StringWriter();
                    //write the xml content to utf8 string
                    trans.transform(new DOMSource(doc), new StreamResult(sw));
                    String result = sw.toString();

                    //remove the old entry from container
                    zip.removeFile(target);

                    //create new entry
                    ZipParameters p = new ZipParameters();
                    p.setFileNameInZip(target);
                    //write new entry stream
                    zip.addStream(new ByteArrayInputStream(result.getBytes(StandardCharsets.UTF_8)), p);

                    JOptionPane.showMessageDialog(frm, "Success! Removed read and write password.\n\nWritten file to:\n\n" + save, "Success", JOptionPane.INFORMATION_MESSAGE);
                } else
                //currently only excel,word,ppt files are supported for write-protection removal
                {
                    JOptionPane.showMessageDialog(frm, "Warning! Unsupported file format for removing write password. Removed only read password.\n\nWritten file to:\n\n" + save, "Success", JOptionPane.WARNING_MESSAGE);
                }

            } catch (Throwable ex) {
                JOptionPane.showMessageDialog(frm, "Error processing file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                ex.printStackTrace();
            }
        });

        frm.setLocationRelativeTo(null);
        frm.setVisible(true);
    }

}
