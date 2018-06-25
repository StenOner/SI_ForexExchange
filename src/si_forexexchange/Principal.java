package si_forexexchange;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.auth.oauth2.ServiceAccountCredentials;
import com.google.cloud.storage.Blob;
import com.google.cloud.storage.BlobId;
import com.google.cloud.storage.BlobInfo;
import com.google.cloud.storage.Bucket;
import com.google.cloud.storage.Storage;
import com.google.cloud.storage.StorageOptions;
import java.awt.Color;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.ByteArrayOutputStream;
import org.apache.commons.codec.binary.Base64;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.security.InvalidAlgorithmParameterException;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.crypto.BadPaddingException;
import javax.crypto.Cipher;
import javax.crypto.IllegalBlockSizeException;
import javax.crypto.NoSuchPaddingException;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.event.HyperlinkEvent;
import javax.swing.event.HyperlinkListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import sun.java2d.loops.ProcessPath.ProcessHandler;

public class Principal extends javax.swing.JFrame {

    //private DefaultTableModel dtmExcel;
    private Workbook workbook, encryptedWorkbook;
    private Sheet sheet, encryptedSheet;
    private final String key = "ForexExchangeUPT"; // 128 bit key
    private final String initVector = "RandomInitVector"; // 16 bytes IV
    private final String encriptado = "_ENCRYPTED";
    private String nombreArchivo = "";
    private boolean statusEncrypted = false;
    private final String projectId = "nadeko-186003";
    private final String bucketName = "nadeko-186003.appspot.com";
    private Storage storage;
    private final String jsonPath = new File("src/Nadeko-7bc026a81f2f.json").getAbsolutePath();

    public Principal() throws IOException {
        initComponents();
        lblExcel.setVisible(false);
        btnEncriptar.setVisible(false);
        System.out.println(jsonPath);
        storage = StorageOptions.newBuilder()
                .setProjectId(projectId)
                .setCredentials(ServiceAccountCredentials.fromStream(new FileInputStream(jsonPath)))
                .build()
                .getService();
        /*
        Bucket bucket = storage.get(bucketName);
        for (Blob blob : bucket.list().iterateAll()) {
            System.out.println(blob.getSelfLink());
            System.out.println(blob.getMediaLink());
        }
         */
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        btnSalir = new javax.swing.JButton();
        lblExcel = new javax.swing.JLabel();
        btnFile = new javax.swing.JButton();
        btnEncriptar = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Forex Exchange");
        setResizable(false);

        btnSalir.setText("Salir");
        btnSalir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSalirActionPerformed(evt);
            }
        });

        lblExcel.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N
        lblExcel.setText("1");

        btnFile.setText("Elegir archivo");
        btnFile.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnFileActionPerformed(evt);
            }
        });

        btnEncriptar.setText("jButton1");
        btnEncriptar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEncriptarActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btnSalir, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnEncriptar, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnFile, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(18, 18, 18)
                .addComponent(lblExcel)
                .addContainerGap(480, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(lblExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                    .addComponent(btnFile, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(7, 7, 7)
                .addComponent(btnEncriptar)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(btnSalir)
                .addContainerGap(16, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnSalirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSalirActionPerformed
        System.exit(0);
    }//GEN-LAST:event_btnSalirActionPerformed

    private void btnFileActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnFileActionPerformed
        try {
            JFileChooser fileChooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel", "xls", "xlsx");
            fileChooser.setFileFilter(filter);
            fileChooser.setCurrentDirectory(javax.swing.filechooser.FileSystemView.getFileSystemView().getHomeDirectory());
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                lblExcel.setText(selectedFile.getName());
                nombreArchivo = selectedFile.getName();
                cargar(selectedFile.getAbsolutePath());
                lblExcel.setVisible(true);
            } else {
                throw new IOException("Error al seleccionar archivo.");
            }
        } catch (IOException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btnFileActionPerformed

    private void btnEncriptarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEncriptarActionPerformed
        String[] segmentosNombreArchivo = nombreArchivo.split("\\.");
        try {
            if (!statusEncrypted) {
                String rutaEncrypted = javax.swing.filechooser.FileSystemView.getFileSystemView().getHomeDirectory().toString();
                nombreArchivo = segmentosNombreArchivo[0];
                nombreArchivo += encriptado;
                nombreArchivo += "." + segmentosNombreArchivo[segmentosNombreArchivo.length - 1];
                rutaEncrypted += "\\" + nombreArchivo;
                FileOutputStream foutEncrypted = new FileOutputStream(rutaEncrypted);
                encryptedWorkbook.write(foutEncrypted);
                foutEncrypted.close();
                encryptedWorkbook.close();
                subirCloudStorage(rutaEncrypted);
            } else {
                String ruta = javax.swing.filechooser.FileSystemView.getFileSystemView().getHomeDirectory().toString();
                nombreArchivo = segmentosNombreArchivo[0];
                nombreArchivo += "." + segmentosNombreArchivo[segmentosNombreArchivo.length - 1];
                ruta += "\\" + nombreArchivo;
                FileOutputStream fout = new FileOutputStream(ruta);
                workbook.write(fout);
                fout.close();
                workbook.close();
                JOptionPane.showMessageDialog(null, "Se grabo en: " + ruta);
            }
            btnEncriptar.setVisible(false);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btnEncriptarActionPerformed

    private void cargar(String path) throws IOException, InvalidFormatException {
        ZipSecureFile.setMinInflateRatio(-1.0d);
        FileInputStream fis = new FileInputStream(new File(path));

        if (path.contains(encriptado)) {
            statusEncrypted = true;
            btnEncriptar.setText("Desencriptar");
            nombreArchivo = nombreArchivo.replace(encriptado, "");
        } else {
            statusEncrypted = false;
            btnEncriptar.setText("Encriptar");
        }
        //ArrayList<String> titulos = new ArrayList();
        //ArrayList<String> contenido = new ArrayList();
        //boolean firstRow = true;
        if (statusEncrypted) {
            workbook = WorkbookFactory.create(fis);
            sheet = workbook.getSheetAt(0);
            Iterator<Row> iteratorRow = sheet.iterator();
            while (iteratorRow.hasNext()) {
                Row encryptedRow = iteratorRow.next();
                Iterator<Cell> iteratorCell = encryptedRow.iterator();
                while (iteratorCell.hasNext()) {
                    Cell encryptedCell = iteratorCell.next();
                    String value = desencriptarData(encryptedCell.toString());
                    if (value.length() > 0 && String.valueOf(value.charAt(0)).equals("=")) {
                        value = value.replaceFirst("=", "");
                        encryptedCell.setCellType(CellType.FORMULA);
                        encryptedCell.setCellFormula(value);
                    } else {
                        encryptedCell.setCellValue(value);
                    }
                    /*if (firstRow) {
                        titulos.add(value);
                    } else {
                        contenido.add(value);
                    }*/
                }
                //firstRow = false;
            }
        } else {
            encryptedWorkbook = WorkbookFactory.create(fis);
            encryptedSheet = encryptedWorkbook.getSheetAt(0);
            Iterator<Row> iteratorRow = encryptedSheet.iterator();
            while (iteratorRow.hasNext()) {
                Row currentRow = iteratorRow.next();
                Iterator<Cell> iteratorCell = currentRow.iterator();
                while (iteratorCell.hasNext()) {
                    Cell currentCell = iteratorCell.next();
                    String value = "";
                    if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
                        value += "=";
                        value += currentCell.getCellFormula();
                        currentCell.setCellType(CellType.STRING);
                        //currentCell.setCellFormula(encriptarData(value));
                    } else {
                        value += currentCell.toString();
                    }
                    /*if (firstRow) {
                        titulos.add(value);
                    } else {
                        contenido.add(value);
                    }*/
                    currentCell.setCellValue(encriptarData(value));
                }
                //firstRow = false;
            }
            /*
            int rowSize = titulos.size();
            String[] arrayTitulos = new String[rowSize];
            arrayTitulos = titulos.toArray(arrayTitulos);
            dtmExcel = new DefaultTableModel(null, arrayTitulos);
            tblExcel.setModel(dtmExcel);
            int contadorRowEnd = 0;
            Object[] fila = new Object[rowSize];
            for (String currentContent : contenido) {
                if (contadorRowEnd == rowSize) {
                    contadorRowEnd = 0;
                    dtmExcel.addRow(fila);
                    fila = new Object[rowSize];
                }
                fila[contadorRowEnd] = currentContent;
                contadorRowEnd++;
            }
             */
        }
        btnEncriptar.setVisible(true);
    }

    private void subirCloudStorage(String path) {
        try {
            File file = new File(path);
            ByteArrayOutputStream output = new ByteArrayOutputStream(1024);
            int byteReads;
            InputStream inputStream = new FileInputStream(file);
            while ((byteReads = inputStream.read()) != -1) {
                output.write(byteReads);
            }
            byte[] bytes = output.toByteArray();

            Bucket bucket = storage.get(bucketName);
            BlobId blobId = BlobId.of(bucket.getName(), nombreArchivo);
            BlobInfo blobInfo = BlobInfo.newBuilder(blobId).setContentType("application/octet-stream").build();
            storage.create(blobInfo, bytes);
            String link = "https://storage.cloud.google.com/" + bucketName + "/" + nombreArchivo;

            JEditorPane ep = new JEditorPane("text/html", "<html><body>"
                    + "Se subio con exito en:<a href=\"" + link + "\">" + link + "</a>" //
                    + "</body></html>");
            ep.addHyperlinkListener(new HyperlinkListener() {
                @Override
                public void hyperlinkUpdate(HyperlinkEvent e) {
                    if (e.getEventType().equals(HyperlinkEvent.EventType.ACTIVATED)) {
                        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                        StringSelection linkSelection = new StringSelection(link);
                        clipboard.setContents(linkSelection, null);
                    }
                }
            });
            ep.setEditable(false);
            ep.setBackground(Color.getColor("00DFD9D6"));
            JOptionPane.showMessageDialog(null, ep);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    private String desencriptarData(String encryptedInput) {
        String result;
        try {
            IvParameterSpec iv = new IvParameterSpec(initVector.getBytes("UTF-8"));
            SecretKeySpec skeySpec = new SecretKeySpec(key.getBytes("UTF-8"), "AES");

            Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5PADDING");
            cipher.init(Cipher.DECRYPT_MODE, skeySpec, iv);

            byte[] original = cipher.doFinal(Base64.decodeBase64(encryptedInput));
            result = new String(original);
            return result;
        } catch (UnsupportedEncodingException | InvalidAlgorithmParameterException | InvalidKeyException | NoSuchAlgorithmException | BadPaddingException | IllegalBlockSizeException | NoSuchPaddingException ex) {
            ex.printStackTrace();
            return null;
        }
    }

    private String encriptarData(String input) {
        String result;
        try {
            IvParameterSpec iv = new IvParameterSpec(initVector.getBytes("UTF-8"));
            SecretKeySpec skeySpec = new SecretKeySpec(key.getBytes("UTF-8"), "AES");

            Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5PADDING");
            cipher.init(Cipher.ENCRYPT_MODE, skeySpec, iv);

            byte[] encrypted = cipher.doFinal(input.getBytes());
            result = Base64.encodeBase64String(encrypted);
            return result;
        } catch (UnsupportedEncodingException | InvalidAlgorithmParameterException | InvalidKeyException | NoSuchAlgorithmException | BadPaddingException | IllegalBlockSizeException | NoSuchPaddingException ex) {
            ex.printStackTrace();
            return null;
        }
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    new Principal().setVisible(true);
                } catch (IOException ex) {
                    Logger.getLogger(Principal.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnEncriptar;
    private javax.swing.JButton btnFile;
    private javax.swing.JButton btnSalir;
    private javax.swing.JLabel lblExcel;
    // End of variables declaration//GEN-END:variables

}
