/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package infowriter;

import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import static javafx.application.Platform.exit;
import javafx.stage.FileChooser;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import org.apache.commons.compress.archivers.dump.DumpArchiveEntry;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakClear;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
/**
 *
 * @author HP
 */
public class InfoWriter extends JFrame{
    private JButton choose_excel;
    private JButton choose_word;
    private JButton fill;
    private JLabel label_excel, label_word;
    private File file_excel,file_word;
    private String function;
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new InfoWriter().setVisible(true);
            }
        });
        // TODO code application logic here
    }
    public InfoWriter(){
        choose_excel = new JButton("Choisir un fichier excel");
        label_excel = new JLabel();
        choose_word = new JButton("choisir un fichier word");
        label_word = new JLabel();
        fill = new JButton("Remplir");
        choose_excel.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                JFileChooser c = new JFileChooser();
                int rVal = c.showOpenDialog(InfoWriter.this);
                if (rVal == JFileChooser.APPROVE_OPTION) {
                   file_excel = new File("");
                   file_excel = c.getSelectedFile(); 
                   label_excel.setText(file_excel.getName());
                } 
         }});
        choose_word.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                JFileChooser c = new JFileChooser();
                int rVal = c.showOpenDialog(InfoWriter.this);
                if (rVal == JFileChooser.APPROVE_OPTION) {
                   file_word = new File("");
                   file_word = c.getSelectedFile(); 
                   label_word.setText(file_word.getName());
                } 
         }});
        Path path = Paths.get("C:\\Optimum");
         if (!Files.exists(path)) {
            try {
                Files.createDirectories(path);
            } catch (IOException e) {
                //fail to create directory
                e.printStackTrace();
            }
        }
         fill.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                try {
                    Fill();
                } catch (IOException ex) {
                    Logger.getLogger(InfoWriter.class.getName()).log(Level.SEVERE, null, ex);
                } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException ex) {
                    Logger.getLogger(InfoWriter.class.getName()).log(Level.SEVERE, null, ex);
                }
         }});
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints(0, 0, 1, 1, 1.0, 1.0,
            GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(
                  50, 10, 0, 0), 0, 0);
        GridBagConstraints gbc2 = new GridBagConstraints(0, 0, 1, 1, 1.0, 1.0,
            GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(
                  0, 0, 0, 0), 0, 0);
        panel.add(choose_excel,gbc);
        gbc.gridx = 0;
        gbc.gridy = 1;
       // panel.add(choose_word,gbc);
        gbc.gridx = 1;
        gbc.gridy = 0;
        panel.add(label_excel,gbc);
        gbc.gridx = 1;
        gbc.gridy = 1;
       // panel.add(label_word,gbc);
        gbc.gridx = 0;
        gbc.gridy = 2;
        panel.add(fill,gbc);
        this.setLayout(new GridBagLayout());
        this.getContentPane().add(panel,gbc2);
        panel.setOpaque(false);
        this.setSize(500, 500);
        this.setResizable(false);
        this.setFocusable(true);
        this.setVisible(true);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
   }
    
    public void Fill() throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException{
       Workbook workbook = null;
         if (file_excel == null){
            JOptionPane.showMessageDialog(null, "Veuillez sélectionner un fichier au format excel");
        }
        else{
            workbook = WorkbookFactory.create(file_excel);
         }
        Sheet firstSheet = workbook.getSheetAt(0);
        Row row;
       // XWPFDocument document = new XWPFDocument();

      //Write the Document in file system
      //FileOutputStream out = null;
        /*if (file_word == null){
            JOptionPane.showMessageDialog(null, "Veuillez sélectionner un fichier au format word");
        }
        else{
            out = new FileOutputStream(file_word);
        }*/
      //create Paragraph
      XWPFParagraph paragraph ;
      XWPFRun run;
     
      for (int i = 1; i <= firstSheet.getLastRowNum (); i++) {
           
            row=(Row) firstSheet.getRow(i);
            String surname = row.getCell(0).getStringCellValue();
            String name = row.getCell(1).getStringCellValue();
            String ID = row.getCell(4).getStringCellValue();
            File f = new File("C:\\Optimum\\contrat"+" "+ID+" "+surname+" "+name+".docx");
            f.createNewFile();
            XWPFDocument document = new XWPFDocument();
            
            CTDocument1 document1 = document.getDocument();
            CTBody body = document1.getBody();
            if (!body.isSetSectPr()) {
                 body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();
            if(!section.isSetPgSz()) {
                section.addNewPgSz();       
            }
            CTPageSz pageSize = section.getPgSz();
            pageSize.setW(BigInteger.valueOf(12240));
            pageSize.setH(BigInteger.valueOf(20160));
            FileOutputStream out = null;
            out = new FileOutputStream(f);
            /*if (policy.getDefaultHeader() == null) {
   // Need to create some new headers
   // The easy way, gives a single empty paragraph
   XWPFHeader headerD = policy.createHeader(policy.DEFAULT);
   headerD.addPictureData(new FileInputStream("image.png"), NORMAL);
            }*/            
            //System.out.println(row.getCell(0).getStringCellValue());
            paragraph = document.createParagraph();
             run = paragraph.createRun();
             //run.addTab();
             run.setText("                   ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("CONTRAT DE TRAVAIL À DURÉE DÉTERMINÉE À TERME IMPRÉCIS");
             run.addBreak();
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("          I-          ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("IDENTIFICATION DES PARTIES");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Entre les soussignés");
             run.addBreak();
             run.addBreak();
             run.setText("OPTIMUM INTERNATIONAL");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", SARL, au capital de 1 000 000 FCFA, dont le siège est sis ");
             run.addBreak();
             run.setText("Abidjan, Cocody Opération Latrille II plateaux, Aghien Las Palmas, 01 BP 5755 Abidjan 01, ");
             run.addBreak();
             run.setText("téléphones  59 11 00 00/22 52 26 45, immatriculée au Régistre du Commerce et du Crédit Mobilier ");
             run.setText("sous le numéro ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("CI- ABJ-2014-B-21785");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", représentée par ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Monsieur TRAORE Oumar");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", en sa ");
             run.setText("qualité de ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Co-gérant ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(";");
             run.addBreak();
             run.addBreak();
             run.setText("Ci-après désignée « l’Employeur »                                                                             d’une part ;");
             run.addBreak();
             run.addBreak();
             run.setText("ET ");
             run.addBreak();
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             //run.setBold(false);
             //run.setUnderline(UnderlinePatterns.NONE);
             
             run.setText("Nom: "+surname);
             run.addBreak();
             run.addBreak();
             
             run.setText("Prénom(s): "+name);
             run.addBreak();
             run.addBreak();
             String birthday = row.getCell(2).getStringCellValue();
             run.setText("Date et lieu de naissance: "+birthday);
             run.addBreak();
             run.addBreak();
             String nationality = row.getCell(3).getStringCellValue();
             run.setText("Nationalité: "+nationality);
             run.addBreak();
             run.addBreak();
             
             run.setText("Référence pièce d'identité: "+ID);
             run.addBreak();
             run.addBreak();
             String home = row.getCell(5).getStringCellValue();
             run.setText("Domicile: "+home);
             run.addBreak();
             run.addBreak();
             String representative = row.getCell(6).getStringCellValue();
             run.setText("représentant(e) légal(e): "+representative);
             run.addBreak();
             run.addBreak();
             String representative_phone = row.getCell(7).getStringCellValue();
             run.setText("Contacts représentant(e) légal(e): "+representative_phone);
             run.addBreak();
             run.addBreak();
             String email = row.getCell(8).getStringCellValue();
             run.setText("Adresse email(Obligatoire): "+email);
             run.addBreak();
             run.addBreak();
             String marital_situation = row.getCell(9).getStringCellValue();
             run.setText("Situation matrimoniale: "+marital_situation);
             run.addBreak();
             run.addBreak();
             String phone = row.getCell(10).getStringCellValue();
             run.setText("Contacts téléphoniques: "+phone);
             run.addBreak();
             run.addBreak();
             run.setText("Ci- après désigné « ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("L’employé(e)");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("»                                                                              d’autre part ;");
             run.addBreak();
             run.addBreak();
             //run.addBreak();
             run.setText("Il a été convenu ce qui suit :");
             run.addBreak(BreakType.PAGE);
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("     II-      ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("LE CONTRAT");
             run.addBreak();
             run.setText("ARTICLE 1:");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Objet du contrat");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("Le travailleur est engagé pour être mis à la disposition du ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Client");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", en qualité de ");
             function = row.getCell(11).getStringCellValue();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText(function);
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", classé en ");
             run.addBreak();
             String category = row.getCell(12).getStringCellValue();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText(category);
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("de la Convention Collective relative aux Industries Extractives et Prospection Minière.");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("ARTICLE 2: ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Durée du contrat et renouvellement");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("Le travailleur est engagé pour la durée des travaux de forage miniers à accomplir par le ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Client");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", à");
             run.addBreak();
             run.setText("compter de la date de mise en service effective de la foreuse.");
             run.addBreak();
             run.setText("Toutefois, l’engagement peut être limité à la durée des travaux spécifiques confiés au travailleur");
             run.addBreak();
             run.setText("La date de la fin des travaux fixée par l’entreprise utilisatrice sera retenue comme celle du terme de ");
             run.addBreak();
             run.setText("la présente convention.");
             run.addBreak();
             run.setText("L’engagement du travailleur est renouvelable. ");
             run.addBreak();
             run.setText("À la fin du contrat, aucun agent ne peut poursuivre sans avoir préalablement signé un nouveau ");
             run.addBreak();
             run.setText("contrat.");
             run.addBreak();
             run.setText("À défaut, si l’agent poursuit son contrat sans en avoir signé un nouveau, la poursuite du contrat ne ");
             run.addBreak();
             run.setText("peut excéder un (1) mois à compter de la date du dernier contrat ou au lendemain de la date ");
             run.addBreak();
             run.setText("d’expiration du dernier contrat.");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("ARTICLE 3 : ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Rémunération ");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("L’employeur s’engage pendant la durée d’exécution du contrat à payer au travailleur");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("un salaire");
             run.addBreak();
             run.setText("mensuel brut imposable ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("(sans heures supplémentaires et sans bonus) d’un montant global de ");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             double salary = row.getCell(13).getNumericCellValue();
             run.setText(salary+" CFA");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("qui se décompose comme suit :");
             run.addBreak();
             run.setText("- Salaire de base: ");
             double baseSalary = row.getCell(14).getNumericCellValue();
             run.setText(baseSalary+" F CFA");
             run.addBreak();
             run.setText("- SurSalaire: ");
             double overSalary = row.getCell(15).getNumericCellValue();
             run.setText(overSalary+" F CFA");
             run.addBreak();
             run.setText("- Prime de panier et de restauration : ");
             run.addBreak();
             run.setText("     o Prime(s) de panier (taux) :  ");
             double basketRate = row.getCell(16).getNumericCellValue();
             run.setText(basketRate+" F CFA");
             run.addBreak();
             run.setText("     o Prime(s) de restauration :  ");
             double restoration = row.getCell(17).getNumericCellValue();
             run.setText(restoration+" F CFA");
             run.addBreak();
             run.setText("- Indemnité de transport : ");
             double transport = row.getCell(18).getNumericCellValue();
             run.setText(transport+" F CFA");
             run.addBreak();
             run.setText("- Prime de salissure : ");
             double smear = row.getCell(19).getNumericCellValue();
             run.setText(smear+" F CFA");
             run.addBreak();
             run.setText("- Prime d'entretien vêtement : ");
             double clothe = row.getCell(20).getNumericCellValue();
             run.setText(clothe+" F CFA");
             run.addBreak();
             run.setText("- Prime de logement : ");
             double housing = row.getCell(21).getNumericCellValue();
             run.setText(housing+" F CFA");
             run.addBreak();
             run.setText("- Prime de chantier : ");
             double construction = row.getCell(22).getNumericCellValue();
             run.setText(construction+" F CFA");
             run.addBreak();
             run.setText("- Prime de responsabilité : ");
             double responsibility = row.getCell(23).getNumericCellValue();
             run.setText(responsibility+" F CFA");
             run.addBreak();
             run.setText("- Prime de représentation : ");
             double representation = row.getCell(24).getNumericCellValue();
             run.setText(representation+" F CFA");
             run.addBreak();
             run.setText("- Avantage 1 : ");
             double advantage1 = row.getCell(25).getNumericCellValue();
             run.setText(advantage1+" F CFA");
             run.addBreak();
             run.setText("- Avantage 2 : ");
             double advantage2 = row.getCell(26).getNumericCellValue();
             run.setText(advantage2+" F CFA");
             run.addBreak();
             run.setText("- Avantage 3 : ");
             double advantage3 = row.getCell(27).getNumericCellValue();
             run.setText(advantage3+" F CFA");
             run.addBreak();
             run.setText("- Avantage 4 : ");
             double advantage4 = row.getCell(28).getNumericCellValue();
             run.setText(advantage4+" F CFA");
             run.addBreak();
             run.setText("La prime de restauration n’est pas due si l’entreprise utilisatrice fourni par l’entremise d’un cuisinier ");
             run.addBreak();
             run.setText("le petit déjeuner, le déjeuner et le diner.");
             run.addBreak();
             run.setText("L’employeur s’engage à payer au travailleur, au terme de son contrat, outre le dernier salaire");
             run.addBreak();
             run.setText("l’indemnité de fin de contrat et l’indemnité compensatrice de congé-payé, dans les conditions ");
             run.addBreak();
             run.setText("déterminées par les articles 15.8 et 25.10 du Code du travail.");
             //run.addBreak(BreakType.PAGE);
             //run.addBreak();
             //run.addBreak();
             //run.addBreak();
             //run.addBreak();
             CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
             XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
             XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
             paragraph = header.createParagraph();
             run = paragraph.createRun();

             if(containsIgnoreCase(function,"SECRETAIRE DE CHANTIER")){
                URL header_img=getClass().getResource("header_construction_secretary.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_construction_secretary.png", Units.toEMU(500), Units.toEMU(100));
                
        
             }
             if(containsIgnoreCase(function,"CUISINIER")){
                URL header_img=getClass().getResource("header_cook.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_cook.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"COOK ASSISTANT")){
                URL header_img=getClass().getResource("header_cook_assistant.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_cook_assistant.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"FOREUR 1")){
                URL header_img=getClass().getResource("header_driller1.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_driller1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"FOREUR 2")){
                URL header_img=getClass().getResource("header_driller2.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_driller2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ELECTRICIEN CATEGORIE 1")){
                URL header_img=getClass().getResource("header_electrician1.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_electrician1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ELECTRICIEN CATEGORIE 2")){
                URL header_img=getClass().getResource("header_electrician2.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_electrician2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"AGENT DE SAISIE")){
                URL header_img=getClass().getResource("header_entry_clerk.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_entry_clerk.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"GESTIONNAIRE DE FLOTTE")){
                URL header_img=getClass().getResource("header_fleet_manager.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_fleet_manager.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"HSE")){
                URL header_img=getClass().getResource("header_hse.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_hse.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ASSISTANT MAINTENANCE PLANNER")){
                URL header_img=getClass().getResource("header_maintenace_assistant_planner.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_maintenace_assitant_planner.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ASSISTANT MAINTENANCE MANAGER")){
                URL header_img=getClass().getResource("header_manager_maintenace_assistant.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_manager_maintenace_assitant.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MECANICIEN 1")){
                URL header_img=getClass().getResource("header_manager_mechanic1.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_mechanic1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MECANICIEN 2")){
                URL header_img=getClass().getResource("header_manager_mechanic2.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_mechanic2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER-CLERK")){
                URL header_img=getClass().getResource("header_offsider_clerk.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_offsider_clerk.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 1A")){
                URL header_img=getClass().getResource("header_offsider1A.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_offsider1A.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 1B")){
                URL header_img=getClass().getResource("header_offsider1B.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_offsider1B.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 2")){
                URL header_img=getClass().getResource("header_offsider2.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_offsider2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"PEINTRE")){
                URL header_img=getClass().getResource("header_painter.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_painter.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MAGASINIER")){
                URL header_img=getClass().getResource("header_storekeeper.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_storekeeper.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"TOLIER-PEINTRE")){
                URL header_img=getClass().getResource("header_tolier_painter.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_tolier_painter.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"SOUDEUR 1")){
                URL header_img=getClass().getResource("header_welder1.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_welder1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"SOUDEUR 2")){
                URL header_img=getClass().getResource("header_welder2.png");
                run.addPicture(header_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header_welder2.png", Units.toEMU(500), Units.toEMU(100));
             }
             XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
             paragraph = footer.createParagraph();
             run = paragraph.createRun();
             if(containsIgnoreCase(function,"SECRETAIRE DE CHANTIER")){
                URL footer_img=getClass().getResource("footer_construction_secretary.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_construction_secretary.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"CUISINIER")){
                URL footer_img=getClass().getResource("footer_cook.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_cook.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"COOK ASSISTANT")){
                URL footer_img=getClass().getResource("footer_cook_assistant.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_cook_assistant.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"FOREUR 1")){
                URL footer_img=getClass().getResource("footer_driller1.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_driller1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"FOREUR 2")){
                URL footer_img=getClass().getResource("footer_driller2.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_driller2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ELECTRICIEN CATEGORIE 1")){
                URL footer_img=getClass().getResource("footer_electrician1.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_electrician1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ELECTRICIEN CATEGORIE 2")){
                URL footer_img=getClass().getResource("footer_electrician2.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_electrician2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"AGENT DE SAISIE")){
                URL footer_img=getClass().getResource("footer_entry_clerk.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_entry_clerk.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"GESTIONNAIRE DE FLOTTE")){
                URL footer_img=getClass().getResource("footer_fleet_manager.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_fleet_manager.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"HSE")){
                URL footer_img=getClass().getResource("footer_hse.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_hse.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ASSISTANT MAINTENANCE PLANNER")){
                URL footer_img=getClass().getResource("footer_maintenace_assistant_planner.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_maintenace_assitant_planner.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"ASSISTANT MAINTENANCE MANAGER")){
                URL footer_img=getClass().getResource("footer_manager_maintenace_assistant.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_manager_maintenace_assitant.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MECANICIEN 1")){
                URL footer_img=getClass().getResource("footer_mechanic1.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_mechanic1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MECANICIEN 2")){
                URL footer_img=getClass().getResource("footer_mechanic2.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_mechanic2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER-CLERK")){
                URL footer_img=getClass().getResource("footer_offsider_clerk.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_offsider_clerk.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 1A")){
                URL footer_img=getClass().getResource("footer_offsider1A.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_offsider1A.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 1B")){
                URL footer_img=getClass().getResource("footer_offsider1B.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_offsider1B.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"OFFSIDER 2")){
                URL footer_img=getClass().getResource("footer_offsider2.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_offsider2.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"PEINTRE")){
                URL footer_img=getClass().getResource("footer_painter.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_painter.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"MAGASINIER")){
                URL footer_img=getClass().getResource("footer_storekeeper.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_storekeeper.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"TOLIER-PEINTRE")){
                URL footer_img=getClass().getResource("footer_tolier_painter.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_tolier_painter.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"SOUDEUR 1")){
                URL footer_img=getClass().getResource("footer_welder1.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_welder1.png", Units.toEMU(500), Units.toEMU(100));
             }
             if(containsIgnoreCase(function,"SOUDEUR 2")){
                URL footer_img=getClass().getResource("footer_welder2.png");
                run.addPicture(footer_img.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer_welder2.png", Units.toEMU(500), Units.toEMU(100));
             }
            document.write(out);
            out.close();
                    
      }
            
             /*run.addBreak();
             URL logo=getClass().getResource("logo.png");
             run.addPicture(logo.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "logo.png", Units.toEMU(50), Units.toEMU(50));*/
           // document.write(out);
             //out.close();
             //exit();
     
      JOptionPane.showMessageDialog(null, "Remplissage terminé"); 
    }
     public static boolean containsIgnoreCase(String str, String subString) {
        return str.toLowerCase().contains(subString.toLowerCase());
    }   
    }
