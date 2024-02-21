package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.UnaryOperator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.shape.Circle;
import javafx.scene.shape.Rectangle;
import javafx.scene.control.Button;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.Label;
import javafx.scene.control.MenuButton;
import javafx.scene.control.MenuItem;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.control.TextFormatter.Change;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.image.WritableImage;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.Background;
import javafx.scene.layout.BackgroundFill;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import javafx.stage.FileChooser.ExtensionFilter;

public class mainController {
	@FXML
	private Text inputFileAddress;
	@FXML
	private Text dateShow;
	@FXML
	private Text outputFileAddress;
	@FXML
	private Button startMondify;
	@FXML
	private TextField companyName;
	@FXML
	private TextField companyAddress;
	@FXML
	private TextField managerName;
	@FXML
	private TextField platformName;
	@FXML
	private TextField jobPosition;
	
	private String needCompanyName = null;
	private String needCompanyAddress = null;
	private String needmanagerName = null;
	private String needplatformName = null;
	private String needjobPosition = null;
	private String date = null;
	
	private String originalFileileLocation = null;
	private String outputFileileLocation = null;
	
	@FXML
	public void initialize() {
		//getParameters();
		platformName.setText("Seek");
		managerName.setText("HR Team");
		
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd MMM yyyy");  
		LocalDateTime now = LocalDateTime.now(); 
		date = dtf.format(now);
		
		System.out.println(date);  
		
		originalFileileLocation = "/Users/liuzheng/OneDrive - LZ-HOME/MY CV/Zheng's Cover Letter Tamplate.docx";
		
		inputFileAddress.setText("From: " + originalFileileLocation);
		outputFileAddress.setText("To: " + outputFileileLocation);
		}
	
	@FXML	
    private void startMondify(ActionEvent event) throws InvalidFormatException
    {
		getParameters();
		try {
			modifyTheCV();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	public void getParameters() {
		needCompanyName = companyName.getText();
		needCompanyAddress = companyAddress.getText();
		needmanagerName = managerName.getText();
		needplatformName = platformName.getText();
		needjobPosition = jobPosition.getText();
		
		System.out.println(needCompanyName);
		System.out.println(needmanagerName);
		System.out.println(needplatformName);
		System.out.println(needjobPosition);
		
		outputFileileLocation = "/Users/liuzheng/Desktop/Zheng's Cover Letter for " + needCompanyName + ".docx";
	}
	
	public void modifyTheCV() throws IOException,org.apache.poi.openxml4j.exceptions.InvalidFormatException {
	      try {

	          /**
	           * if uploaded doc then use HWPF else if uploaded Docx file use
	           * XWPFDocument
	           */
	          XWPFDocument doc = new XWPFDocument(
	            OPCPackage.open(originalFileileLocation));
	          for (XWPFParagraph p : doc.getParagraphs()) {
	           List<XWPFRun> runs = p.getRuns();
	           if (runs != null) {
	            for (XWPFRun r : runs) {
	             String text = r.getText(0);
	             if (text != null && text.contains("<Date>")) {
		              text = text.replace("<Date>", date);//your content
		              r.setText(text, 0);
		             }
	             
	             if (text != null && text.contains("<Hiring manager’s name>")) {
	              text = text.replace("<Hiring manager’s name>", needmanagerName);//your content
	              r.setText(text, 0);
	             }
	             
	             if (text != null && text.contains("<Company>")) {
		              text = text.replace("<Company>", needCompanyName);//your content
		              r.setText(text, 0);
		             }
	             
	             if (text != null && text.contains("<Company address>")) {
		              text = text.replace("<Company address>", needCompanyAddress);//your content
		              r.setText(text, 0);
		             }
	             
	             if (text != null && text.contains("<the name of the position>")) {
		              text = text.replace("<the name of the position>", needjobPosition);//your content
		              r.setText(text, 0);
		             }
	             
	             if (text != null && text.contains("<Platform name>")) {
		              text = text.replace("<Platform name>", needplatformName);//your content
		              r.setText(text, 0);
		             }
	            }
	           }
	          }

	          for (XWPFTable tbl : doc.getTables()) {
	           for (XWPFTableRow row : tbl.getRows()) {
	            for (XWPFTableCell cell : row.getTableCells()) {
	             for (XWPFParagraph p : cell.getParagraphs()) {
	              for (XWPFRun r : p.getRuns()) {
	               String text = r.getText(0);
		            if (text != null && text.contains("<Date>")) {
			              text = text.replace("<Date>", date);//your content
			              r.setText(text, 0);
			             }
	               
	               if (text != null && text.contains("<Hiring manager’s name>")) {
	                   text = text.replace("<Hiring manager’s name>", needmanagerName);   
	                r.setText(text, 0);
	               }
	               
		             if (text != null && text.contains("<Company>")) {
			              text = text.replace("<Company>", needCompanyName);//your content
			              r.setText(text, 0);
			             }
		             
		             if (text != null && text.contains("<Company address>")) {
			              text = text.replace("<Company address>", needCompanyAddress);//your content
			              r.setText(text, 0);
			             }
		             
		             if (text != null && text.contains("<the name of the position>")) {
			              text = text.replace("<the name of the position>", needjobPosition);//your content
			              r.setText(text, 0);
			             }
		             
		             if (text != null && text.contains("<Platform name>")) {
			              text = text.replace("<Platform name>", needplatformName);//your content
			              r.setText(text, 0);
			             } 
	              }
	             }
	            }
	           }
	          }

	          doc.write(new FileOutputStream(outputFileileLocation));
	         } finally {
	        	 dateShow.setText(date);
	        	 
	     		inputFileAddress.setText("From: " + originalFileileLocation);
	    		outputFileAddress.setText("To: " + outputFileileLocation);
	         }
		
	}
	
}
