package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.Optional;
import java.util.concurrent.CountDownLatch;

import javax.swing.JOptionPane;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import com.google.auth.oauth2.GoogleCredentials;
import com.google.firebase.FirebaseApp;
import com.google.firebase.FirebaseOptions;
import com.google.firebase.database.DataSnapshot;
import com.google.firebase.database.DatabaseError;
import com.google.firebase.database.DatabaseReference;
import com.google.firebase.database.FirebaseDatabase;
import com.google.firebase.database.ValueEventListener;

import application.DashBoardController.Person;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Node;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.TableColumn.CellEditEvent;
import javafx.scene.control.TableRow;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.util.Pair;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
public class Attendance 
{
			


   String sub;
   String[] subs;
 //  String res;
   	String dur = "";
	String strno = "";
	String strcost = "";
	
	Boolean exp = false;

	@FXML
	TitledPane tpatt;

	@FXML
	ListView<String> list=new ListView<String>();
	@FXML
	Label in1;
	@FXML
	TextField p_num=new TextField();
	@FXML
    TextField p_name=new TextField();
	@FXML
    TextField pcost=new TextField();
	@FXML
	Label lb;
	@FXML 
	AnchorPane ap, ap_attendance;
	@FXML
	HBox hb;
	@FXML
	
	TextField subject=new TextField();
	@FXML
	VBox utility_attend;
	String rootpath = "C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\";
	String coursecode ="";
	@FXML
	private TableView table = new TableView();
	TableColumn pnocol1,nameCol1,costcol;
	static ArrayList<String> studdat = new ArrayList<String>();
	ObservableList<String> abc=  FXCollections.observableArrayList();
	
	
	  
	    
	    private ObservableList data =
		        FXCollections.observableArrayList();
	    
	  
	    
	    @FXML
		Button add,delete, search;
	    
	    public void initialize() throws IOException
		{		
	    	tpatt.setExpanded(true);
	    	//in1.setText(in);
	    	in1.setWrapText(true);
	        in1.setStyle("-fx-font-family: \"Times New Roman \"; -fx-font-size: 20; -fx-text-fill: white; -fx-background-color:#6464a5;");
	    	 Rectangle2D screenBounds = Screen.getPrimary().getBounds();
			 int width = (int) screenBounds.getWidth();
		        int utilsize = (width/100)*20;
		        int lab_size = (width/100)*60;
		        utility_attend.setPrefWidth(utilsize);
		        table.setPrefWidth(lab_size);
		        table.setPrefHeight(screenBounds.getHeight());
		        
		        add.setPrefWidth(utilsize);
		        delete.setPrefWidth(utilsize);
		        search.setPrefWidth(utilsize);
		       // syncsave.setPrefWidth(utilsize);
		      //  addTotalClasses.setPrefWidth(utilsize);
		        
		        
		        
		      
		        table.setEditable(true);
		    	
				pnocol1 = new TableColumn("P.No");
		        pnocol1.setMinWidth(150);
		        pnocol1.setCellValueFactory(
		                new PropertyValueFactory("P.No"));
		       
		        nameCol1 = new TableColumn("NAME");
		        nameCol1.setMinWidth(250);
		        nameCol1.setCellValueFactory(
		                new PropertyValueFactory("name"));
		 
		        costcol = new TableColumn("Cost");
		        costcol.setMinWidth(200);
		        costcol.setCellValueFactory(
		                new PropertyValueFactory("cost"));
		       // costcol.setCellFactory(TextFieldTableCell.forTableColumn());
		        
		        		table.setEditable(true);
		        		 
		        	costcol.setCellFactory(TextFieldTableCell.forTableColumn());
		        		/*costcol.setOnEditCommit(new EventHandler<CellEditEvent>() {
		        		                 
		        		    @Override
		        		    public void handle(CellEditEvent t) {
		        		                     
		        		        ( t.getTableView().getItems().get(
		        		            t.getTablePosition().getRow())).setTitle(t.getNewValue());
		        		    }
		        		});
		        		*/
		        
		       
		        table.setItems(data);
		        table.getColumns().addAll(pnocol1,nameCol1, costcol);
		       
		    	
		    	
		    	
		    	 /*	pnocol1.setVisible(false);
		    	nameCol1.setVisible(false);
		    	costcol.setVisible(false);
		    	list.setVisible(false);
		    	list.setItems(abc);
		    	*/
		    	
		    
		    	
		       }
	    
	    public void openDetails(ActionEvent e)
	    {
	    	if(!exp==true)
	    	{
	    	exp =true;
	    	tpatt.setExpanded(true);
	    	}
	    	else
	    	{
	    		exp=false;
	    		tpatt.setExpanded(false);
	    	}
	    }
	    
	    public void SaveFirebaseAttendance(ActionEvent e) throws IOException
		{
	    	coursecode=subject.getText().toString().toUpperCase();
			//int tc = Integer.parseInt(addTotalClasses.getText().toString());
	   	 			
					
			
		    try {
		    	
		    	ArrayList<ArrayList<String>> big = new ArrayList<ArrayList<String>>();
		      	
		      	
		      	
		      	
		            final CountDownLatch latch1 = new CountDownLatch(1);
		            DatabaseReference ref1= FirebaseDatabase.getInstance().getReference();
		            DatabaseReference ref2;    
		             ref2 = ref1.child("Products");

		        	 String tchr_name = coursecode;
		        	
		        	
		        	// String att = String.join(",", attend);
		        	// String perc = String.join(",", percent);
		        	// DatabaseReference ref = FirebaseDatabase.getInstance().getReference("Products");
		        	 	 DatabaseReference child_name = FirebaseDatabase.getInstance().getReference();
		        	 	
		        	 child_name=ref2.child("pnumber");
		        	 child_name.setValueAsync(p_num.getText());
		        	 child_name=ref2.child(" name");
		        	 child_name.setValueAsync(p_name);
		        	 child_name=ref2.child("cost");
		        	 child_name.setValueAsync(pcost.getText());
		        	 latch1.countDown();
		        	 
		        	System.out.println("Succesfull");
		        	 
		        	latch1.await();
		    			   } 
		    			 catch (InterruptedException ef) {
		    			        ef.printStackTrace();
		    			    }
		    Alert alerts=new Alert(AlertType.INFORMATION);
	        alerts.setTitle("Information Dialog");
	        alerts.setHeaderText(null);
	        alerts.setContentText("Saved Online!");
	        alerts.showAndWait();
	        return;
				
		}
		
		
		
		
		public void LoadFirebaseAttendance(ActionEvent e) throws IOException
		{
					
				

		     ArrayList<DataSnapshot> Userlist = new ArrayList<DataSnapshot>();

		      try {
		              final CountDownLatch latch1 = new CountDownLatch(1);
		              DatabaseReference ref1= FirebaseDatabase.getInstance().getReference();
		              DatabaseReference ref2;    
		               ref2 = ref1.child("Products");


		               ref2.addListenerForSingleValueEvent(
		            		   new ValueEventListener() {
		                public void onDataChange(DataSnapshot dataSnapshot) {

		                    //ArrayList<Object> Userlist = new ArrayList<Object>();   
		                    ArrayList<ArrayList<String>> big_arr = new ArrayList<ArrayList<String>>();
		                   	   	for (DataSnapshot dsp : dataSnapshot.getChildren()) {
		                   	   		System.out.println(dsp.getKey());
		                   	      if(dsp.getKey().equalsIgnoreCase(strcost))  
		                   	    	  Userlist.add(dsp); 
		                         
		                        }
		                   	                                                latch1.countDown();
		     				     }

		    			        	  public void onCancelled(DatabaseError error) {
		    			        		  latch1.countDown();
		    			        		  
		    			        	  }
		    			        	});
		    			        	 latch1.await();
		     			   } 
		     			 catch (InterruptedException en) {
		    				        en.printStackTrace();
		    				    }
		     			
		      	ArrayList<ArrayList<String>> big = new ArrayList<ArrayList<String>>();
		      	ArrayList<String> smol = new ArrayList<String>();
		      	
		      	for(DataSnapshot d: Userlist.get(0).getChildren())
		      	{
		      		
		      		smol = new ArrayList<String>();
		      		smol.add(d.getValue().toString());
		      		big.add(smol);
		      		
		      	
		      		
		      	
		      	}
		    	
		      	System.out.println(big);
		      	//combine(big);
		      	 Alert alerts=new Alert(AlertType.INFORMATION);
			        alerts.setTitle("Information Dialog");
			        alerts.setHeaderText(null);
			        alerts.setContentText("blaghj");
			        alerts.showAndWait();
		
		}	
		
		public void saveAttendance(ActionEvent e) throws IOException
		{
		 

		    studdat.clear();
		    strno = p_num.getText().toString();
		    strno = strno.toUpperCase();
		    strcost = pcost.getText().toString();
		    strcost = strcost.toUpperCase();
		    studdat.add(strno);
		    studdat.add(strcost);
		   
	    
	    	  	  String sub=subject.getText().toString().toUpperCase();
		        System.out.println(sub);
		       		    	
	 	/*  InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+".xls");
			HSSFWorkbook  workbook = new HSSFWorkbook(ExcelFileToRead);
	        HSSFSheet spreadsheet = workbook.getSheetAt(0);
	 	   /*Workbook workbook = new HSSFWorkbook();
	        Sheet spreadsheet = workbook.createSheet("sample");

	        Row row = spreadsheet.createRow(0);
	      
	      spreadsheet.getRow(0).createCell(0).setCellValue("DAYANANDA SAGAR COLLEGE OF ENGINEERING");
	        spreadsheet.getRow(1).createCell(0).setCellValue("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING");
	        spreadsheet.getRow(2).createCell(0).setCellValue(studdat.get(0)+studdat.get(1)+" ATTENDANCE: "+finalDate);
	        spreadsheet.getRow(3).createCell(0).setCellValue("SUBJECT: "+sub);
	        spreadsheet.addMergedRegionUnsafe(new CellRangeAddress(0, 0, 0, 5));
	        spreadsheet.addMergedRegionUnsafe(new CellRangeAddress(1, 1, 0, 5));
	        spreadsheet.addMergedRegionUnsafe(new CellRangeAddress(2, 2, 0, 5));
	        spreadsheet.addMergedRegionUnsafe(new CellRangeAddress(3, 3, 0, 5));
	        System.out.println(table.getItems().size());
	        
	        int i=5;
	        for(Person dsce: data)
			{
				if(!dsce.getClasses().equals(""))
				{
					System.out.println(spreadsheet.getRow(i).getCell(0).getStringCellValue());
					spreadsheet.getRow(i).createCell(2).setCellValue(dsce.getClasses());
				System.out.println(dsce.getClasses());
				Double percentage = Double.parseDouble((dsce.getClasses().toString()))/tc;
				percentage= percentage *100;
				
				int perc = (int) Math.round(percentage);
				
				spreadsheet.getRow(i).createCell(3).setCellValue(perc+"");
				}
				i++;
				
			}
	        
	      
	        
	        String directoryName=rootpath+studdat.get(0)+studdat.get(1)+"-"+sub;
	        File directory = new File(directoryName);
	        System.out.println(directoryName);
	        if (! directory.exists()){
	            directory.mkdir();
	        }
	        String fileName="\\"+finalDate+".xls";
	        */
		
		        HSSFWorkbook  workbook = new HSSFWorkbook();
		        HSSFSheet sheet = workbook.getSheetAt(0);
	        
		 String directoryName="wholesale directory";
	        File directory = new File(directoryName);
	        String fileName="\\"+"prod details"+".xls";
	  
	        FileOutputStream fileOut = new FileOutputStream(directoryName+fileName);
	        workbook.write(fileOut);
	        fileOut.close();
	       	       // InputStream ExcelFileToRead1 = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+"-"+sub+"\\"+finalDate+".xls");
			
			
			
			HSSFRow row; 
			HSSFCell cell;

			Iterator rows = sheet.rowIterator();
			
			//data.clear();
			int k = 5;
			System.out.println(sheet.getPhysicalNumberOfRows());
			while(k<sheet.getPhysicalNumberOfRows())
			{
				data.add(sheet.getRow(k).getCell(0).getStringCellValue());
				//System.out.println(sheet.getRow(k).getCell(2).getStringCellValue());
				k++;
			}
			 table.setItems(data);
			/* int i1=0;
			 for (Node n: table.lookupAll("TableRow")) {
			      if (n instanceof TableRow) {
			        TableRow row1 = (TableRow) n;
			        if (Integer.parseInt(table.getItems().get(i1).getPer().toString().trim())<75) {
			          row1.getStyleClass().add("red");
			          //row1.setDisable(false);
			        } else if(Integer.parseInt(table.getItems().get(i1).getPer().toString().trim())>=75 && Integer.parseInt(table.getItems().get(i1).getPer().toString().trim())<85){
			          row1.getStyleClass().add("yellow");
			          //row1.setDisable(true);
			        }
			        else if(Integer.parseInt(table.getItems().get(i1).getPer().toString().trim())>=85)
			        {
			        row1.getStyleClass().add("green");
			        }
			        i1++;
			        if (i1 == table.getItems().size())
			          break;
			      }
			    }*/
			 
			
	        Alert alert=new Alert(AlertType.INFORMATION);
	        alert.setTitle("Information Dialog");
	        alert.setHeaderText(null);
	        alert.setContentText("Spreadsheet Saved!");
	        alert.showAndWait();
	        }

	protected void setText(Object object) {
			// TODO Auto-generated method stub
			
		}

	protected void setStyle(String string) {
			// TODO Auto-generated method stub
			
		}

	public void loadAttendance(ActionEvent e)throws IOException
	{
		p_num.setVisible(true);
    	p_name.setVisible(true);
    	pcost.setVisible(true);
    	//perCol.setVisible(false);
    	list.setVisible(true);
    	list.setEditable(true);
    	pnocol1.setVisible(true);
    	nameCol1.setVisible(true);
    	costcol.setVisible(true);
    	
    	
	    studdat.clear();
	    strno = p_num.getText().toString();
	    strno = strno.toUpperCase();
	    strcost = pcost.getText().toString();
	    strcost = strcost.toUpperCase();
	    studdat.add(strno);
	    studdat.add(strcost);
	    String sub=subject.getText().toString();
	   
	    if(strno.equals("")||strcost.equals("")||sub.equals(""))
		{
			Alert alerts=new Alert(AlertType.WARNING);
	        alerts.setTitle("Warning Dialog");
	        alerts.setHeaderText(null);
	        alerts.setContentText("Kindly enter all the text fields!");
	        alerts.showAndWait();
	        return;
	        }
		table.getColumns().clear();
		table.getColumns().addAll(pnocol1, nameCol1, costcol);
		table.setItems(data);
		
		
		
	    
		
			data.clear();
			table.setItems(data);
		
			
			table.setItems(data);
	}
	
	/*public void importAttendanceFile(ActionEvent e) throws IOException
	{
		p_num.setVisible(true);
    	p_name.setVisible(true);
    	pcost.setVisible(true);
    	//perCol.setVisible(false);
    	list.setVisible(true);
    	list.setEditable(true);
    	pnocol1.setVisible(true);
    	nameCol1.setVisible(true);
    	costcol.setVisible(true);
    	
    	
	    studdat.clear();
	    strno = p_num.getText().toString();
	    strno = strno.toUpperCase();
	    strcost = pcost.getText().toString();
	    strcost = strcost.toUpperCase();
	    studdat.add(strno);
	    studdat.add(strcost);
	    String sub=subject.getText().toString();
	    
		table.getColumns().clear();
		table.getColumns().addAll(pnocol1, nameCol1, costcol);
		table.setItems(data);
		
		
		
	    
		
		data.clear();
		table.setItems(data);
		
		
			table.setItems(data);
	}

	/*public void combine( ArrayList<ArrayList<String>> big) throws IOException
	{
		 String directoryName=rootpath+"Consolidated";
	     File directory = new File(directoryName);
	    // System.out.println(directoryName);
	     if (! directory.exists()){
	         directory.mkdir();
	        }
	     String directoryName1=rootpath+"Consolidated\\"+studdat.get(0)+studdat.get(1);
	     File directory1 = new File(directoryName1);
	    // System.out.println(directoryName);
	     if (! directory1.exists()){
	         directory1.mkdir();
	        }
	     
		 int n;
		 XWPFDocument docX2 = new XWPFDocument();
		 
		 CTBody body = docX2.getDocument().getBody();
		 if(!body.isSetSectPr()){
			 body.addNewSectPr();
			 }
			  
			 CTSectPr pcost = body.getSectPr();
			 if(!pcost.isSetPgSz()){
			 pcost.addNewPgSz();
			 }
			  
			 CTPageSz pageSize = pcost.getPgSz();
			 pageSize.setOrient(STPageOrientation.LANDSCAPE);
			 //A4 = 595x842 / multiply 20 since BigInteger represents 1/20 Point
			 pageSize.setW(BigInteger.valueOf(16840));
			 pageSize.setH(BigInteger.valueOf(11900));
	        
	      XWPFParagraph paragraph = docX2.createParagraph();
	      paragraph.setAlignment(ParagraphAlignment.CENTER);	      
	      XWPFRun paragraphOneRunOne = paragraph.createRun();
	      paragraphOneRunOne.setBold(true);
	      paragraphOneRunOne.setText("DAYANANDA SAGAR COLLEGE OF ENGINEERING");
	      paragraphOneRunOne.addBreak();
	      
	     
	      XWPFRun paragraphOneRunTwo = paragraph.createRun();
	      paragraphOneRunTwo.setBold(true);
	      paragraphOneRunTwo.setText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING");
	      paragraphOneRunTwo.addBreak();
	      paragraphOneRunTwo.addBreak();
	      
	      
	      XWPFRun paragraphOneRunThree = paragraph.createRun();
	      paragraphOneRunThree.setBold(true);
	      paragraphOneRunThree.setText("ATTENDANCE DISPLAY");
	      paragraphOneRunThree.addBreak();
	      
	      XWPFRun paragraphOneRunFour = paragraph.createRun();
	      paragraphOneRunFour.setBold(true);
	      paragraphOneRunFour.setText("(Session: "+dur+")");
	      paragraphOneRunFour.addBreak();
	      paragraphOneRunFour.addBreak();
	      paragraphOneRunFour.addBreak();
	      
	      XWPFParagraph paragraph1 = docX2.createParagraph();
	      paragraph1.setAlignment(ParagraphAlignment.LEFT);
	      String finalDate="";
			LocalDate date = datePicker1.getValue();
	       DateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
	        Date conv_date = java.sql.Date.valueOf(date);
	        finalDate = formatter.format(conv_date);
	        finalDate = finalDate.replace('/', '-');
	     
	      
	      XWPFRun paragraphTwoRunOne = paragraph1.createRun();
	      paragraphTwoRunOne.setBold(true);
	      paragraphTwoRunOne.setText("Class: "+strno+strcost+"                                                                                              Cumulative Attendance Record: "+finalDate);
	      
	     
	      
	      
	      //create table
	      XWPFTable table = docX2.createTable();
	      
	      
	      
	      
	      XWPFTableRow tableRowOne = table.createRow();
	      table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(1000));
	      
	      XWPFTableCell cell2=tableRowOne.getCell(0);
	  	cell2.setText("Sl#");
	  	CTTcPr tcpr = cell2.getCTTc().addNewTcPr();
	  	CTVMerge vMerge=tcpr.addNewVMerge();
	  	vMerge.setVal(STMerge.RESTART); 
	      
	  	XWPFTableCell cell3=tableRowOne.createCell();
	  	cell3.setText("USN");
	  	CTTcPr tcpr1 = cell3.getCTTc().addNewTcPr();
	  	CTVMerge vMerge1=tcpr1.addNewVMerge();
	  	vMerge1.setVal(STMerge.RESTART); 
	      
	  	
	  	XWPFTableCell c2 = tableRowOne.createCell();
	  	XWPFRun run = c2.addParagraph().createRun();
	  	run.setBold(true);run.setText("Subject ->");run.setFontSize(12);
	  	c2.removeParagraph(0);
	  	
	  	try {
            final CountDownLatch latch1 = new CountDownLatch(1);
            DatabaseReference ref= FirebaseDatabase.getInstance().getReference().child("Subjects/");


             ref.addListenerForSingleValueEvent(
          		new ValueEventListener() {
	              public void onDataChange(DataSnapshot d) {
	            	  if(d.hasChild(strno))
	            	  {
	            		  System.out.println(d.child(strno).getValue().toString());
	            		  subs = d.child(strno).getValue().toString().split(",");
	            		  
	            	  }
	            	  else
	            	  {
	            		  Alert alerts=new Alert(AlertType.WARNING);
	  			        alerts.setTitle("Warning Dialog");
	  			        alerts.setHeaderText(null);
	  			        alerts.setContentText("Kindly enter the subjects for this p_num in Student Setup!");
	  			        alerts.showAndWait();
	  			        
	            	  }
	                  latch1.countDown();
	   				}
	
	  			  public void onCancelled(DatabaseError error) {
	  			      latch1.countDown();
	  			        		  
	  			  }
  			  });
  			  latch1.await();
   			} 
   			catch (InterruptedException en) {
  				en.printStackTrace();
  			}
	  	
	  	XWPFRun run1;
	  	for(int x=0; x<subs.length; x++)
	  	{
	  		XWPFTableCell cell4=tableRowOne.createCell();
		  	run1 = cell4.addParagraph().createRun();
		  	run1.setBold(true);run1.setText(subs[x]);run1.setFontSize(12);
		  	cell4.removeParagraph(0);
		  	CTTcPr tcpr2 = cell4.getCTTc().addNewTcPr();
		  	CTHMerge vMerge2=tcpr2.addNewHMerge();
		  	vMerge2.setVal(STMerge.RESTART); 
		  	
		  	XWPFTableCell cell5=tableRowOne.createCell();
		  	CTTcPr tcpr3 = cell5.getCTTc().addNewTcPr();
		  	CTHMerge vMerge3=tcpr3.addNewHMerge();
		  	vMerge3.setVal(STMerge.CONTINUE);
	  	}
	  	
	  	/*XWPFTableCell cell4=tableRowOne.createCell();
	  	XWPFRun run1 = cell4.addParagraph().createRun();
	  	run1.setBold(true);run1.setText(subs[0]);run1.setFontSize(12);
	  	cell4.removeParagraph(0);
	  	CTTcPr tcpr2 = cell4.getCTTc().addNewTcPr();
	  	CTHMerge vMerge2=tcpr2.addNewHMerge();
	  	vMerge2.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell5=tableRowOne.createCell();
	  	CTTcPr tcpr3 = cell5.getCTTc().addNewTcPr();
	  	CTHMerge vMerge3=tcpr3.addNewHMerge();
	  	vMerge3.setVal(STMerge.CONTINUE); 
	  	
	  	
	  	XWPFTableCell cell6=tableRowOne.createCell();
	  	XWPFRun run2 = cell6.addParagraph().createRun();
	  	run2.setBold(true);run2.setText(subs[1]);run2.setFontSize(12);
	  	cell6.removeParagraph(0);
	  	CTTcPr tcpr4 = cell6.getCTTc().addNewTcPr();
	  	CTHMerge vMerge4=tcpr4.addNewHMerge();
	  	vMerge4.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell7=tableRowOne.createCell();
	  	CTTcPr tcpr5 = cell7.getCTTc().addNewTcPr();
	  	CTHMerge vMerge5=tcpr5.addNewHMerge();
	  	vMerge5.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell8=tableRowOne.createCell();
	  	XWPFRun run3 = cell8.addParagraph().createRun();
	  	run3.setBold(true);run3.setText(subs[2]);run3.setFontSize(12);
	  	cell8.removeParagraph(0);
	  	CTTcPr tcpr6 = cell8.getCTTc().addNewTcPr();
	  	CTHMerge vMerge6=tcpr6.addNewHMerge();
	  	vMerge6.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell9=tableRowOne.createCell();
	  	CTTcPr tcpr7 = cell9.getCTTc().addNewTcPr();
	  	CTHMerge vMerge7=tcpr7.addNewHMerge();
	  	vMerge7.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell10=tableRowOne.createCell();
	  	XWPFRun run4 = cell10.addParagraph().createRun();
	  	run4.setBold(true);run4.setText(subs[3]);run4.setFontSize(12);
	  	cell10.removeParagraph(0);
	  	CTTcPr tcpr8 = cell10.getCTTc().addNewTcPr();
	  	CTHMerge vMerge8=tcpr8.addNewHMerge();
	  	vMerge8.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell11=tableRowOne.createCell();
	  	CTTcPr tcpr9 = cell11.getCTTc().addNewTcPr();
	  	CTHMerge vMerge9=tcpr9.addNewHMerge();
	  	vMerge9.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell12=tableRowOne.createCell();
	  	XWPFRun run5 = cell12.addParagraph().createRun();
	  	run5.setBold(true);run5.setText(subs[4]);run5.setFontSize(12);
	  	cell12.removeParagraph(0);
	  	CTTcPr tcpr10 = cell12.getCTTc().addNewTcPr();
	  	CTHMerge vMerge10=tcpr10.addNewHMerge();
	  	vMerge10.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell13=tableRowOne.createCell();
	  	CTTcPr tcpr11 = cell13.getCTTc().addNewTcPr();
	  	CTHMerge vMerge11=tcpr11.addNewHMerge();
	  	vMerge11.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell14=tableRowOne.createCell();
	  	XWPFRun run6 = cell14.addParagraph().createRun();
	  	run6.setBold(true);run6.setText(subs[5]);run6.setFontSize(12);
	  	cell14.removeParagraph(0);
	  	CTTcPr tcpr12 = cell14.getCTTc().addNewTcPr();
	  	CTHMerge vMerge12=tcpr12.addNewHMerge();
	  	vMerge12.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell15=tableRowOne.createCell();
	  	CTTcPr tcpr13 = cell15.getCTTc().addNewTcPr();
	  	CTHMerge vMerge13=tcpr13.addNewHMerge();
	  	vMerge13.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell16=tableRowOne.createCell();
	  	XWPFRun run7 = cell16.addParagraph().createRun();
	  	run7.setBold(true);run7.setText(subs[6]);run7.setFontSize(12);
	  	cell16.removeParagraph(0);
	  	CTTcPr tcpr14 = cell16.getCTTc().addNewTcPr();
	  	CTHMerge vMerge14=tcpr14.addNewHMerge();
	  	vMerge14.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell17=tableRowOne.createCell();
	  	CTTcPr tcpr15 = cell17.getCTTc().addNewTcPr();
	  	CTHMerge vMerge15=tcpr15.addNewHMerge();
	  	vMerge15.setVal(STMerge.CONTINUE);
	  	
	  	XWPFTableCell cell18=tableRowOne.createCell();
	  	XWPFRun run8 = cell18.addParagraph().createRun();
	  	run8.setBold(true);run8.setText(subs[7]);run8.setFontSize(12);
	  	cell18.removeParagraph(0);
	  	CTTcPr tcpr16 = cell18.getCTTc().addNewTcPr();
	  	CTHMerge vMerge16=tcpr16.addNewHMerge();
	  	vMerge16.setVal(STMerge.RESTART); 
	  	
	  	XWPFTableCell cell19=tableRowOne.createCell();
	  	CTTcPr tcpr17 = cell19.getCTTc().addNewTcPr();
	  	CTHMerge vMerge17=tcpr17.addNewHMerge();
	  	vMerge17.setVal(STMerge.CONTINUE);
	  	
	  	
	  	XWPFTableRow tableRowOne1 = table.createRow();
	  	int twipsPerInch =  1440;
	  	tableRowOne1.setHeight((int)(twipsPerInch*2/10)); //set height 1/10 inch.
	  	tableRowOne1.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT); //set w:hRule="exact"

	  	
	  	
	  	XWPFTableCell cell20=tableRowOne1.getCell(0);
	  	CTTcPr tcpr18 = cell20.getCTTc().addNewTcPr();
	  	CTVMerge vMerge18=tcpr18.addNewVMerge();
	  	vMerge18.setVal(STMerge.CONTINUE); 
	  	
	  	XWPFTableCell cell21=tableRowOne1.createCell();
	  	CTTcPr tcpr19 = cell21.getCTTc().addNewTcPr();
	  	CTVMerge vMerge19=tcpr19.addNewVMerge();
	  	vMerge19.setVal(STMerge.CONTINUE); 
	  	
	  	
	  	XWPFTableCell cel = tableRowOne1.createCell();
	  	//cel.setText("Classes Conducted ->");
	  	XWPFRun run9 = cel.addParagraph().createRun();
	  	run9.setBold(true);run9.setText(" ");run9.setFontSize(9);
	  	cel.removeParagraph(0);
	  	
        
	  	for(int i = 0;i<subs.length;i++)
	  	{
	  		XWPFTableCell cell22=tableRowOne1.createCell();
		  	cell22.setText(" ");
		  	CTTcPr tcpr20 = cell22.getCTTc().addNewTcPr();
		  	CTHMerge vMerge20=tcpr20.addNewHMerge();
		  	vMerge20.setVal(STMerge.RESTART); 
		  	
		  	XWPFTableCell cell23=tableRowOne1.createCell();
		  	CTTcPr tcpr21 = cell23.getCTTc().addNewTcPr();
		  	CTHMerge vMerge21=tcpr21.addNewHMerge();
		  	vMerge21.setVal(STMerge.CONTINUE);
	  	}
	  		
	  	
	  	XWPFTableRow tableRowOne2 = table.createRow();
	  	tableRowOne2.setHeight((int)(twipsPerInch*2/10)); //set height 1/10 inch.
	  	tableRowOne2.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT); //set w:hRule="exact"
	  	
	  	
	  	XWPFTableCell cell24=tableRowOne2.getCell(0);
	  	CTTcPr tcpr21 = cell24.getCTTc().addNewTcPr();
	  	CTVMerge vMerge21=tcpr21.addNewVMerge();
	  	vMerge21.setVal(STMerge.CONTINUE); 
	  	
	  	XWPFTableCell cell25=tableRowOne2.createCell();
	  	CTTcPr tcpr22 = cell25.getCTTc().addNewTcPr();
	  	CTVMerge vMerge22=tcpr22.addNewVMerge();
	  	vMerge22.setVal(STMerge.CONTINUE); 
	  	
	  	XWPFTableCell c3 = tableRowOne2.createCell();
	  	run1 = c3.addParagraph().createRun();
	  	run1.setBold(true);run1.setText("Name");run1.setFontSize(12);
	  	c3.removeParagraph(0);
	  	
	  	for(int i=0;i<subs.length;i++) {
	      c3 = tableRowOne2.addNewTableCell();
	      run1 = c3.addParagraph().createRun();
		  	run1.setBold(true);run1.setText("A");run1.setFontSize(12);
		  	c3.removeParagraph(0);
		  	c3 = tableRowOne2.addNewTableCell();
		      run1 = c3.addParagraph().createRun();
			  	run1.setBold(true);run1.setText("%");run1.setFontSize(12);
			  	c3.removeParagraph(0);
	  	}
	     
	  	int[] cols = new int[3+(subs.length*2)];
	    cols[0] = 8000;
	    cols[1] = 20000;
	    cols[2] = 20000;
	    for(int x=0; x<(subs.length*2); x++)
	    	cols[3+x] = 8000;
		     
	      for(int i = 0; i < table.getNumberOfRows(); i++){ 
	            XWPFTableRow row = table.getRow(i); 
	            int numCells = row.getTableCells().size(); 
	            for(int j = 0; j < numCells; j++){ 
	                XWPFTableCell cell = row.getCell(j); 
	                cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(cols[j])); 
	            } 
	        } 
	       
	     
	                                                                                                                                    // "+studdat.get(0)+studdat.get(1)+"-"+sub+"\\"+finalDate+".xls");        
	      InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\"+studdat.get(0)+studdat.get(1)+"-"+coursecode+"\\"+finalDate+".xls");
	  	HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
	  	HSSFSheet sheet = wb.getSheetAt(0);
	  	HSSFRow row; 
	  	
	  	ArrayList<String> names = new ArrayList<String>();
	  	ArrayList<String> usns = new ArrayList<String>();
	  	
	  	for(int i=5;i<sheet.getPhysicalNumberOfRows();i++)
	  	{
	  		
	  	
	  	String usn = sheet.getRow(i).getCell(0).toString();
	  	String name = sheet.getRow(i).getCell(1).toString();
	    
	  	names.add(name);
	  	usns.add(usn);
	      
	  	}

	  	System.out.println(names);
	  	
	  	
	  	
	  	for(int i=0;i<names.size();i++)
	  	{
	  		XWPFTableRow r1 =  table.createRow();
	  		r1.getCell(0).setText(String.valueOf(i+1));
	  		r1.setHeight((int)(twipsPerInch*2/10)); //set height 1/10 inch.
	  		r1.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT); //set w:hRule="exact"
	  		table.getRow(i+4).createCell().setText(usns.get(i));
	  		table.getRow(i+4).createCell().setText(names.get(i));
	  		
	  		int k = 0;
	  		
	  		while(k < (subs.length*2))
	  		{
	  			table.getRow(i+4).createCell().setText("");
	  			table.getRow(i+4).createCell().setText("");
	  			for(int m = 0; m<big.size();  m++)
	  			System.out.println(big.get(m));
		  		for(int x=0; x < big.size(); x=x+3)
		  		{
		  			//System.out.println(big.get(x).get(0)+"-"+table.getRow(1).getCell(k+3).getText().toString()+"-"+((k/2)+3));
			  		if(big.get(x).get(0).equalsIgnoreCase(table.getRow(1).getCell(k+3).getText().toString()))
			  		{
			  			table.getRow(i+4).getCell(k+3).setText(big.get(x+1).get(i).toString());
			  			table.getRow(i+4).getCell(k+4).setText(big.get(x+2).get(i).toString());
			  		}
		  		}
	  			k = k + 2;
	  		}
	  	}
	  	
	  	File path=new File("C:\\Users\\"+System.getProperty("user.name")+"\\Documents\\SDM\\Attendance\\Consolidated\\"+studdat.get(0)+studdat.get(1)+"\\"+studdat.get(0)+studdat.get(1)+"consolidated["+finalDate+"].docx");
	  	FileOutputStream fileOut = new FileOutputStream(path);
        docX2.write(fileOut);
        fileOut.close();
		            
	            System.out.println(".docx written successully");
	              
	}

	   
	    
	public static class Person {
	   	 
        private final SimpleStringProperty usn;
    	private final SimpleStringProperty name;
        private final SimpleStringProperty classes;
        private final SimpleStringProperty per;
        
        
        private Person(String string,String string1, String string2,String string3) {
        	this.usn = new SimpleStringProperty(string);
        	this.name =new SimpleStringProperty(string1);
            this.classes =new SimpleStringProperty(string2);
            this.per =new SimpleStringProperty(string3);
            
        }
        
        

		public void setStyle(String string) {
			// TODO Auto-generated method stub
			
		}



		public String getUsn() {
            
            return usn.get();
        }
        
        public void setUsn(String u) {
           usn.set(u);
        }

        public String getName() {
            
            return name.get();
        }

        
        public void setName(String u) {
            name.set(u);
        }

        public String getClasses() {
        	
        	return classes.get();
        }

        public void setClasses(String u) {
            classes.set(u);
           
        }
        
        public String getPer() {
        	
        	return per.get();
        }

        public void setPer(String u) {
            per.set(u);
           
        }
        
        
       
	}
	*/

}
