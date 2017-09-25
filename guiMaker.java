//DIEGO MARTINEZ

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;

public class guiMaker extends JFrame implements ActionListener {

	protected JLabel masterLocLabel;
	protected JTextField masterFileField;
	protected JLabel masterFileLabel;
	protected JButton masterFileButton;
	protected JTextField masterLocationField;
	
	protected JLabel vendorLocLabel;
	protected JTextField vendorFileField;
	protected JLabel vendorFileLabel;
	protected JButton vendorFileButton;
	protected JTextField vendorLocationField;
	
	protected JLabel saveLabel;
	protected JTextField saveLocationField;
	protected JButton runButton;
	protected JButton saveFileButton;
	protected JFileChooser fileCho1;
	protected JFileChooser fileCho2;
	protected JFileChooser fileCho3;
	protected String mFileName;
	protected String vFileName;
	protected File readFileM = null;
	protected File readFileV = null;
	protected File saveFile = null;
	public String[] skuM;
	public double[] mapM;
	public String[] skuV;
	public double[] mapV;
	protected String[] details;
	protected File newMaster;
	protected File newMasterDetails;
	
	guiMaker() {
		setTitle("Master File Updater");
		masterLocLabel = new JLabel("Master File Location:");
		masterLocationField = new JTextField(30);
		masterLocationField.setText("");
		masterLocationField.setEditable(false);
		masterFileField = new JTextField(30);
		masterFileField.setEditable(false);
		masterFileLabel = new JLabel("Master File Name:");
		masterFileButton = new JButton("Open Master File");
		masterFileButton.addActionListener(this);
		
		
		vendorLocLabel = new JLabel("Vendor File Location:");
		vendorLocationField = new JTextField(30);
		vendorLocationField.setText("");
		vendorLocationField.setEditable(false);
		vendorFileField = new JTextField(30);
		vendorFileField.setEditable(false);
		vendorFileLabel = new JLabel("Vendor File Name:");
		vendorFileButton = new JButton("Open Vendor File");
		vendorFileButton.addActionListener(this);
		
		saveLabel = new JLabel("Save Location:");
		saveLocationField = new JTextField(30);
		saveLocationField.setText("");
		saveLocationField.setEditable(false);
		saveFileButton = new JButton("Save Location");
		saveFileButton.addActionListener(this);
		
		runButton = new JButton("Run");
		runButton.addActionListener(this);
		
		fileCho1 = new JFileChooser();
		fileCho2 = new JFileChooser();
		fileCho3 = new JFileChooser();
		GridBagConstraints layout = new GridBagConstraints();
		setLayout(new GridBagLayout());
		
		layout.gridx = 0;
		layout.gridy = 0;
		layout.insets = new Insets(10,10,10,10);
		add(masterFileButton, layout);
		
		layout.gridx = 1;
		layout.gridy = 0;
		layout.insets = new Insets(10,10,10,10);
		add(masterLocLabel, layout);
		
		layout.gridx = 2;
		layout.gridy = 0;
		layout.insets = new Insets(10,10,10,10);
		add(masterLocationField, layout);
		
		layout.gridx = 1;
		layout.gridy = 1;
		layout.insets = new Insets(10,10,10,10);
		add(masterFileLabel, layout);
		
		layout.gridx = 2;
		layout.gridy = 1;
		layout.insets = new Insets(10,10,10,10);
		add(masterFileField, layout);
		
		layout.gridx = 0;
		layout.gridy = 3;
		layout.insets = new Insets(10,10,10,10);
		add(vendorFileButton, layout);
		
		layout.gridx = 1;
		layout.gridy = 3;
		layout.insets = new Insets(10,10,10,10);
		add(vendorLocLabel, layout);
		
		layout.gridx = 2;
		layout.gridy = 3;
		layout.insets = new Insets(10,10,10,10);
		add(vendorLocationField, layout);
		
		layout.gridx = 1;
		layout.gridy = 4;
		layout.insets = new Insets(10,10,10,10);
		add(vendorFileLabel, layout);
		
		layout.gridx = 2;
		layout.gridy = 4;
		layout.insets = new Insets(10,10,10,10);
		add(vendorFileField, layout);
		
		layout.gridx = 0;
		layout.gridy = 5;
		layout.insets = new Insets(10,10,10,10);
		add(saveFileButton, layout);
		
		layout.gridx = 1;
		layout.gridy = 5;
		layout.insets = new Insets(10,10,10,10);
		add(saveLabel, layout);
		
		layout.gridx = 2;
		layout.gridy = 5;
		layout.insets = new Insets(10,10,10,10);
		add(saveLocationField, layout);
		
		layout.gridx = 2;
		layout.gridy = 6;
		layout.insets = new Insets(10,10,10,10);
		add(runButton, layout);
		
			}
		@Override
		public void actionPerformed(ActionEvent event) {
			int fileOpenerValue = 0;
			String errorType = "Selected File is NOT .xlsx Format!";
			JButton source = (JButton) event.getSource();
			int skuColM = -1;
			int mapColM = -1;

			if(source == saveFileButton) {
				fileCho3.setCurrentDirectory(saveFile);
				fileCho3.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				fileOpenerValue = fileCho3.showOpenDialog(this);
				if (fileOpenerValue == JFileChooser.CANCEL_OPTION) {
				    JOptionPane.showMessageDialog(this,"No directory selected!");
				    saveLocationField.setText("");
				}
				else {
					saveFile = fileCho3.getSelectedFile();
					saveLocationField.setText(saveFile.getPath());
					String location = saveFile.getPath() + "\\newMaster.xlsx";
					String location2 = saveFile.getPath() + "\\newMasterDetails.xlsx";
					newMaster = new File(location);
					newMasterDetails = new File(location2);
				}
			}
			if (source == runButton) {
				if(readFileM == null || readFileV == null) {
				    JOptionPane.showMessageDialog(this,"Missing file(s)!");

				}
				
				else {
					try {
						FileInputStream fs = new FileInputStream(readFileM);
						XSSFWorkbook workbook = new XSSFWorkbook(fs);
						XSSFSheet mSheet = workbook.getSheetAt(0);
						Iterator<Row> rowIterator = mSheet.iterator();
										
						Row row = rowIterator.next();
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							if(cell.getStringCellValue().equals("SKU")){
								skuColM = cell.getColumnIndex();
							}
							if(cell.getStringCellValue().equals("MAP")){
								mapColM = cell.getColumnIndex();
							}
						}
						mapM = new double[mSheet.getLastRowNum()+1];
						skuM = new String[mSheet.getLastRowNum()+1];
						if(skuColM != -1 && mapColM != -1) {
							for(int i = 1; i < mSheet.getLastRowNum()+1; i++) {
								row = mSheet.getRow(i);
								Cell cell = row.getCell(skuColM);
								skuM[i] = cell.getStringCellValue();
								Cell cell2 = row.getCell(mapColM);
								mapM[i] = cell2.getNumericCellValue();
							}
						}
						else{
						    JOptionPane.showMessageDialog(this,"Master Missing SKU/MAP!");
						    return;

						}			
						fs.close();
						workbook.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
					try {
						FileInputStream fs = new FileInputStream(readFileV);
						XSSFWorkbook workbook = new XSSFWorkbook(fs);
						XSSFSheet mSheet = workbook.getSheetAt(0);
						Iterator<Row> rowIterator = mSheet.iterator();
										
						skuColM = -1;
						mapColM = -1;
						Row row = rowIterator.next();
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							if(cell.getStringCellValue().equals("SKU")){
								skuColM = cell.getColumnIndex();
							}
							if(cell.getStringCellValue().equals("MAP")){
								mapColM = cell.getColumnIndex();
							}
						}
						if(skuColM == -1 || mapColM == -1) {
						    JOptionPane.showMessageDialog(this,"Vendor Missing SKU/MAP!");
						    return;
						}
						else{
							skuV = new String[mSheet.getLastRowNum()+1];
							mapV = new double[mSheet.getLastRowNum()+1];
							for(int i = 1; i < mSheet.getLastRowNum()+1; i++) {
								row = mSheet.getRow(i);
								Cell cell = row.getCell(skuColM);
								skuV[i] = cell.getStringCellValue();
								Cell cell2 = row.getCell(mapColM);
								mapV[i] = cell2.getNumericCellValue();
								
							}
						}			
						fs.close();
						workbook.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
					details = new String[skuM.length];
					for(int i = 1; i < skuM.length; i++) {
						details[i] = "Not found";
					}
					for (int i = 1; i < skuM.length; i++) {
						for (int j = 1; j < skuV.length; j++) {
							if (skuM[i].equals(skuV[j])) {
								if(mapM[i] > mapV[j]) {
									mapM[i] = mapV[j];
									details[i] = "MAP decreased";
								}
								if(mapM[i] < mapV[j]) {
									mapM[i] = mapV[j];
									details[i] = "MAP increased";
								}
							}
						}
					}			
					try {
						Files.copy(Paths.get(readFileM.getPath()), Paths.get(newMaster.getPath()), StandardCopyOption.REPLACE_EXISTING);
						Files.copy(Paths.get(readFileM.getPath()), Paths.get(newMasterDetails.getPath()), StandardCopyOption.REPLACE_EXISTING);

						FileInputStream inS = new FileInputStream(newMasterDetails);
						FileInputStream inS2 = new FileInputStream(newMaster);
		                XSSFWorkbook workbook = new XSSFWorkbook(inS); 
		                XSSFWorkbook workbook2 = new XSSFWorkbook(inS2); 
		                XSSFSheet sheet = workbook.getSheetAt(0); 
		                XSSFSheet sheet2 = workbook2.getSheetAt(0); 
		            	Iterator<Row> rowIterator = sheet.iterator();
						Row row = rowIterator.next();
						Row row2 = null;
						int detailsCol = row.getLastCellNum();
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							if(cell.getStringCellValue().equals("MAP")){
								mapColM = cell.getColumnIndex();
							}
						}
						for(int i = 1; i < sheet.getLastRowNum()+1; i++) {
							row = sheet.getRow(i); //for 2 outputs
							row2 = sheet2.getRow(i);
							Cell cell = row.getCell(mapColM);
							cell.setCellValue(mapM[i]);
							Cell cell2 = row2.getCell(mapColM);
							cell2.setCellValue(mapM[i]);
							cell = row.createCell(detailsCol); // adds details
							cell.setCellValue(details[i]);
						}
						row = sheet.getRow(0);
						Cell cell = row.createCell(detailsCol);
						cell.setCellValue("Details");
						
		                inS.close();			
		                inS2.close();
		                FileOutputStream fs = new FileOutputStream(newMasterDetails);
		                FileOutputStream fs2 = new FileOutputStream(newMaster);
		                workbook.write(fs);
		                workbook2.write(fs2);
		                fs.close(); 
		                fs2.close();
					    JOptionPane.showMessageDialog(this,"DONE!");
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}

			if (fileOpenerValue == JFileChooser.APPROVE_OPTION) {
				if(source == masterFileButton) {
					fileCho1.setFileSelectionMode(JFileChooser.FILES_ONLY);
					fileOpenerValue = fileCho1.showOpenDialog(this);
					if (fileOpenerValue == JFileChooser.CANCEL_OPTION) {
					    JOptionPane.showMessageDialog(this,"No file selected!");
					}
					else{
						readFileM = fileCho1.getSelectedFile();
						masterFileField.setText(readFileM.getName());
						mFileName = readFileM.getName();
						String end = mFileName.substring(mFileName.length() - 5);
						if(end.equals(".xlsx")) {
							masterLocationField.setText(readFileM.getPath());
						}
						else {
							JOptionPane.showMessageDialog(this, errorType);
							masterFileField.setText("");
							readFileM = null;
						}
					}
				}
				if (source == vendorFileButton) {
					fileCho2.setFileSelectionMode(JFileChooser.FILES_ONLY);
					fileOpenerValue = fileCho2.showOpenDialog(this);
					if (fileOpenerValue == JFileChooser.CANCEL_OPTION) {
					    JOptionPane.showMessageDialog(this,"No file selected!");
					}
					else {
						readFileV = fileCho2.getSelectedFile();
						vendorFileField.setText(readFileV.getName());
						vFileName = readFileV.getName();
						String end = vFileName.substring(vFileName.length() - 5);
						if(end.equals(".xlsx")) {
							vendorLocationField.setText(readFileV.getPath());
						}
						else {
							JOptionPane.showMessageDialog(this, errorType);
							vendorFileField.setText("");
							readFileV = null;
						}
					}
				}
			}
			

			return;
			
		}
}

