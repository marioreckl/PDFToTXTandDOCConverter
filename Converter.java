/*
 * This program takes a directory of .pdf files and extracts the data and creates
 * a .txt and a .doc file containing the same text data as the original .pdf file. If the
 * .pdf file does not have selectable text, it will run Optical Character Recognition (OCR)
 * on the .pdf file to get the data. 
 * Input is a directory containing the .pdf files 
 * Output are two new directories contained in a "Converted" directory for each of the new file types
 * (.txt and .doc) with the directory named TXTFiles and DocxFiles respectively 
 */

import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.*;

import net.sourceforge.tess4j.*;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.WindowConstants;

public class Converter {
	
	//minimum number of characters in a string for a file not be be OCRed 
	private final static int MIN_DOC_LENGTH = 1500;  
	
	public static void main(String[] args) throws IOException, Docx4JException {
		
		String text;//holds the data from the .pdf file 
		
		//gets the input directory if it exist, if no exits 
		String  dir = JOptionPane.showInputDialog (null, "What is the directory? ");
		if (dir == null)
			System.exit(0);

		//gets the directory and creates a new file type to store it 
		File directory = new File(dir);
		
		//gets all the file in the directory with the .pdf extension
		File files[] = directory.listFiles(new FilenameFilter() {
			@Override
			public boolean accept(File dir, String name) {
				return name.endsWith(".pdf");
			}
		});
		
		//checks if it is able to create a new directory 
		if (new File(directory + "/Converted/").mkdir())
		{
			
			//creates a GUI for the user to see what is happening while the program runs 
			JFrame frame = new JFrame("Log");//GUI has title "Log" 
			frame.setSize(1000, 700); 
			JTextArea outTextArea = new JTextArea(25,50);//the text area
			JScrollPane scrollPane = new JScrollPane(outTextArea); //allow scrolling 
			scrollPane.setPreferredSize(new Dimension(450, 50));
			outTextArea.setEditable(false);//makes the text not editable 
			outTextArea.setVisible(true);//shows the text area 
			frame.add(scrollPane);//the text area to the GUI 
			frame.setVisible(true);//shows the GUI 
			
			//creates new directories for the two different output file types 
			String txtDir = directory + "/Converted/TXTFiles/";
			String docxDir = directory + "/Converted/DocxFiles/";
			
			//creates a log file 
			String logDir = directory + "/Converted/";
			File logOutFile = new File(logDir + "log.txt");
			//creates a new BufferedWriter object to write to a log file 
			BufferedWriter logWriter = new BufferedWriter(new FileWriter(logOutFile));
			
			new File(txtDir).mkdir();
			outTextArea.append(txtDir + " directory created \n");
			new File(docxDir).mkdir();
			outTextArea.append(docxDir + " directory created \n");
			
			//loops for all the file with a .pdf extension in the directory 
			for ( int i = 0; i < files.length; i++){

				//creates a new stripper object to get the text from the pdf file 
				PDFTextStripper textStripper=new PDFTextStripper();
				
				//loads the current pdf file
				PDDocument document=PDDocument.load(files[i]);
				outTextArea.append("File loaded: " + files[i] + "\n");
				
				//sets string to null 
				text = null;
				
				//creates output with that same name as the input files but with different extensions 
				File txtOutFile = new File(txtDir + files[i].getName() + ".txt");
				outTextArea.append("File created: " + txtOutFile + "\n");
				File docxOutFile = new File(docxDir + files[i].getName() + ".doc");
				outTextArea.append("File created: " + docxOutFile + "\n");
				
				//creates a new BufferedWriter object to output the data from the .pdf file 
				BufferedWriter txtWriter = new BufferedWriter(new FileWriter(txtOutFile));
				
				//creates a new word document to store the data from the .pdf file 
				WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
				MainDocumentPart docx = wordMLPackage.getMainDocumentPart();
				
				//gets data form the pdf file 
				text = textStripper.getText(document);
				
				//checks if the string of text is less then the minimum length for a file to have OCR run on it 
				if(text.length() < MIN_DOC_LENGTH){

					//notifies the user that OCR was needed on the file 
					outTextArea.append("OCR needed:  " + files[i] + "\n");
					
					//creates a new PDFRenderer object to be able to read the pdf file 
					PDFRenderer pdfRenderer = new PDFRenderer(document);
					
					//creates a new PDDocument
					PDDocument doc = new PDDocument();
					
					//Declares a Tesseract object to run OCR 
					Tesseract tesseract;
					//Declares a PDPage to store each page of the pdf file 
					PDPage ocrPage;
					
					//loops for each page in the pdf file 
					for (int page = 0; page < document.getNumberOfPages(); ++page)
					{ 
						//gets a image from each page of the pdf file 
						BufferedImage img = pdfRenderer.renderImageWithDPI(page, 300);
						
						//Initializes the Tesseract object
						tesseract = new Tesseract();
				        ocrPage = new PDPage();
				        
				        //adds a page to the doc Object
				        doc.addPage(ocrPage);
				      
						try {
							//runs OCR on the image and saves it to a sting
							String result = tesseract.doOCR(img);
							//outputs the string to a .txt file 
							txtWriter.write (result);
							//outputs the string to a .doc file 
							docx.addParagraphOfText(result);

						} catch (TesseractException e) //catches any exception in the OCR
						{
							outTextArea.append("Error: OCR in  " + files[i] + "\n");
							System.err.println(e.getMessage());
						}
					}
					
					//closes and saves all the files used and Objects needed to extract the data 
			        doc.close();
			        outTextArea.append("File saved: " + txtOutFile + "\n");
					document.close();
					txtWriter.close();
					wordMLPackage.save(docxOutFile);
					outTextArea.append("File saved: " + docxOutFile + "\n");
					}
				else//if OCR is not needed 
				{
					outTextArea.append("File saved: " + txtOutFile + "\n");
					//output the data to the new .txt file and saves it
					txtWriter.write(text.trim());

					//adds the data to the new .doc file 
					wordMLPackage.getMainDocumentPart().addParagraphOfText(text.trim());
					//closes and saves all the file 
					document.close();
					txtWriter.close();
					wordMLPackage.save(docxOutFile);
					outTextArea.append("File saved: " + docxOutFile + "\n");
					
				}
				}
		
			//checks to make sure all the files have been converted 
			String notOut = "";
			int count = 0;
			for ( int x = 0; x < files.length; x++){
				File convertedFile = new File(txtDir + files[x].getName() + ".txt");
				if (!(convertedFile.exists()))
				{
					count++; 
					notOut = notOut + files[x] + "\n";
				}
			}
			
			//notifies the user the conversion is complete and exits when the user closes the GUI 
			outTextArea.append("\n \n \nConversion Completed\n \n");
			outTextArea.append(count + " files not converted: \n");
			outTextArea.append(notOut);
			logWriter.write(outTextArea.getText());
			logWriter.close();
			frame.setVisible(true);
			frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
		}
		else //output a message that the directory already exits 
		{
			JOptionPane.showMessageDialog(null, "Converted directory alread exist, please remove before running again.");
		}
	
	}
}