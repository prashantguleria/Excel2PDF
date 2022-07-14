package com.prashant.excel2pdf;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.WorksheetVisibility;

import org.apache.pdfbox.multipdf.PDFMergerUtility;

import java.io.File;
import java.io.FileNotFoundException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;

public class App {
    public static void main(String[] args) {
      try{
          if(args.length < 2){
              System.out.println("Accepts 2 parameters");
              System.out.println("First parameter is input file path with fileName");
              System.out.println("Second Parameter is output directory where PDF file has to be written.");
              return;
          }
          String inputFilePath = args[0];
          String destinationDirectory = args[1];

          createPDFFromExcel(inputFilePath, destinationDirectory);
      }catch (Exception ex){
          ex.printStackTrace();
      }

    }

    static void createPDFFromExcel(String filePath, String destinationDirectory) throws Exception{
        Path tempFilePath = Files.createTempDirectory("pdfOut");
        try{
            File inputFile  = new File(filePath);
            String fileName  = inputFile.getName();
            String outputFileName = (fileName.split("\\."))[0] + ".pdf";
            //Create a Workbook instance and load an Excel file
            Workbook workbook = new Workbook();
            workbook.loadFromFile(filePath);
            int sheets = workbook.getWorksheets().getCount();
            ArrayList<String> filePaths = new ArrayList<>();
            for (int i = 0; i< sheets ; i++){
                Worksheet worksheet = workbook.getWorksheets().get(i);
                worksheet.getPageSetup().isFitToPage(true);
                if(worksheet.getVisibility().equals(WorksheetVisibility.Hidden) || worksheet.getVisibility().equals(WorksheetVisibility.StrongHidden)){
                    continue;
                }
                String filePagePath = tempFilePath.toString() + "/" +  i + ".pdf";
                worksheet.saveToPdf(filePagePath);
                filePaths.add(filePagePath);
            }

            mergePDFs(filePaths, outputFileName, destinationDirectory);
        }catch (Exception ex){
            throw ex;
        }
        finally {
            Files.list(tempFilePath).forEach(x -> {
                x.toFile().delete();
            });
            Files.delete(tempFilePath);
        }
    }


    static void mergePDFs(ArrayList<String> files, String fileName, String folderPath){
        String outputPDF = folderPath + "/" +fileName;
        try{
            PDFMergerUtility pdfMergerUtility = new PDFMergerUtility();
            pdfMergerUtility.setDestinationFileName(outputPDF);
            files.forEach( file -> {
                try {
                    pdfMergerUtility.addSource(new File(file));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            });
            pdfMergerUtility.mergeDocuments();
            System.out.println("File created succeully at " + folderPath);
        }catch (Exception ex){
            System.out.println("Error creating file at " + folderPath);
            ex.printStackTrace();
        }
    }
}
