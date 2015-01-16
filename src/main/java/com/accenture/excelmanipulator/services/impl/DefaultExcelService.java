package com.accenture.excelmanipulator.services.impl;

import com.accenture.excelmanipulator.mapper.ExcelFormat;
import com.accenture.excelmanipulator.services.ExcelService;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import java.io.*;
import java.util.*;

/**
 * Created by abhay.narain.phougat on 12/12/2014.
 */
@Component(value = "excelService")
public class DefaultExcelService implements ExcelService {


    @Override
    public <T extends ExcelFormat> Workbook readSheet(T format, FileInputStream excelInputStream, int sheetNumber) throws IOException {
        HSSFWorkbook wb = null;
        try {
            POIFSFileSystem fs = new POIFSFileSystem(excelInputStream);
            wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(sheetNumber);
            HSSFRow row;
            HSSFCell cell;
            Sheet pivotSheet = wb.createSheet();
            int rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;
            HashMap<String, HashMap<String, ArrayList<String>>> pivotMap = new HashMap<String, HashMap<String, ArrayList<String>>>();
            Map<String,Integer> pivotColumnsNumber = new HashMap<String,Integer>();
            int pivotRowColumnNumber = 0;
            int pivotValueColumnNumber = 0;

            Set<String> columnsNames = new HashSet<String>() ;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows

            if (rows > 0) {

                row = sheet.getRow(0);
                if (row != null) {
                    tmp = sheet.getRow(0).getPhysicalNumberOfCells();
                    if (tmp > cols) cols = tmp;
                    for (int c = 0; c < cols; c++) {
                        cell = row.getCell(c);
                        if (cell != null) {
                            if (format.getColumnNames().contains(cell.getStringCellValue())) {
                                pivotColumnsNumber.put(cell.getStringCellValue(),c);
                            } else if (format.getRowName().contains(cell.getStringCellValue())) {
                                pivotRowColumnNumber = c;
                            } else if (format.getValueName().contains(cell.getStringCellValue())) {
                                pivotValueColumnNumber = c;
                            }

                        }
                    }

                }

            }
            int startingCell = 0;
            int startingRow = 0;
            Row pivotHeaderRow = pivotSheet.createRow(0);
            Cell pivotCell = pivotHeaderRow.createCell(startingCell);
            pivotCell.setCellValue(format.getRowName());
           /* for (String columnName : format.getColumnNames()) {
                Cell pivotColumnCell = pivotHeaderRow.createCell(startingCell + 1);
                pivotColumnCell.setCellValue(columnName);
            }*/


            for (int r = 1; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    if (!CollectionUtils.isEmpty(pivotMap) && pivotMap.containsKey(row.getCell(pivotRowColumnNumber).getStringCellValue())) {
                        for (Map.Entry<String,Integer> column : pivotColumnsNumber.entrySet()) {

                            if(!CollectionUtils.isEmpty(pivotMap.get(row.getCell(pivotRowColumnNumber).getStringCellValue()).get(row.getCell(column.getValue()).getStringCellValue())))
                            {
                                pivotMap.get(row.getCell(pivotRowColumnNumber).getStringCellValue()).get(row.getCell(column.getValue()).getStringCellValue()).add(row.getCell(pivotValueColumnNumber).getStringCellValue());

                            }else
                            {
                                HashMap<String, ArrayList<String>> columnValues = pivotMap.get(row.getCell(pivotRowColumnNumber).getStringCellValue());
                                ArrayList<String> carrierList = new ArrayList<String>();
                                carrierList.add(row.getCell(pivotValueColumnNumber).getStringCellValue());
                                columnValues.put(row.getCell(column.getValue()).getStringCellValue(), carrierList);
                                columnsNames.add(row.getCell(column.getValue()).getStringCellValue());
                            }
                                                       //columnValues.put(pivotColumn,carrierList);
                        }

                    } else {
                        HashMap<String, ArrayList<String>> columnValues = new HashMap<String, ArrayList<String>>();
                        for (Map.Entry<String,Integer> column : pivotColumnsNumber.entrySet()) {
                            ArrayList<String> carrierList = new ArrayList<String>();
                            carrierList.add(row.getCell(pivotValueColumnNumber).getStringCellValue());
                            columnValues.put(row.getCell(column.getValue()).getStringCellValue(), carrierList);
                            columnsNames.add(row.getCell(column.getValue()).getStringCellValue());
                        }
                        pivotMap.put(row.getCell(pivotRowColumnNumber).getStringCellValue(), columnValues);
                    }


                }
            }

            for (int run = 0; run < columnsNames.size() ; run++) {
                Cell pivotColumnCell = pivotHeaderRow.createCell(run + 1);
                pivotColumnCell.setCellValue(columnsNames.toArray()[run].toString());
            }
            for (Map.Entry<String, HashMap<String, ArrayList<String>>> entry : pivotMap.entrySet()) {
                int cellSequence = 0;
                Row pivotRow = pivotSheet.createRow(++startingRow );
                Cell pivotRowNameCell = pivotRow.createCell(cellSequence);
                pivotRowNameCell.setCellValue(entry.getKey());
                for (Map.Entry<String, ArrayList<String>> columnEntries : entry.getValue().entrySet()) {
                    for(int i=0;i< columnsNames.toArray().length ;i++)
                    {
                        if(columnsNames.toArray()[i].equals(columnEntries.getKey()))
                        {
                            Cell pivotValueCell = pivotRow.createCell(i + 1);
                            pivotValueCell.setCellValue(columnEntries.getValue().toString());
                        }


                    }

                   // Cell pivotColumnCell = pivotHeaderRow.createCell(startingCell + 1);
                   // pivotColumnCell.setCellValue(columnEntries.getKey());

                }

            }

        } catch (Exception ioe) {
            ioe.printStackTrace();
        } finally {
            excelInputStream.close();
        }


        return wb;
    }

    @Override
    public <T extends ExcelFormat> T getExcelFormat(Class<T> formatClass, File formatFile) throws JAXBException, ClassCastException {
        //String packageName = formatClass.getPackage().getName();

        JAXBContext jc = JAXBContext.newInstance(formatClass);
        Unmarshaller u = jc.createUnmarshaller();
        T format = (T) u.unmarshal(formatFile);
        return format;
    }

    @Override
    public boolean writeSheet(Workbook sheet, FileOutputStream outputStream) {

        try {
            sheet.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return false;
    }


}
