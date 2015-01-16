package com.itembriyo.excelmanipulator.services;

import com.itembriyo.excelmanipulator.mapper.ExcelFormat;
import org.apache.poi.ss.usermodel.Workbook;

import javax.xml.bind.JAXBException;
import java.io.*;

/**
 * Created by abhay.narain.phougat on 12/12/2014.
 */
public interface ExcelService {

    public <T extends ExcelFormat> Workbook readSheet(T format, FileInputStream excelInputStream, int sheetNumber) throws IOException;

    public <T extends ExcelFormat> T getExcelFormat(Class<T> formatClass, File formatInputStream) throws JAXBException;

    public boolean writeSheet(Workbook sheet, FileOutputStream outputStream);

}
