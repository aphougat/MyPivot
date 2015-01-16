package com.itembriyo.excelmanipulator.controllers;


import com.itembriyo.excelmanipulator.mapper.ExcelFormat;
import com.itembriyo.excelmanipulator.services.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by abhay.narain.phougat on 12/12/2014.
 */
@Controller
@RequestMapping("/getpivot")
public class ExcelRendererController {

    @Autowired
    ExcelService excelService;

    private static final String rootPath = System.getProperty("catalina.home");

    /**
     * Size of a byte buffer to read/write file
     */
    private static final int BUFFER_SIZE = 4096;


    @RequestMapping(method = RequestMethod.POST)
    public String handlePost(Model model, @RequestParam("format") MultipartFile format, @RequestParam("excel") MultipartFile excel, @RequestParam("sheetNumber") String sheetNumber) {
        if (!format.isEmpty() && !excel.isEmpty()) {
            try {
                byte[] formatBytes = format.getBytes();
                byte[] excelBytes = excel.getBytes();

                // Creating the directory to store file

                File dir = new File(rootPath + File.separator + "tmpFiles");
                if (!dir.exists())
                    dir.mkdirs();

                // Create the file on server
                File serverFormatFile = new File(dir.getAbsolutePath()
                        + File.separator + format.getName()+System.currentTimeMillis());
                BufferedOutputStream formatStream = new BufferedOutputStream(
                        new FileOutputStream(serverFormatFile));
                formatStream.write(formatBytes);
                formatStream.close();

                File serverExcelFile = new File(dir.getAbsolutePath()
                        + File.separator + excel.getName()+System.currentTimeMillis()+".xls");
                BufferedOutputStream excelStream = new BufferedOutputStream(
                        new FileOutputStream(serverExcelFile));
                excelStream.write(excelBytes);
                excelStream.close();

                    /*logger.info("Server File Location="
                            + serverFile.getAbsolutePath());
*/
                ExcelFormat newFormat = excelService.getExcelFormat(ExcelFormat.class, serverFormatFile);

               Workbook workbook=  excelService.readSheet(newFormat, new FileInputStream(serverExcelFile), Integer.parseInt(sheetNumber));
                excelService.writeSheet(workbook,new FileOutputStream(serverExcelFile));
                model.addAttribute("file",serverExcelFile.getName()+".xls");
                return "downloadPivot";
            } catch (JAXBException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "failure";

    }

    @RequestMapping(method = RequestMethod.GET)
    public String handleGet(Model model) {
        return "uploadFile";
    }

    @RequestMapping(value = "/printFormat",method = RequestMethod.GET ,produces = "application/xml")
    @ResponseBody
    public ExcelFormat printFormat()
    {
        ExcelFormat format = new ExcelFormat();
        List<String> columns = new ArrayList<String>();
        columns.add("Wave1");
        columns.add("Wave2");
        columns.add("Wave3");
        columns.add("Wave4");
        format.setColumnNames(columns);
        format.setRowName("Modules");
        format.setValueName("Capability");
        try {


            JAXBContext jaxbContext = JAXBContext.newInstance(ExcelFormat.class);
            Marshaller jaxbMarshaller = jaxbContext.createMarshaller();

            // output pretty printed
            jaxbMarshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);


            jaxbMarshaller.marshal(format, System.out);

        } catch (JAXBException e) {
            e.printStackTrace();
        }
        return format;
    }

    @RequestMapping(value="/{id}", method=RequestMethod.GET)
    public void downloadPivot(HttpServletRequest request,
                              HttpServletResponse response, @PathVariable("id") String fileName) throws IOException
    {
        // construct the complete absolute path of the file
        String fullPath = rootPath + File.separator + "tmpFiles" +File.separator + fileName;
        File downloadFile = new File(fullPath);
        FileInputStream inputStream = new FileInputStream(downloadFile);
        ServletContext context = request.getServletContext();
        // get MIME type of the file
        String mimeType = context.getMimeType(fullPath);
        if (mimeType == null) {
            // set to binary type if MIME mapping not found
            mimeType = "application/octet-stream";
        }
        System.out.println("MIME type: " + mimeType);

        // set content attributes for the response
        response.setContentType(mimeType);
        response.setContentLength((int) downloadFile.length());

        // set headers for the response
        String headerKey = "Content-Disposition";
        String headerValue = String.format("attachment; filename=\"%s\"",
                downloadFile.getName());
        response.setHeader(headerKey, headerValue);

        // get output stream of the response
        OutputStream outStream = response.getOutputStream();

        byte[] buffer = new byte[BUFFER_SIZE];
        int bytesRead = -1;

        // write bytes read from the input stream into the output stream
        while ((bytesRead = inputStream.read(buffer)) != -1) {
            outStream.write(buffer, 0, bytesRead);
        }

        inputStream.close();
        outStream.close();

    }
}
