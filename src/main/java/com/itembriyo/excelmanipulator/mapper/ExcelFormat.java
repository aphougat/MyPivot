package com.itembriyo.excelmanipulator.mapper;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

import java.util.List;


/**
 * Created by abhay.narain.phougat on 12/12/2014.
 */
@XmlRootElement(name = "excelFormat")
public class ExcelFormat {

    List<String> columnNames;
    String rowName;
    String valueName;


    public List<String> getColumnNames() {
        return columnNames;
    }

    @XmlElement
    public void setColumnNames(List<String> columnNames) {
        this.columnNames = columnNames;
    }

    public String getValueName() {
        return valueName;
    }

    @XmlElement
    public void setValueName(String valueName) {
        this.valueName = valueName;
    }

    public String getRowName() {
        return rowName;
    }

    @XmlElement
    public void setRowName(String rowName) {
        this.rowName = rowName;
    }


}
