package com.smlf.excel;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.smlf.excel.annotation.DataField;
import com.smlf.excel.annotation.Dataset;
import com.smlf.excel.util.ValueUtil;

public class ExportUtil {
	
	public static Workbook export(Workbook workbook, int sheetIndex, int dataStartRowIndex, List<?>... dataset){
		// dataset 不允许为空
        if (dataset==null || dataset.length==0) {
            throw new RuntimeException("导出错误 ------------- 数据集不允许为空");
        }
        List<?> sheetDataList = dataset[0];
        
        Class<?> dataModelClass = sheetDataList.get(0).getClass();
        Dataset excelSheet = dataModelClass.getAnnotation(Dataset.class);

        // 设置 sheet 名称
        String sheetName = excelSheet.sheetName();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if(sheet == null){
        	sheet = workbook.createSheet();
        }
        workbook.setSheetName(sheetIndex, sheetName);
        for(List<?> dataList : dataset){        	
        	write2Sheet(workbook, dataStartRowIndex, dataList, sheet);
        }
        
		return workbook;
	}

	private static void write2Sheet(Workbook workbook, int dataStartRowIndex, List<?> dataList,
			Sheet sheet) {
		// 获取数据模型类
		Class<?> dataModelClass = dataList.get(0).getClass();
        // 获取数据模型类的所有字段
		List<Field> fields = new ArrayList<Field>();
        if (dataModelClass.getDeclaredFields()!=null && dataModelClass.getDeclaredFields().length>0) {
            for (Field field: dataModelClass.getDeclaredFields()) {
            	// 过滤静态字段
                if (Modifier.isStatic(field.getModifiers())) {
                    continue;
                }
                DataField dataField = field.getAnnotation(DataField.class);
                // 过滤未使用 DataField 注解的字段
                if(dataField != null){
                	fields.add(field);
                }
            }
        }
        if (fields==null || fields.size()==0) {
            throw new RuntimeException("导出错误 ------------- 数据域不允许为空");
        }
        
        CellStyle[] dataFieldStyles = new CellStyle[fields.size()];
        int[] dataFieldWidth = new int[fields.size()];
        for (int i = 0; i < fields.size(); i++) {

            // 字段
            Field field = fields.get(i);
            DataField dataField = field.getAnnotation(DataField.class);
            int fieldWidth = 0;
            HorizontalAlignment align = null;
            if (dataField != null) {
                fieldWidth = dataField.width();
                align = dataField.align();
            }

            // 字段宽度
            dataFieldWidth[i] = fieldWidth;

            // 数据样式
            CellStyle dataFieldStyle = workbook.createCellStyle();
            if (align != null) {
                dataFieldStyle.setAlignment(align);
            }
            dataFieldStyles[i] = dataFieldStyle;
        }
        
        // 数据行
        for (int dataIndex = 0; dataIndex < dataList.size(); dataIndex++) {
            int rowIndex = dataIndex+dataStartRowIndex;
            Object rowData = dataList.get(dataIndex);

            Row row = sheet.getRow(rowIndex);
            if(row==null){
            	row = sheet.createRow(rowIndex);
            }

            for (int i = 0; i < fields.size(); i++) {
                Field field = fields.get(i);
                try {
                    field.setAccessible(true);
                    Object fieldValue = field.get(rowData);

                    String fieldValueString = ValueUtil.formatValue(field, fieldValue);
                    
                    // 根据索引获取单元格
                    DataField dataField = field.getAnnotation(DataField.class);
                    int columnIndex = dataField.columnIndex();
                    Cell cell = row.getCell(columnIndex);
                    if(cell==null){
                    	cell = row.createCell(columnIndex, CellType.STRING);
                    }
                    cell.setCellValue(fieldValueString);
                    cell.setCellStyle(dataFieldStyles[i]);
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        // 设置宽度
        for (int i = 0; i < dataFieldWidth.length; i++) {
            int fieldWidth = dataFieldWidth[i];
            if (fieldWidth > 0) {
                sheet.setColumnWidth(i, fieldWidth);
            } else {
                sheet.autoSizeColumn((short)i);
            }
        }
		
	}

}
