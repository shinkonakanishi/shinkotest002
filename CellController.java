package jp.co.snknet.common.excel.controller;

import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel Cellコントローラ<br>
 * <br>
 * ExcelのCell制御
 * 
 * @author Shinko
 * @version 1.0
 */
public class CellController {

	Cell mclsCell;

	/**
	 * コンストラクタ
	 * 
	 * @param cell
	 */
	public CellController(Cell cell) {
		mclsCell = cell;
	}
	/**
	 * セルに値を代入
	 * 
	 * @param value
	 */
	public void setValue(Object value) {
		
        if ( value != null) {
        	// 型に合わせて値を代入
            if ( value instanceof String) {
            	// 文字列
            	mclsCell.setCellValue((String) value);
            } else if ( value instanceof Number) {
            	// 
                Number lnumValue = (Number) value;
                if (lnumValue instanceof Float) {
                    Float lfloatValue = (Float) lnumValue;
                    lnumValue = new Double(String.valueOf(lfloatValue));
                }
                
                mclsCell.setCellValue(lnumValue.doubleValue());
            } else if ( value instanceof Date) {
                Date ldateValue = (Date) value;
                mclsCell.setCellValue(ldateValue);
            } else if ( value instanceof Boolean) {
                Boolean lboolValue = (Boolean) value;
                mclsCell.setCellValue(lboolValue);
            }
        } else {
            CellStyle lclsCellStyle = mclsCell.getCellStyle();

        	mclsCell.setCellType(Cell.CELL_TYPE_BLANK);
        	mclsCell.setCellStyle(lclsCellStyle);
        }
	}
	/**
	 * セルの値をString型で取得
	 * 
	 * @return　String セルの値
	 */
	public  String getCellStringValue() {
		String lsResult = "";
		
		if (mclsCell != null) {
			switch (mclsCell.getCellType()) {
				case HSSFCell.CELL_TYPE_BLANK :
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN :
					if (mclsCell.getBooleanCellValue() == true) {
						
					} else {

					}
					break;
				case HSSFCell.CELL_TYPE_ERROR :
					if (mclsCell.getErrorCellValue() == 0) {
						
					} else {

					}

					break;
				case HSSFCell.CELL_TYPE_FORMULA :
					lsResult = mclsCell.getStringCellValue();
					break;
				case HSSFCell.CELL_TYPE_NUMERIC :
					lsResult = String.valueOf(mclsCell.getNumericCellValue());
					break;
				case HSSFCell.CELL_TYPE_STRING :
					lsResult = mclsCell.getStringCellValue();
					break;
				default :
			}
		}
		
		return lsResult;
	}
}
