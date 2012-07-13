package jp.co.snknet.common.excel.controller;

import jp.co.snknet.common.utility.StringUtility;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excelユーティリティ<br>
 * <br>
 * Excelのユーティリティ
 * 
 * @author Shinko
 * @version 1.0
 */
public final class ExcelUtility {

	/**
     * Excel列名から列インデックス番号を取得
     * 
     * @param	columnName String 列名（例：A, B, ... CB, ... IV）
     * @return	int 列インデックス番号
     * @throws	Exception
     */
	public static int getExcelColumnIndex(String columnName) throws Exception {
		int liResult = 0;
		
		int liStartArha = 65;//アルファベット開始　ASC番号
		int liArhaNum = 26;//アルファベットの数
		
		try {
			if ((StringUtility.getByteCount(columnName) < 1
						|| StringUtility.getByteCount(columnName) > 2)
					|| !StringUtility.isExistZenkaku(columnName)) {
				//
				throw new Exception();
			}
			
	        switch (StringUtility.getByteCount(columnName)) {
	        	case 1:
	        		int liChar = ((int) columnName.charAt(0)) - liStartArha;

	        		return liChar;

	        	case 2:
	        		int liChar1 = ((int) columnName.charAt(0)) - liStartArha;
	        		int liChar2 = ((int) columnName.charAt(1)) - liStartArha;

	        		return liChar2 + ((liChar1 + 1) * liArhaNum);
	        	default:
	        } 
			
		} catch (Exception ex) {
			throw new Exception("指定された列名：　" + columnName + " は正しくありません。");	
		}
		return liResult;
	}
	/**
     * Excel列インデックス番号から列名を取得
     * 
     * @param	列インデックス番号
     * @return	列名（例：A, B, ... CB, ... IV）
     * @throws	Exception
     */
	public static String getExcelColumnName(int columnIndex) throws Exception{
		String lsResult = "";
		
		int liStartArha = 65;//アルファベット開始　ASC番号
		int liArhaNum = 26;//アルファベットの数
		
		try {
			int liMod = columnIndex % liArhaNum;
			int liWaru = (columnIndex - liMod) / liArhaNum;

			
			int liModPlus = liMod + liStartArha;
			int liWaruPlus = (liWaru - 1) + liStartArha;
			if (liWaru > 0) {
				lsResult = (new Character((char) liWaruPlus).toString()) + (new Character((char) liModPlus).toString());
			} else {
				lsResult = new Character((char) (columnIndex + liStartArha)).toString();
			}
		} catch (Exception ex) {
			throw new Exception("指定された列番号：　" + String.valueOf(columnIndex) + " は正しくありません。");	
		}
		return lsResult;

	}
	/**
	 * 印刷範囲の選択範囲を取得
	 */
	public static CellRangeAddress getPrintArea(String printArea) throws Exception{
		String lsSheetName = "";
		String lsRowStartName = "";
		String lsRowEndName = "";
		String lsColumnStartName = "";
		String lsColumnEndName = "";

		//
		// 印刷範囲の文字列を解析
		//
		int liReadPlace = 0;
		for (int i = 0 ; i < printArea.length() ; i++) {
			String lsData = printArea.substring(i, i + 1);
			
			if (lsData.equals("!")
					|| lsData.equals("$")
					|| lsData.equals(":")) {
				liReadPlace ++;
			} else {
				// 印刷範囲文字列　＝　「'シート名'!$A$1:$C$4」
		        switch (liReadPlace) {

		        	case 0:
		        		// シート名
		        		lsSheetName += lsData;
		        		break;
		        	case 1:
		        		// !$
		        		break;
		        	case 2:
		        		// 開始セル列
		        		lsColumnStartName += lsData;		
		        		break;
		        	case 3:
		        		// 開始セル行
		        		lsRowStartName += lsData;		        		
		        		break;
		        	case 4:
		        		// $:
		        		break;
		        	case 5:
		        		// 終了セル列
		        		lsColumnEndName += lsData;
		        		break;
		        	case 6:
		        		// 終了セル行
		        		lsRowEndName += lsData;
		        		break;
		        	default:
		        }
			}
		}
		
		// 印刷範囲をインデックスに変換
		int liRowStartIndex = Integer.valueOf(lsRowStartName) - 1;
		int liRowEndIndex = ExcelUtility.getExcelColumnIndex(lsColumnStartName);
		int liColumnStartIndex = Integer.valueOf(lsRowEndName) - 1;
		int liColumnEndIndex = ExcelUtility.getExcelColumnIndex(lsColumnEndName);

		return new CellRangeAddress(liRowStartIndex, liColumnStartIndex, liRowEndIndex, liColumnEndIndex);
	}
	/**
	 * 印刷範囲の全文字列を取得
	 * 
	 * @param workBook Workbook ワークブック
	 * @param sheetIndex int シートインデックス
	 * @return String 
	 */
	public static String getPringAreaStringFull(Workbook workBook, int sheetIndex) {
		int liNowActiveSheetIndex = workBook.getActiveSheetIndex();
		// 該当シートをアクティブにする（こうしないと、nullしか取得できないため）
		workBook.setActiveSheet(sheetIndex);
		// 印刷範囲を取得
		String lsPrintAreaString = workBook.getPrintArea(sheetIndex);
		// 元のシートをアクティブにする
		workBook.setActiveSheet(liNowActiveSheetIndex);
		
		return lsPrintAreaString;
	}
	/**
	 * 印刷範囲の範囲文字列を取得
	 * 
	 * @param cellRangeAddress
	 * @return
	 * @throws Exception
	 */
	public static String getPrintAreaString(CellRangeAddress cellRangeAddress) throws Exception{
		String lsResult = "$" + ExcelUtility.getExcelColumnName(cellRangeAddress.getFirstColumn())
								+ "$" + String.valueOf(cellRangeAddress.getFirstRow() + 1)
								+ ":"
								+ "$" + ExcelUtility.getExcelColumnName(cellRangeAddress.getLastColumn())
								+ "$" + String.valueOf(cellRangeAddress.getLastRow() + 1);

		return lsResult;
	}
}
