package jp.co.snknet.common.excel.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excelコントローラ<br>
 * <br>
 * Excelの制御
 * 
 * @author Shinko
 * @version 1.0
 */
public class ExcelController {

	private Workbook workBook;		// ワーク用ブック

	String msFilePath = "";	// 対象Excelファイルパス
	private static final String DEF_SHEET_NAME_DEFUALT = "シート１";	// デフォルトシート名（新規作成時のみ使用）
	
    /**
     * コンストラクタ
     */
	public ExcelController() {
		//　初期化
		this.initialize();
	}
    /**
     * コンストラクタ
     * 
     * @param filePath ファイルパス
     */
	public ExcelController(String filePath) {
		//　初期化
		this.initialize();
		//　値を保持
		msFilePath = filePath;
	}
    /**
     * 初期化
     */
	private void initialize() {
		//
		this.clear();
		//
		msFilePath = "";
	}
    /**
     * クリア
     */	
	public void clear() {
		// ブックを新規
		workBook =  new HSSFWorkbook();
		// シートを作成
		workBook.createSheet(DEF_SHEET_NAME_DEFUALT);
	}
	/**
	 * ファイルパスを取得
	 * @return ファイルパス
	 */
	public String getFilePath() {
		return msFilePath;
	}
    /**
     * 初期化処理
     */
	public void open() throws Exception, IOException {
		
		//
		//　クリア
		//
		this.clear();
		
		//
		// Excelを開く
		//
		if (msFilePath.equals("")) {
			//　新規作成
			workBook = new HSSFWorkbook();
			
		} else {
			//　編集
			
			// Excelファイルの読み込み
			FileInputStream fis = new FileInputStream(msFilePath);
			POIFSFileSystem fs = new POIFSFileSystem(fis);

			// ワークブック・オブジェクトの取得
			workBook = new HSSFWorkbook(fs);
		}
	}
    /**
     * ファイルへ書き出し
     */
	public void save() throws Exception, FileNotFoundException{
		try{
			
			// ワークブック・オブジェクトをファイルとして出力 			 		 
			FileOutputStream fileOut = new FileOutputStream(msFilePath);
			workBook.write(fileOut);
			
			//ファイルを閉じる
			fileOut.close();

		} catch (Exception ex){
			throw ex;
		}
	}
    /**
     * ファイルへ書き出し
     * 
     * @param outputFilePath String 出力ファイルパス
     */
	public void save(String outputFilePath) throws Exception, FileNotFoundException{
		// ファイルパスを上書き保持
		msFilePath = outputFilePath;
		// ファイルへ書き出し
		this.save();		
	}
    /**
     * データを閉じる
     */
	public void close(){
		// ブックを開放
		workBook = null;
	}
	/**
	 * シート名を変更
	 * 
	 * @param sheetName String シート名
	 */
	public void setSheetName(String sheetName) {
		workBook.setSheetName(workBook.getActiveSheetIndex(), sheetName);
	}
	/**
	 * シート名を変更
	 * 
	 * @param sheetIndex int シートインデックス
	 * @param sheetName String シート名
	 */
	public void setSheetName(int sheetIndex, String sheetName) {
		workBook.setSheetName(sheetIndex, sheetName);		
	}
	/**
	 * シートを削除
	 * 
	 * @param sheetIndex
	 */
	public void deleteSheet(int sheetIndex) {
		workBook.removeSheetAt(sheetIndex);
	}
	/**
	 * シートのクローンを作成
	 * 
	 * @param sheetName String クローン元のシート名
	 */
	public void cloneSheet(String sheetName) throws Exception{
		int liSheetIdx = workBook.getSheetIndex(sheetName);
		this.cloneSheet(liSheetIdx);		
	}
	/**
	 * シートのクローンを作成
	 * 
	 * @param sheetIndex int クローン元のシートインデックス
	 */
	public void cloneSheet(int sheetIndex) throws Exception{
		// シートのクローンを作成
		workBook.cloneSheet(sheetIndex);
		// 印刷範囲もコピー
		CellRangeAddress lclsRange = ExcelUtility.getPrintArea(ExcelUtility.getPringAreaStringFull(workBook, sheetIndex));
		String lsPrintArea = ExcelUtility.getPrintAreaString(lclsRange);//"$A$1:$F$14";//
		
		workBook.setPrintArea(this.getSheetCount() - 1, lsPrintArea);
	}
	/**
	 * Sheetコントローラを取得
	 * 
	 * @return SheetController シートコントローラ
	 */
	public SheetController getSheet() {
		return this.getSheet(workBook.getActiveSheetIndex());		
	}
	/**
	 * Sheetコントローラを取得
	 * 
	 * @param sheetIndex int シートインデックス
	 * @return SheetController シートコントローラ
	 */
	public SheetController getSheet(int sheetIndex) {
		return new SheetController(workBook.getSheetAt(sheetIndex));		
	}
	/**
	 * シート数を取得
	 */
	public int getSheetCount() {
		return workBook.getNumberOfSheets();
	}
	/**
	 * 選択シートインデックスを取得
	 * 
	 * @return int 選択しているシートのインデックス
	 */
	public int getSelectedSheetIndex() {
		return workBook.getActiveSheetIndex();
	}
	/**
	 * 選択シートインデックスを代入
	 * 
	 * @param sheetIndex String 選択するシートのインデックス
	 */
	public void setSelectedSheetIndex(int sheetIndex) {
		workBook.setActiveSheet(sheetIndex);
	}
	/**
	 * セルに値を代入
	 * 
	 * @param rowIndex
	 * @param column
	 * @param value
	 */
	public void setCellValue(int rowIndex, int columnIndex, String value) {

		SheetController lclsSheet = this.getSheet();
		// セルに値を代入
		lclsSheet.setCellValue(rowIndex, columnIndex, value);
	}
	/**
	 * セルの値をString型で取得
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @return
	 */
	public String getCellStringValue(int rowIndex, int columnIndex){
		return this.getSheet().getCellValue(rowIndex, columnIndex);
	}

	
	public void setPrintArea(int aiAddColumnNum) {

		String lsSheetName = "";
		String lsRowStartName = "";
		String lsColumnStartName = "";
		String lsRowEndName = "";
		String lsColumnEndName = "";
		
		String lsPrintArea = workBook.getPrintArea(0);

		int liReadPlace = 0;
		for (int i = 0 ; i < lsPrintArea.length() ; i++) {
			String lsData = lsPrintArea.substring(i, i + 1);
			
			if (lsData.equals("!")
					|| lsData.equals("$")
					|| lsData.equals(":")) {
				liReadPlace ++;
			} else {
				// 'シート名'!$A$1:$C$4
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
		
		int liAddRowIndex = Integer.valueOf(lsRowEndName) + aiAddColumnNum;
		
		String lsNewPrintArea = "$" + lsColumnStartName + "$" + lsRowStartName + ":$" + lsColumnEndName + "$" + String.valueOf(liAddRowIndex);
		
		workBook.setPrintArea(0, lsNewPrintArea);
	}
	/**
	 * 列数の取得
	 * 
	 * @return int 列数
	 */
	public int getColumnCount() {
		return this.getSheet().getColumnCount();
	}
	/**
	 * 列数の取得
	 * 
	 * @param シートインデックス
	 * @return int 列数
	 */
	public int getColumnCount(int sheetIndex) {
		return this.getSheet(sheetIndex).getColumnCount();
	}
	/**
	 * 行数の取得
	 * 
	 * @return　int 行数
	 */
	public int getRowCount() {
		return this.getSheet().getRowCount();
	}
	/**
	 * 行数の取得
	 * 
	 * @param シートインデックス
	 * @return　int 行数
	 */
	public int getRowCount(int sheetIndex) {
		return this.getSheet(sheetIndex).getRowCount();
	}
	public int getPrintRowCount(int sheetIndex) throws Exception{
		CellRangeAddress lclsRange = ExcelUtility.getPrintArea(workBook.getPrintArea(sheetIndex));
		
		return lclsRange.getLastRow() + 1;
	}

	/**
	 * テンプレートとなるシートの内容をコピー
	 * 
	 * @param templateSheetIndex
	 * @param pageNum
	 */
	public void copyTemplateSheet(int templateSheetIndex, int pageNum) throws Exception{
		this.copyTemplateSheet(this.getSelectedSheetIndex(), templateSheetIndex, pageNum);
	}
	/**
	 * テンプレートとなるシートの内容をコピー
	 * 
	 * @param targetSheetIndex
	 * @param templateSheetIndex
	 * @param pageNum
	 */
	public void copyTemplateSheet(int targetSheetIndex, int templateSheetIndex, int pageNum) throws Exception{
		Sheet lclsTargetSheet = workBook.getSheetAt(targetSheetIndex);
		Sheet lclsTemplateSheet = workBook.getSheetAt(templateSheetIndex);

		//
		// 対象の開始行インデックスを取得
		//
		CellRangeAddress lclsPrintRange = ExcelUtility.getPrintArea(ExcelUtility.getPringAreaStringFull(workBook, templateSheetIndex));
		int liTemplateRowCount = lclsPrintRange.getLastRow() + 1;
		int liTemplateColumnCount = lclsPrintRange.getLastColumn() + 1;
		int liStartRowIndex = (liTemplateRowCount * (pageNum - 1));
		
		//
		//　結合セルのコピー
		//
		int liMargedCount = lclsTemplateSheet.getNumMergedRegions();
		
		for (int i = 0 ; i < liMargedCount ; i++) {
			// 結合範囲を取得
			CellRangeAddress lclsRange = lclsTemplateSheet.getMergedRegion(i);
			// 結合範囲からコピー先の結合範囲インデックスを取得
			int liFirstRowIndex = lclsRange.getFirstRow() + liStartRowIndex;
			int liFirstColumnIndex = lclsRange.getFirstColumn();
			int liLastRowIndex = lclsRange.getLastRow() + liStartRowIndex;
			int liLastColumnIndex = lclsRange.getLastColumn();
			// 結合範囲を追加
			lclsTargetSheet.addMergedRegion(new CellRangeAddress(liFirstRowIndex, liLastRowIndex, liFirstColumnIndex, liLastColumnIndex));
		}
		
		//
		// 行のコピー
		//
		
		//
		// セルのコピー
		//
		for (int liRow = 0 ; liRow < liTemplateRowCount ; liRow++) {
		
			Row lclsTargetRow = lclsTargetSheet.createRow(liStartRowIndex + liRow);
			Row lclsTempRow = lclsTemplateSheet.getRow(liRow);
					
			if (lclsTempRow != null) {
				for (int liCol = 0 ; liCol < liTemplateColumnCount ; liCol++) {
					Cell lclsTargetCell = lclsTargetRow.createCell(liCol);
					Cell lclsTempCell = lclsTempRow.getCell(liCol);
					
					// 値を取得
					if (lclsTempCell != null) {
						switch (lclsTempCell.getCellType()) {
							case HSSFCell.CELL_TYPE_BLANK :
								break;
							case HSSFCell.CELL_TYPE_BOOLEAN :
								lclsTargetCell.setCellValue(lclsTempCell.getBooleanCellValue());
								break;
							case HSSFCell.CELL_TYPE_ERROR :
								lclsTargetCell.setCellValue(lclsTempCell.getErrorCellValue());
								break;
							case HSSFCell.CELL_TYPE_FORMULA :
								lclsTargetCell.setCellValue(lclsTempCell.getStringCellValue());
								break;
							case HSSFCell.CELL_TYPE_NUMERIC :
								lclsTargetCell.setCellValue(lclsTempCell.getNumericCellValue());
								break;
							case HSSFCell.CELL_TYPE_STRING :
								lclsTargetCell.setCellValue(lclsTempCell.getStringCellValue());
								break;
							default :
						}

						// スタイルを取得
						lclsTargetCell.setCellStyle(lclsTempCell.getCellStyle());

					}
				}			
			}
		}
		
		//
		// 改ページの挿入
		//
		lclsTargetSheet.setRowBreak(liStartRowIndex + (liTemplateRowCount - 1));

		//
		// 印刷範囲の挿入
		//
		this.setPrintArea(liStartRowIndex + liTemplateRowCount);

	}

}