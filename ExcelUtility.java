package jp.co.snknet.common.excel.controller;

import jp.co.snknet.common.utility.StringUtility;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excel���[�e�B���e�B<br>
 * <br>
 * Excel�̃��[�e�B���e�B
 * 
 * @author Shinko
 * @version 1.0
 */
public final class ExcelUtility {

	/**
     * Excel�񖼂����C���f�b�N�X�ԍ����擾
     * 
     * @param	columnName String �񖼁i��FA, B, ... CB, ... IV�j
     * @return	int ��C���f�b�N�X�ԍ�
     * @throws	Exception
     */
	public static int getExcelColumnIndex(String columnName) throws Exception {
		int liResult = 0;
		
		int liStartArha = 65;//�A���t�@�x�b�g�J�n�@ASC�ԍ�
		int liArhaNum = 26;//�A���t�@�x�b�g�̐�
		
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
			throw new Exception("�w�肳�ꂽ�񖼁F�@" + columnName + " �͐���������܂���B");	
		}
		return liResult;
	}
	/**
     * Excel��C���f�b�N�X�ԍ�����񖼂��擾
     * 
     * @param	��C���f�b�N�X�ԍ�
     * @return	�񖼁i��FA, B, ... CB, ... IV�j
     * @throws	Exception
     */
	public static String getExcelColumnName(int columnIndex) throws Exception{
		String lsResult = "";
		
		int liStartArha = 65;//�A���t�@�x�b�g�J�n�@ASC�ԍ�
		int liArhaNum = 26;//�A���t�@�x�b�g�̐�
		
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
			throw new Exception("�w�肳�ꂽ��ԍ��F�@" + String.valueOf(columnIndex) + " �͐���������܂���B");	
		}
		return lsResult;

	}
	/**
	 * ����͈͂̑I��͈͂��擾
	 */
	public static CellRangeAddress getPrintArea(String printArea) throws Exception{
		String lsSheetName = "";
		String lsRowStartName = "";
		String lsRowEndName = "";
		String lsColumnStartName = "";
		String lsColumnEndName = "";

		//
		// ����͈͂̕���������
		//
		int liReadPlace = 0;
		for (int i = 0 ; i < printArea.length() ; i++) {
			String lsData = printArea.substring(i, i + 1);
			
			if (lsData.equals("!")
					|| lsData.equals("$")
					|| lsData.equals(":")) {
				liReadPlace ++;
			} else {
				// ����͈͕�����@���@�u'�V�[�g��'!$A$1:$C$4�v
		        switch (liReadPlace) {

		        	case 0:
		        		// �V�[�g��
		        		lsSheetName += lsData;
		        		break;
		        	case 1:
		        		// !$
		        		break;
		        	case 2:
		        		// �J�n�Z����
		        		lsColumnStartName += lsData;		
		        		break;
		        	case 3:
		        		// �J�n�Z���s
		        		lsRowStartName += lsData;		        		
		        		break;
		        	case 4:
		        		// $:
		        		break;
		        	case 5:
		        		// �I���Z����
		        		lsColumnEndName += lsData;
		        		break;
		        	case 6:
		        		// �I���Z���s
		        		lsRowEndName += lsData;
		        		break;
		        	default:
		        }
			}
		}
		
		// ����͈͂��C���f�b�N�X�ɕϊ�
		int liRowStartIndex = Integer.valueOf(lsRowStartName) - 1;
		int liRowEndIndex = ExcelUtility.getExcelColumnIndex(lsColumnStartName);
		int liColumnStartIndex = Integer.valueOf(lsRowEndName) - 1;
		int liColumnEndIndex = ExcelUtility.getExcelColumnIndex(lsColumnEndName);

		return new CellRangeAddress(liRowStartIndex, liColumnStartIndex, liRowEndIndex, liColumnEndIndex);
	}
	/**
	 * ����͈͂̑S��������擾
	 * 
	 * @param workBook Workbook ���[�N�u�b�N
	 * @param sheetIndex int �V�[�g�C���f�b�N�X
	 * @return String 
	 */
	public static String getPringAreaStringFull(Workbook workBook, int sheetIndex) {
		int liNowActiveSheetIndex = workBook.getActiveSheetIndex();
		// �Y���V�[�g���A�N�e�B�u�ɂ���i�������Ȃ��ƁAnull�����擾�ł��Ȃ����߁j
		workBook.setActiveSheet(sheetIndex);
		// ����͈͂��擾
		String lsPrintAreaString = workBook.getPrintArea(sheetIndex);
		// ���̃V�[�g���A�N�e�B�u�ɂ���
		workBook.setActiveSheet(liNowActiveSheetIndex);
		
		return lsPrintAreaString;
	}
	/**
	 * ����͈͈͕͂̔�������擾
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
