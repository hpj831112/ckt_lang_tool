/*
*Author: Wentu.Zheng
*Date: 2013/3/18 14:11:10
*/

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import com.ximpleware.AutoPilot;
import com.ximpleware.BookMark;
import com.ximpleware.EOFException;
import com.ximpleware.EncodingException;
import com.ximpleware.EntityException;
import com.ximpleware.ModifyException;
import com.ximpleware.NavException;
import com.ximpleware.ParseException;
import com.ximpleware.TranscodeException;
import com.ximpleware.VTDGen;
import com.ximpleware.VTDNav;
import com.ximpleware.XMLModifier;
import com.ximpleware.XPathEvalException;
import com.ximpleware.XPathParseException;

class SPOIExcel{
	private STable mTable;
	private SPOIUtility mUtility;
	private String mSheetName = "All";
	private SProgramOpt mOpt;
	private static final boolean DEBUG = true;

	static class SPOIUtility{
		public static final short STYLE_ID_TITLE = 0;
		public static final short STYLE_ID_TITLE_LANG = 1;
		public static final short STYLE_ID_LANG_NOID = 2;
		public static final short STYLE_ID_LANG_EMPTY = 3;
		public static final short STYLE_ID_LANG_NOTEMPTY = 4;
		public static final short STYLE_ID_PATH = 5;
		public static final short STYLE_ID_ARRAY = 6;
		public static final short STYLE_ID_CATEGORY = 7;
		public static final short STYLE_ID_HYPERLINK = 8;
		
		private HSSFFont fontDefault = null;
		private HSSFFont fontHyperLink = null;
		private HSSFCellStyle styleTitle = null;
		private HSSFCellStyle styleTitleLang = null;
		private HSSFCellStyle styleNoID = null;
		private HSSFCellStyle styleEmpty = null;
		private HSSFCellStyle styleNotEmpty = null;
		private HSSFCellStyle stylePath = null;
		private HSSFCellStyle stylePath1 = null;
		private HSSFCellStyle stylePath2 = null;
		private HSSFCellStyle styleArray = null;
		private HSSFCellStyle styleArray1 = null;
		private HSSFCellStyle styleArray2 = null;
		private HSSFCellStyle styleCategory = null;
		private HSSFCellStyle styleHyperLink = null;
		
		public static String int2Column(int i){
			String ret = "";
			int n = i;
			int d = 26;
			int r = 0;
			do{
				r = n%d;
				n = n/d;
				ret += String.format("%c", 'A' + r);
			}while(n != 0);
			return ret;
		}
		public void ChangeArrayStyle(){
			if(styleArray != styleArray1){
				styleArray = styleArray1;
			}else{
				styleArray = styleArray2;
			}
		}
		public void ChangePathStyle(){
			if(stylePath != stylePath1){
				stylePath = stylePath1;
			}else{
				stylePath = stylePath2;
			}
		}
		public HSSFCellStyle CreateStyle(HSSFWorkbook wb, short id){
			HSSFCellStyle styleRet = null;
			//Test and create default font
			if(null == fontDefault){
				fontDefault = wb.createFont();
				fontDefault.setFontHeightInPoints((short)10);
				fontDefault.setFontName("Arial");
			}
			
			//Test and create style according to id
			switch(id){
			case STYLE_ID_TITLE:{
				if(null == styleTitle){
					styleTitle = wb.createCellStyle();
					styleTitle.setFillForegroundColor(HSSFColor.RED.index);
					styleTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					styleTitle.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleTitle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleTitle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleTitle.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleTitle.setTopBorderColor(HSSFColor.BLACK.index);
					styleTitle.setBottomBorderColor(HSSFColor.BLACK.index);
					styleTitle.setLeftBorderColor(HSSFColor.BLACK.index);
					styleTitle.setRightBorderColor(HSSFColor.BLACK.index);
					styleTitle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleTitle.setFont(fontDefault);
				}
				styleRet = styleTitle;
				break;
			}
			case STYLE_ID_TITLE_LANG:{
				if(null == styleTitleLang){
					styleTitleLang = wb.createCellStyle();
					styleTitleLang.setFillForegroundColor(HSSFColor.GREEN.index);
					styleTitleLang.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleTitleLang.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleTitleLang.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleTitleLang.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleTitleLang.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleTitleLang.setTopBorderColor(HSSFColor.BLACK.index);
					styleTitleLang.setBottomBorderColor(HSSFColor.BLACK.index);
					styleTitleLang.setLeftBorderColor(HSSFColor.BLACK.index);
					styleTitleLang.setRightBorderColor(HSSFColor.BLACK.index);
					styleTitleLang.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleTitleLang.setFont(fontDefault);
				}
				styleRet = styleTitleLang;
				break;
			}
			case STYLE_ID_LANG_NOID:{
				if(null == styleNoID){
					styleNoID = wb.createCellStyle();
					styleNoID.setFillForegroundColor(HSSFColor.BLACK.index);
					styleNoID.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleNoID.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleNoID.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleNoID.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleNoID.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleNoID.setTopBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNoID.setBottomBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNoID.setLeftBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNoID.setRightBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNoID.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleNoID.setWrapText(true);
					styleNoID.setFont(fontDefault);
				}
				styleRet = styleNoID;
				break;
			}
			case STYLE_ID_LANG_EMPTY:{
				if(null == styleEmpty){
					styleEmpty = wb.createCellStyle();
					styleEmpty.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
					styleEmpty.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleEmpty.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleEmpty.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleEmpty.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleEmpty.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleEmpty.setTopBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleEmpty.setBottomBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleEmpty.setLeftBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleEmpty.setRightBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleEmpty.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleEmpty.setWrapText(true);
					styleEmpty.setFont(fontDefault);
				}
				styleRet = styleEmpty;
				break;
			}
			case STYLE_ID_LANG_NOTEMPTY:{
				if(null == styleNotEmpty){
					styleNotEmpty = wb.createCellStyle();
					styleNotEmpty.setFillForegroundColor(HSSFColor.WHITE.index);
					styleNotEmpty.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleNotEmpty.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleNotEmpty.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleNotEmpty.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleNotEmpty.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleNotEmpty.setTopBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNotEmpty.setBottomBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNotEmpty.setLeftBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNotEmpty.setRightBorderColor(HSSFColor.GREY_80_PERCENT.index);
					styleNotEmpty.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleNotEmpty.setWrapText(true);
					styleNotEmpty.setFont(fontDefault);
				}
				styleRet = styleNotEmpty;
				break;
			}
			case STYLE_ID_PATH:{
				if(null == stylePath1){
					stylePath1 = wb.createCellStyle();
					stylePath1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
					stylePath1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					stylePath1.setBorderTop(HSSFCellStyle.BORDER_THIN);
					stylePath1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					stylePath1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					stylePath1.setBorderRight(HSSFCellStyle.BORDER_THIN);
					stylePath1.setTopBorderColor(HSSFColor.TEAL.index);
					stylePath1.setBottomBorderColor(HSSFColor.TEAL.index);
					stylePath1.setLeftBorderColor(HSSFColor.TEAL.index);
					stylePath1.setRightBorderColor(HSSFColor.TEAL.index);
					stylePath1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					stylePath1.setFont(fontDefault);
					
					stylePath = stylePath1;
				}
				if(null == stylePath2){
					stylePath2 = wb.createCellStyle();
					stylePath2.setFillForegroundColor(HSSFColor.AQUA.index);
					stylePath2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					stylePath2.setBorderTop(HSSFCellStyle.BORDER_THIN);
					stylePath2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					stylePath2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					stylePath2.setBorderRight(HSSFCellStyle.BORDER_THIN);
					stylePath2.setTopBorderColor(HSSFColor.TEAL.index);
					stylePath2.setBottomBorderColor(HSSFColor.TEAL.index);
					stylePath2.setLeftBorderColor(HSSFColor.TEAL.index);
					stylePath2.setRightBorderColor(HSSFColor.TEAL.index);
					stylePath2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					stylePath2.setFont(fontDefault);
				}

				styleRet = stylePath;
				break;
			}
			case STYLE_ID_ARRAY:{
				if(null == styleArray1){
					styleArray1 = wb.createCellStyle();
					styleArray1.setFillForegroundColor(HSSFColor.CORNFLOWER_BLUE.index);
					styleArray1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleArray1.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleArray1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleArray1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleArray1.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleArray1.setTopBorderColor(HSSFColor.TEAL.index);
					styleArray1.setBottomBorderColor(HSSFColor.TEAL.index);
					styleArray1.setLeftBorderColor(HSSFColor.TEAL.index);
					styleArray1.setRightBorderColor(HSSFColor.TEAL.index);
					styleArray1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleArray1.setFont(fontDefault);
					
					styleArray = styleArray1;
				}
				if(null == styleArray2){
					styleArray2 = wb.createCellStyle();
					styleArray2.setFillForegroundColor(HSSFColor.LAVENDER.index);
					styleArray2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleArray2.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleArray2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleArray2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleArray2.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleArray2.setTopBorderColor(HSSFColor.TEAL.index);
					styleArray2.setBottomBorderColor(HSSFColor.TEAL.index);
					styleArray2.setLeftBorderColor(HSSFColor.TEAL.index);
					styleArray2.setRightBorderColor(HSSFColor.TEAL.index);
					styleArray2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleArray2.setFont(fontDefault);
				}
				
				styleRet = styleArray;
				break;
			}
			case STYLE_ID_CATEGORY:{
				if(null == styleCategory){
					styleCategory = wb.createCellStyle();
					styleCategory.setFillForegroundColor(HSSFColor.LAVENDER.index);
					styleCategory.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleCategory.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleCategory.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleCategory.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleCategory.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleCategory.setTopBorderColor(HSSFColor.TEAL.index);
					styleCategory.setBottomBorderColor(HSSFColor.TEAL.index);
					styleCategory.setLeftBorderColor(HSSFColor.TEAL.index);
					styleCategory.setRightBorderColor(HSSFColor.TEAL.index);
					styleCategory.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleCategory.setFont(fontDefault);
				}
				styleRet = styleCategory;
				break;
			}
			case STYLE_ID_HYPERLINK:{
				if(null == fontHyperLink){
					fontHyperLink = wb.createFont();
					fontHyperLink.setFontHeightInPoints((short)10);
					fontHyperLink.setFontName("Arial");
					fontHyperLink.setColor(HSSFColor.BLUE.index);
				}
				if(null == styleHyperLink){
					styleHyperLink = wb.createCellStyle();
					styleHyperLink.setFillForegroundColor(HSSFColor.LAVENDER.index);
					styleHyperLink.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					styleHyperLink.setBorderTop(HSSFCellStyle.BORDER_THIN);
					styleHyperLink.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					styleHyperLink.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					styleHyperLink.setBorderRight(HSSFCellStyle.BORDER_THIN);
					styleHyperLink.setTopBorderColor(HSSFColor.TEAL.index);
					styleHyperLink.setBottomBorderColor(HSSFColor.TEAL.index);
					styleHyperLink.setLeftBorderColor(HSSFColor.TEAL.index);
					styleHyperLink.setRightBorderColor(HSSFColor.TEAL.index);
					styleHyperLink.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					styleHyperLink.setFont(fontHyperLink);
				}
				styleRet = styleHyperLink;
				break;
			}
			}
			return styleRet;
		}
	}

	public SPOIExcel() {
		super();
		mUtility = new SPOIUtility();
		mTable = new STable();
		mOpt = new SProgramOpt();
	}
	public SPOIExcel(STable table) {
		super();
		mUtility = new SPOIUtility();
		mTable = table;
		mOpt = new SProgramOpt();
	}
	public SPOIExcel(STable table, SProgramOpt opt) {
		super();
		mUtility = new SPOIUtility();
		mTable = table;
		mOpt = opt;
	}
	public SProgramOpt GetOpt() {
		return mOpt;
	}
	public void SetOpt(SProgramOpt opt) {
		mOpt = opt;
	}
	private HSSFSheet CreateCategorySheet(HSSFWorkbook wb, String name){
		HSSFSheet sheet = wb.createSheet(name);
		sheet.setColumnWidth(0, 100*256);
		sheet.setColumnWidth(1, 10*256);
		sheet.setColumnWidth(2, 10*256);
		sheet.setColumnWidth(3, 10*256);
		HSSFCellStyle style = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_TITLE);
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(style);
		cell.setCellValue("Path List");
		return sheet;
	}
	private HSSFSheet CreateSheet(HSSFWorkbook wb, String name){
		HSSFSheet sheet = wb.createSheet(name);

		//Add AutoFilter file range
		CellRangeAddress rangeAddress = new CellRangeAddress(0, 0, 2, 0);
		sheet.setAutoFilter(rangeAddress);

		sheet.setColumnWidth(0, SExcelHeader.widthPath*256);
		sheet.setColumnWidth(1, SExcelHeader.widthFile*256);
		sheet.setColumnWidth(2, SExcelHeader.widthType*256);
		sheet.setColumnWidth(3, SExcelHeader.widthName*256);
		sheet.setColumnWidth(4, SExcelHeader.widthMsgID*256);
		sheet.setColumnWidth(5, SExcelHeader.widthLength*256);
		sheet.setColumnWidth(6, SExcelHeader.widthIndex*256);
		sheet.setColumnWidth(7, SExcelHeader.widthProduct*256);

		sheet.groupColumn(1, 2);
		sheet.groupColumn(4, 7);

		sheet.createFreezePane(0, 1);
		return sheet;
	}
	private void CreateTitle(HSSFWorkbook wb, HSSFSheet sheet, HSSFSheet refSheet, int nRow, boolean bLink){
		SExcelHeader titleHeder = new SExcelHeader();
		HSSFCellStyle styleDefault = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_TITLE);
		HSSFCellStyle styleLang = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_TITLE_LANG);
		HSSFCellStyle styleHyperLink = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_HYPERLINK);
		
		HSSFRow row = sheet.createRow(0);
		
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetPath());
		
		cell = row.createCell(1);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetFile());
		
		cell = row.createCell(2);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetType());
		
		cell = row.createCell(3);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetName());
		
		cell = row.createCell(4);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetMsgID());
		
		cell = row.createCell(5);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetLength());
		
		cell = row.createCell(6);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetIndex());
		
		cell = row.createCell(7);
		cell.setCellStyle(styleDefault);
		cell.setCellValue(titleHeder.GetProduct());
		
		ArrayList<String> langs = mTable.GetLangs();
		for(int i = 0; i < langs.size(); i++){
			sheet.setColumnWidth(SItemHeader.getLangIndex() + i, 30*256);
			cell = row.createCell(SItemHeader.getLangIndex() + i);
			cell.setCellStyle(styleLang);
			cell.setCellValue(langs.get(i));
		}
		
		if(bLink){
			int idx = SItemHeader.getLangIndex() + langs.size();
			
			HSSFHyperlink linkRef = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
			linkRef.setAddress(String.format("%s!%s%d:%s%d",refSheet.getSheetName(), SPOIUtility.int2Column(0), nRow, SPOIUtility.int2Column(3), nRow));
			cell = row.createCell(idx + 1);
			cell.setCellStyle(styleHyperLink);
			cell.setCellValue("Category");
			cell.setHyperlink(linkRef);
		}
	}
	private void InsertCategory(HSSFWorkbook wb, HSSFSheet sheet, int nRow, String path, HSSFSheet refSheet, HSSFSheet allSheet, int allRow){
		HSSFCellStyle style = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_CATEGORY);
		HSSFCellStyle styleHyperLink = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_HYPERLINK);

		HSSFRow row = sheet.createRow(nRow);
		
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(style);
		cell.setCellValue(path);
		
		int idx = 1;
		if(null != refSheet){
			HSSFHyperlink linkRef = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
			linkRef.setAddress(String.format("%s!A1",refSheet.getSheetName()));
			cell = row.createCell(idx++);
			cell.setCellStyle(styleHyperLink);
			cell.setCellValue("To Table");
			cell.setHyperlink(linkRef);
		}
		if(null != allSheet){
			HSSFHyperlink linkAll = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
			linkAll.setAddress(String.format("%s!A%d:H%d",allSheet.getSheetName(), allRow + 1, allRow + 1));
			cell = row.createCell(idx++);
			cell.setCellStyle(styleHyperLink);
			cell.setCellValue("To All");
			cell.setHyperlink(linkAll);
		}
		if(null != path){
			HSSFHyperlink linkFile = new HSSFHyperlink(HSSFHyperlink.LINK_FILE);
			linkFile.setAddress(".." + File.separator + path + "");
			cell = row.createCell(idx++);
			cell.setCellStyle(styleHyperLink);
			cell.setCellValue("To Folder");
			cell.setHyperlink(linkFile);
		}
	}
	private void InsertLine(HSSFWorkbook wb, HSSFSheet sheet, SItem item, int nRow){
		HSSFCellStyle styleNoID = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_NOID);
		HSSFCellStyle styleEmpty = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_EMPTY);
		HSSFCellStyle styleNotEmpty = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_NOTEMPTY);
		HSSFCellStyle stylePath = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_PATH);
		HSSFCellStyle styleArray = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_ARRAY);
		
		HSSFRow row = sheet.createRow(nRow);
		
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetPath());
		
		cell = row.createCell(1);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetFile());
		
		cell = row.createCell(2);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetType());
		
		cell = row.createCell(3);
		if(item.GetType().equals("string-array")){
			cell.setCellStyle(styleArray);
			cell.setCellValue(item.GetName());
		}else{
			cell.setCellStyle(stylePath);
			cell.setCellValue(item.GetName());
		}
		
		cell = row.createCell(4);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetMsgID());
		
		cell = row.createCell(5);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetLength());
		
		cell = row.createCell(6);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetIndex());
		
		cell = row.createCell(7);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetProduct());
		
		ArrayList<String> langs = mTable.GetLangs();
		for(int i = 0; i < langs.size(); i++){
			String value = item.GetTrans().get(langs.get(i));
			int langStartIdx = SItemHeader.getLangIndex();
			cell = row.createCell(langStartIdx + i);
			if((null == value)){
				cell.setCellStyle(styleNoID);
			}
			else if(value.isEmpty()){
				cell.setCellStyle(styleEmpty);
			}else{
				cell.setCellStyle(styleNotEmpty);
			}
			cell.setCellValue(value);
		}
	}
	private void InsertLine(HSSFWorkbook wb, HSSFSheet sheet, SItem item, int nRow, HSSFSheet refSheet, int refRow, boolean bRef){
		HSSFCellStyle styleNoID = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_NOID);
		HSSFCellStyle styleEmpty = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_EMPTY);
		HSSFCellStyle styleNotEmpty = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_LANG_NOTEMPTY);
		HSSFCellStyle stylePath = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_PATH);
		HSSFCellStyle styleArray = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_ARRAY);
		HSSFCellStyle styleHyperLink = mUtility.CreateStyle(wb, SPOIUtility.STYLE_ID_HYPERLINK);
		
		HSSFRow row = sheet.createRow(nRow);
		
		HSSFCell cell = row.createCell(0);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetPath());
		
		cell = row.createCell(1);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetFile());
		
		cell = row.createCell(2);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetType());
		
		cell = row.createCell(3);
		if(item.GetType().equals("string-array")){
			cell.setCellStyle(styleArray);
			cell.setCellValue(item.GetName());
		}else{
			cell.setCellStyle(stylePath);
			cell.setCellValue(item.GetName());
		}
		
		cell = row.createCell(4);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetMsgID());
		
		cell = row.createCell(5);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetLength());
		
		cell = row.createCell(6);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetIndex());
		
		cell = row.createCell(7);
		cell.setCellStyle(stylePath);
		cell.setCellValue(item.GetProduct());
		
		ArrayList<String> langs = mTable.GetLangs();
		for(int i = 0; i < langs.size(); i++){
			String value = item.GetTrans().get(langs.get(i));
			int langStartIdx = SItemHeader.getLangIndex();
			cell = row.createCell(langStartIdx + i);
			if((null == value)){
				cell.setCellStyle(styleNoID);
			}
			else if(value.isEmpty()){
				cell.setCellStyle(styleEmpty);
			}else{
				cell.setCellStyle(styleNotEmpty);
			}
			if(bRef){
				cell.setCellFormula(String.format("%s!%s%d", refSheet.getSheetName(), SPOIUtility.int2Column(langStartIdx + i), refRow + 1));
			}else{
				cell.setCellValue(value);
			}
		}
		
		int idx = SItemHeader.getLangIndex() + langs.size();
		
		HSSFHyperlink linkRef = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
		linkRef.setAddress(String.format("%s!%s%d:%s%d",refSheet.getSheetName(), SPOIUtility.int2Column(0), refRow + 1, SPOIUtility.int2Column(idx), refRow + 1));
		cell = row.createCell(idx);
		cell.setCellStyle(styleHyperLink);
		cell.setCellValue("Go To");
		cell.setHyperlink(linkRef);
	}
	private void ReadHeader(HSSFSheet sheet){
		ArrayList<String> langList = new ArrayList<String>();
		
		HSSFRow row = sheet.getRow(0);
		HSSFCell cell = null;

		//Calculate column count
		int nColumnHasContent = 0;
		int nColumnFromExcel = row.getLastCellNum() - row.getFirstCellNum() + 1;
		String cellValue;
		for(int i =0; i < nColumnFromExcel; i++){
			cell = row.getCell(nColumnHasContent);
			
			if(null == cell){
				break;
			}
			
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cellValue = cell.getStringCellValue();
			if((null == cellValue) || cellValue.isEmpty()){
				break;
			}
			nColumnHasContent++;
		}
		int nLangCount = nColumnHasContent - SItemHeader.getLangIndex();
		if(nLangCount <=0 ){
			System.err.println("Error: Excel file error!");
			System.exit(1);
		}
		//Read Language Columns
		for(int i = 0; i < nLangCount; i++){
			cell = row.getCell(SItemHeader.getLangIndex() + i);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			
			String langName = cell.getStringCellValue();
			if((null != langName) && (!langName.isEmpty())){
				langList.add(langName);
			}
		}
		mTable.SetLangs(langList);
	}
	private void ReadLine(HSSFSheet sheet, int idxRow){
		ArrayList<String> langs = mTable.GetLangs();
		
		HSSFRow row = sheet.getRow(idxRow);
		
		SItem item = new SItem();
		HSSFCell cell = null;  

		do{
			cell = row.getCell(0);
			if(null == cell){
				break;
			}
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			item.SetPath(row.getCell(0).getStringCellValue());
			
			cell = row.getCell(1);
			if(null == cell){
				break;
			}
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			item.SetFile(row.getCell(1).getStringCellValue());
			
			cell = row.getCell(2);
			if(null == cell){
				break;
			}
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			item.SetType(row.getCell(2).getStringCellValue());
			
			cell = row.getCell(3);
			if(null == cell){
				break;
			}
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			item.SetName(row.getCell(3).getStringCellValue());
			
			cell = row.getCell(4);
			if(null == cell){
				item.SetMsgID("");
			}else{
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				item.SetMsgID(row.getCell(4).getStringCellValue());
			}
			
			cell = row.getCell(5);
			if(null == cell){
				item.SetLength("0");
			}else{
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				item.SetLength(row.getCell(5).toString());
			}
			
			
			cell = row.getCell(6);
			if(null == cell){
				item.SetIndex("0");
			}else{
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				item.SetIndex(row.getCell(6).toString());
			}
			
			
			cell = row.getCell(7);
			if(null == cell){
				item.SetProduct("");
			}else{
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				item.SetProduct(row.getCell(7).toString());
			}
			
			
			String path = item.GetPath();
			for(int i = 0, count = langs.size(); i < count; i++){
				cell = row.getCell(SItemHeader.getLangIndex() + i);
				if(null == cell){
					item.GetTrans().put(langs.get(i), null);
				}else{
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					String value = cell.getStringCellValue();
					short color = cell.getCellStyle().getFillForegroundColor();
					if(color == HSSFColor.BLACK.index){
						item.GetTrans().put(langs.get(i), null);
					}else{
						item.GetTrans().put(langs.get(i), mOpt.isTranslate() ? convert(value) : value);
					}
				}
			}
			
			String fileName = item.GetFile();
			mTable.CheckAndAddItem(path);
			SPathItem pathItem = mTable.GetItem(path);
			
			pathItem.CheckAndAddItem(fileName);
			SFileItem fileItem = pathItem.GetItem(fileName);
			
			fileItem.InsertItem(item);
		}while(false);
		
	}

    private String convert(String value) {
        String result = value;
        System.out.println("origin: " + value);
        Pattern pattern = Pattern.compile("\\<[\\s\\S]*?\\>");
        Matcher matcher = pattern.matcher(result);
        boolean match = false;
        ArrayList<Integer> indexList = new ArrayList<Integer>();
        ArrayList<String> strList = new ArrayList<String>();
        ArrayList<String> tagList = new ArrayList<String>();
        ArrayList<String> notagList = new ArrayList<String>();

        while (matcher.find()) {
            match = true;
            indexList.add(matcher.start());
            indexList.add(matcher.end());
            tagList.add(value.substring(matcher.start(), matcher.end()));
        }

        if (match) {
            strList.add(value.substring(0, indexList.get(0)));
            for (int i = 0; i < indexList.size(); i++) {
                if (i + 1 < indexList.size()) {
                    strList.add(value.substring(indexList.get(i),
                            indexList.get(i + 1)));
                }
            }
            strList.add(value.substring(indexList.get(indexList.size() - 1)));

            for (int i = 0; i < strList.size(); i++) {

                if (i % 2 == 0 || i == strList.size() - 1) {
                    notagList.add(strList.get(i));
                }
            }

            System.out.println("strList: " + strList.toString());

            System.out.println("before nottagList: " + notagList.toString());
            processNoTagList(notagList);
            System.out.println("after nottagList: " + notagList.toString());

            System.out.println("before tagList: " + tagList.toString());

            ArrayList<String> processedTagList = new ArrayList<String>();

            if(tagList.size() > 2) {
                ArrayList<Integer> spliteIndexList = splitTagList(tagList);
                System.out.println("tagList: " + tagList.toString());

                for(int i = 0; i < spliteIndexList.size(); i = i + 2) {
                    int firstIndex = spliteIndexList.get(i);
                    int secondIndex = spliteIndexList.get(i + 1);
                    ArrayList<String> subTagList = new ArrayList<String>();
                    for(int j = firstIndex; j <= secondIndex; j++ ) {
                        subTagList.add(tagList.get(j));
                    }
                    System.out.println("i: " + i);
                    System.out.println("before subTagList: " + subTagList.toString());
                    processedTagList.addAll(processTagList(subTagList));
                }
            } else {
                processedTagList = processTagList(tagList);
            }

            //System.out.println("allTagList: " + allTagList.toString());

            if (strList.size() != processedTagList.size() + notagList.size()
                    && notagList.size() - processedTagList.size() != 1) {
                System.out
                        .println("error: "
                                + "strList size not equals processedTagList size plus notagList size");
            }

            System.out.println("begin mergerList: " + strList.toString());
            result = mergerList(processedTagList, notagList);

        } else {
            strList.add(value);
            processNoTagList(strList);
            result = strList.get(0);
        }

        System.out.println("result: " + result);
        System.out.println("------------------------------------------------------------------------------");

        return result;
    }

    private ArrayList<Integer> splitTagList(ArrayList<String> tagList) {
        ArrayList<Integer> spliteIndexList = new ArrayList<Integer> ();
        String firstTag = "";
        int theSameTag = 0;
        boolean updateFirstTag = true;

        for(int i = 0; i < tagList.size(); i++) {
            if(updateFirstTag) {
                firstTag = tagList.get(i).substring(0,tagList.get(i).replace(">", " ").indexOf(" ")) + ">";
                spliteIndexList.add(i);
            }

            if(i != 0 &&
                    updateFirstTag == false &&
                    (tagList.get(i).substring(0,tagList.get(i).replace(">", " ").indexOf(" ")) + ">")
                            .equals(firstTag)) {
                theSameTag++;

            }

            updateFirstTag = false;

            if(tagList.get(i).startsWith("</")
                    && ("<" +tagList.get(i).substring(2,tagList.get(i).length() -1) + ">").equals(firstTag)
                    && theSameTag == 0) {

                spliteIndexList.add(i);
                updateFirstTag = true;
            } else if(i != 0 &&
                    tagList.get(i).startsWith("</") &&
                    ("<" +tagList.get(i).substring(2,tagList.get(i).length() -1) + ">").equals(firstTag) &&
                    tagList.get(i).substring(2,tagList.get(i).length() -1)
                            .equals(tagList.get(i - 1).substring(1,tagList.get(i - 1).replace(">", " ").indexOf(" ")))) {
                theSameTag--;

            } else if(tagList.get(i).endsWith("/>") &&
                    (tagList.get(i).substring(0,tagList.get(i).replace(">", " ").indexOf(" ")) + ">").equals(firstTag)) {
                spliteIndexList.add(i);
                updateFirstTag = true;
            } else {
                continue;
            }
        }


        if(spliteIndexList.size() % 2 != 0) {
            spliteIndexList.add(tagList.size() - 1);
        }
        System.out.println("spliteIndexList: " + spliteIndexList);
        return spliteIndexList;
    }

    private String mergerList(ArrayList<String> tagList,
                                     ArrayList<String> notagList) {
        int index = 0;
        StringBuilder build = new StringBuilder();
        while (index < tagList.size()) {
            build.append(notagList.get(index) + tagList.get(index));
            index++;
        }
        build.append(notagList.get(index));

        return build.toString();
    }



    private ArrayList<String> processTagList(ArrayList<String> tagList) {
        ArrayList<String> pairTagList = new ArrayList<String>();
        ArrayList<Integer> pairTagListIndexInTaglist = new ArrayList<Integer>();
        ArrayList<String> nestedTagList = new ArrayList<String>();
        ArrayList<Integer> nestedTagListIndexInTaglist = new ArrayList<Integer>();
        ArrayList<String> mergerTagList = new ArrayList<String>();

        for(int i = 0; i < tagList.size(); i++) {
            String hopeTag = "", nextTag;
            if(pairTagListIndexInTaglist.contains(i)) {
                continue;
            }

            if(nestedTagList.size() > 0 && tagList.get(i).startsWith("</")) {
                hopeTag = "<" + tagList.get(i).substring(2,tagList.get(i).length() - 1);
                if(nestedTagList.get(nestedTagList.size() -1).startsWith(hopeTag)) {
                    pairTagList.add(nestedTagList.remove(nestedTagList.size() -1));
                    pairTagList.add(tagList.get(i));
                    pairTagListIndexInTaglist.add(nestedTagListIndexInTaglist.remove(nestedTagListIndexInTaglist.size() -1));
                    pairTagListIndexInTaglist.add(i);
                    continue;
                }
            }

            hopeTag = "</"+ tagList.get(i).substring(1,tagList.get(i).replace(">", " ").indexOf(" ")) + ">";


            if(i + 1 < tagList.size()) {
                nextTag = tagList.get(i+ 1);
            } else {
                nextTag = null;
            }

            if(nextTag != null && nextTag.equals(hopeTag)) {
                if(!pairTagListIndexInTaglist.contains(i)) {
                    pairTagList.add(tagList.get(i));
                    pairTagList.add(tagList.get(i + 1));
                    pairTagListIndexInTaglist.add(i);
                    pairTagListIndexInTaglist.add(i + 1);
                    i++;
                }

            } else if(tagList.get(i).endsWith("/>")) {
                if(!pairTagListIndexInTaglist.contains(i)) {
                    pairTagList.add(tagList.get(i));
                pairTagListIndexInTaglist.add(i);
            }
            } else {
                nestedTagList.add(tagList.get(i));
                nestedTagListIndexInTaglist.add(i);
            }

        }

        System.out.println("----pairTagList: " + pairTagList);
        System.out.println("----pairTagListIndexInTaglist: " + pairTagListIndexInTaglist);
        System.out.println("----nestedTagList: " + nestedTagList);
        System.out.println("----nestedTagListIndexInTaglist: " + nestedTagListIndexInTaglist);


        if(nestedTagList.size() > 0) {
            int halfTag = nestedTagList.size() % 2 == 0 ? nestedTagList.size() / 2 : (nestedTagList
                    .size() / 2) + 1;
            System.out.println("----halfTag: " + halfTag);
            for (int i = 0; i < nestedTagList.size(); i++) {
                String hopeTag, hopeTagPrefix;
                String realTag;


                if (i < halfTag) {
                    hopeTag = "</"
                            + nestedTagList.get(i).substring(1,
                            nestedTagList.get(i).replace(">", " ").indexOf(" ")).trim()
                            + ">";
                    realTag = nestedTagList.get(nestedTagList.size() - i - 1);
                    if (!hopeTag.equals(realTag)) {
                        System.out.println("----error tagIndex: " + i + " hopeTag: "
                                + hopeTag + " realTag: " + realTag);
                        nestedTagList.set(i, "&lt;" + nestedTagList.get(i).substring(1));
                    }
                } else {
                    hopeTag = "<" + nestedTagList.get(i).substring(2);
                    hopeTagPrefix = hopeTag.substring(0, hopeTag.length() - 1);
                    realTag = nestedTagList.get(nestedTagList.size() - i - 1);
                    if (!realTag.startsWith(hopeTagPrefix)) {
                        System.out.println("----error tagIndex: " + i
                                + " hopeTagPrefix: " + hopeTagPrefix + " realTag: "
                                + realTag);
                        if(!realTag.contains(hopeTagPrefix))
                            nestedTagList.set(i, "&lt;" + nestedTagList.get(i).substring(1));
                    }

                }
            }
        }


        for(int i = 0; i < tagList.size(); i ++) {
            if(pairTagListIndexInTaglist.contains(i)) {
                mergerTagList.add(pairTagList.get(pairTagListIndexInTaglist.indexOf(i)));
            } else if(nestedTagListIndexInTaglist.contains(i)) {
                mergerTagList.add(nestedTagList.get(nestedTagListIndexInTaglist.indexOf(i)));
            }
        }
        System.out.println("----mergerTagList: " + mergerTagList);
        return mergerTagList;

    }

    private ArrayList<String> processNoTagList(
            ArrayList<String> notagList) {
        for (int i = 0; i < notagList.size(); i++) {
            if (notagList.get(i).equals(" ")) {
                continue;
            }
            // do not adjust the order of contains("&")
            if (notagList.get(i).contains("&")) {
                notagList.set(i, notagList.get(i).replace("&", "&amp;"));
            }

            if (notagList.get(i).contains("<")) {
                notagList.set(i, notagList.get(i).replace("<", "&lt;"));
            }

            if (notagList.get(i).contains(">")) {
                notagList.set(i, notagList.get(i).replace(">", "&gt;"));
            }

            if (notagList.get(i).contains("\"")) {
                if(notagList.get(i).startsWith("\"") && notagList.get(i).endsWith("\"")
                        && !notagList.get(i).equals("\"")
                        && !notagList.get(i).equals("\"\"")) {
                    notagList.set(i,
                            notagList.get(i).substring(0,1) +
                                    notagList.get(i).substring(1,notagList.get(i).length() - 1).replace("\"", "&quot;") +
                                    notagList.get(i).substring(notagList.get(i).length() - 1));
                } else {
                    notagList.set(i, notagList.get(i).replace("\"", "&quot;"));
                }

            }

            if (notagList.get(i).contains("'")) {
                if(notagList.get(i).indexOf("'") != 0 &&
                        !notagList.get(i).substring(notagList.get(i).indexOf("'") - 1, notagList.get(i).indexOf("'")).equals("\\")) {
                notagList.set(i, notagList.get(i).replace("'", "\\&apos;"));
            }
        }

        }
        return notagList;
    }



	
	public STable GetTable() {
		return mTable;
	}
	public void Load(String strFile){
		HSSFWorkbook workbook;
		HSSFSheet sheetAll;
		File xlsFile = new File(strFile);
		if(!xlsFile.exists()){
			System.out.println("File missing: " + strFile);
			System.exit(0);
		}
		try {
			FileInputStream fileInputStream = new FileInputStream(xlsFile);
			workbook = new HSSFWorkbook(fileInputStream);
			sheetAll = workbook.getSheet(mSheetName);
			if(null == sheetAll){
				System.err.println(String.format("Error: A sheet named %s should exist!.", mSheetName));
				System.exit(1);
			}
			//Read header
			ReadHeader(sheetAll);
			int nRows = sheetAll.getLastRowNum();
			for(int idxRow = 1; idxRow <= nRows; idxRow++){
				ReadLine(sheetAll, idxRow);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public void Save(String file){
		HSSFWorkbook wb;
		HSSFSheet sheetCnt = null;
		HSSFSheet sheetCat = null;
		HSSFSheet sheetAll = null;

		try {
			wb = new HSSFWorkbook();
			
			//Create Summary Information
			wb.createInformationProperties();
			SummaryInformation si = wb.getSummaryInformation();
			si.setTitle("AWatcher");
			si.setRevNumber("V1.0");
			si.setSubject("String Table");
			si.setLastAuthor("WentuZheng");
			si.setKeywords("ATu");
			si.setComments("A string table for Android project!");
			si.setAuthor("WentuZheng");
			si.setApplicationName("AWatcher");
			si.setSecurity(0);
			DocumentSummaryInformation dsi = wb.getDocumentSummaryInformation();
			dsi.setCompany("TAB");
			dsi.setManager("WentuZheng");
			dsi.setCategory("Tool");
			
			//Store all data into excel file
			ArrayList<String> paths = mTable.GetKeys();
			sheetAll = CreateSheet(wb, mSheetName);
			sheetCat = CreateCategorySheet(wb, "Category");
			int nCurrentRow = 1;
			
			CreateTitle(wb, sheetAll, null, 0, false);
			for(int i = 0; i < paths.size(); i++){
				//Get Path Info
				String path = paths.get(i);
				SPathItem pathItem = mTable.GetItem(path);
				ArrayList<String> fileList = pathItem.GetKeys();
				
				//Create sheet for path
				int nCntSheetCount = 1;
				if(mOpt.isCreateSubTable()){
					sheetCnt = CreateSheet(wb, String.format("%d", i));
					CreateTitle(wb, sheetCnt, sheetCat, i + 2, true);
				}
				
				//Insert a category item
				InsertCategory(wb, sheetCat, i + 1, path, sheetCnt, sheetAll, nCurrentRow);
				
				//Begin insert data for path
				for(int j = 0; j < fileList.size(); j++){
					SFileItem fileItem = pathItem.GetItem(fileList.get(j));
					HashMap<String, SItem> stringIDItemMap = fileItem.GetItemMap();
					ArrayList<String> keys = fileItem.GetKeys();
					System.out.println(String.format("Save file %s%s%s to excel", paths.get(i), File.separator, fileList.get(j)));
					for(int k = 0; k < keys.size(); k++){
						String key = keys.get(k);
						SItem item = stringIDItemMap.get(key);
						//Change array style
						if(item.GetType().equals("string-array") && item.GetIndex().equals("0")){
							mUtility.ChangeArrayStyle();
						}
						//Insert data
						if(mOpt.isCreateSubTable()){
							InsertLine(wb, sheetAll, item, nCurrentRow, sheetCnt, nCntSheetCount, false);
							InsertLine(wb, sheetCnt, item, nCntSheetCount, sheetAll, nCurrentRow, true);
						}else{
							InsertLine(wb, sheetAll, item, nCurrentRow);
						}

						nCurrentRow++;
						nCntSheetCount++;
					}
				}
				mUtility.ChangePathStyle();
			}
			//Test and delete exist file
			File f = new File(file);
			if(f.exists()){
				f.delete();
			}
			//Save excel file
			FileOutputStream fos = new FileOutputStream(file);
			wb.write(fos);
			fos.flush();
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Saved all xml files to excel table, please check it!");
	}
	public void SetTable(STable table) {
		mTable = table;
	}
}

class SExcelHeader extends SItemHeader{
	public static int widthPath = 30;
	public static int widthFile = 20;
	public static int widthType = 12;
	public static int widthName = 30;
	public static int widthMsgID = 10;
	public static int widthLength = 5;
	public static int widthIndex = 5;
	public static int widthProduct = 10;
	
	public SExcelHeader() {
		super();
		SetPath(new String("Path"));
		SetFile(new String("File"));
		SetType(new String("Type"));
		SetName(new String("Name"));
		SetMsgID(new String("MsgID"));
		SetLength(new String("Len"));
		SetIndex(new String("Idx"));
		SetProduct(new String("Product"));
	}
}

class SFileItem{
	private HashMap<String, SItem> mMap_Item;
	private ArrayList<String> mKeys;
	private ArrayList<String> mBuffer;
	public SFileItem() {
		super();
		mMap_Item = new HashMap<String, SItem>();
		mKeys = new ArrayList<String>();
		mBuffer = new ArrayList<String>();
	}
	public HashMap<String, SItem> GetItemMap() {
		return mMap_Item;
	}
	public ArrayList<String> GetKeys() {
		return mKeys;
	}
	public void InsertItem(SItem item){
		if(mKeys.contains(item.GetKey())){
			MergeItem(item);
		}else{
			mKeys.add(item.GetKey());
			mMap_Item.put(item.GetKey(), item);
		}
	}
	public void InsertItemWithSort(SItem item){
		if(mKeys.contains(item.GetKey())){
			int pos = mKeys.indexOf(item.GetKey());
			if(!mBuffer.isEmpty()){
				mKeys.addAll(pos, mBuffer);
				mBuffer.clear();
			}
			MergeItem(item);
		}else{
			mBuffer.add(item.GetKey());
			mMap_Item.put(item.GetKey(), item);
		}
	}
	public void InsertItemWithSortBegin(){
		mBuffer.clear();
	}
	public void InsertItemWithSortEnd(){
		if(!mBuffer.isEmpty()){
			mKeys.addAll(mBuffer);
			mBuffer.clear();
		}
	}
	private void MergeItem(SItem item){
		SItem desItem = mMap_Item.get(item.GetKey());
		desItem.Merge(item);
	}
	public void SetItemMap(HashMap<String, SItem> items) {
		mMap_Item = items;
	}
	public void SetKeys(ArrayList<String> keys) {
		mKeys = keys;
	}
}

class SItem extends SItemHeader{
	private HashMap<String, String> mTrans;
	public SItem() {
		super();
		mTrans = new HashMap<String, String>();
	}
	public HashMap<String, String> GetTrans() {
		return mTrans;
	}
	public void Merge(SItem item){
		if(!item.GetMsgID().isEmpty()){
			SetMsgID(item.GetMsgID());
		}
		mTrans.putAll(item.GetTrans());
	}
}

class SItemHeader{
	private String mPath;
	private String mFile;
	private String mType;
	private String mName;
	private String mMsgID;
	private String mLength;
	private String mIndex;
	private String mProduct;
	private String mKey;
	private static int mLangIndex = 8;
	
	public SItemHeader() {
		super();
		mPath = new String();
		mFile = new String();
		mType = new String();
		mName = new String();
		mMsgID = new String();
		mLength = new String("0");
		mIndex = new String("0");
		mProduct = new String();
		mKey = new String();
	}
	public static int getLangIndex() {
		return mLangIndex;
	}
	public static void setLangIndex(int langIndex) {
		mLangIndex = langIndex;
	}
	public String GetFile() {
		return mFile;
	}
	public String GetIndex() {
		return mIndex;
	}
	public String GetLength() {
		return mLength;
	}
	public String GetMsgID() {
		return mMsgID;
	}
	public String GetName() {
		return mName;
	}
	public String GetPath() {
		return mPath;
	}
	public String GetProduct() {
		return mProduct;
	}
	public String GetType() {
		return mType;
	}
	public String GetKey() {
		mKey = GetName() + "_" + GetType() + "_" + GetProduct() + "_" + GetIndex();
		return mKey;
	}
	public void SetFile(String file) {
		mFile = file.trim();
	}
	public void SetIndex(String index) {
		mIndex = index.trim();
	}
	public void SetLength(String length) {
		mLength = length.trim();
	}
	public void SetMsgID(String msgID) {
		mMsgID = msgID.trim();
	}
	public void SetName(String name) {
		mName = name.trim();
	}
	public void SetPath(String path) {
		mPath = path.trim();
	}
	public void SetProduct(String product) {
		mProduct = product.trim();
	}
	public void SetType(String type) {
		mType = type.trim();
	}
}
class SPathItem{
	private HashMap<String, SFileItem> mMap_Item;
	private ArrayList<String> mKeys;
	public SPathItem() {
		mMap_Item = new HashMap<String, SFileItem>();
		mKeys = new ArrayList<String>();
	}
	public void AddItem(String key){
		mKeys.add(key);
		mMap_Item.put(key, new SFileItem());
	}
	public void CheckAndAddItem(String key){
		if(!mKeys.contains(key)){
			AddItem(key);
		}
	}
	public boolean ContainsItem(String key){
		return mKeys.contains(key);
	}
	public void DelItem(String key){
		if(mKeys.contains(key)){
			mKeys.remove(key);
			mMap_Item.remove(key);
		}
	}
	public SFileItem GetItem(String key){
		return mMap_Item.get(key);
	}
	public ArrayList<String> GetKeys(){
		return mKeys;
	}
}
class SProgramOpt{
	private boolean mMerge = false;
	private boolean mWriteBack = false;
	private boolean mCreateSubTable = false;
	private boolean mCreateNewFile = false;
        private boolean mTranslate = false;
        
	public boolean isMerge() {
		return mMerge;
	}
	public void setMerge(boolean bMerge) {
		mMerge = bMerge;
	}
	public boolean isWriteBack() {
		return mWriteBack;
	}
	public void setWriteBack(boolean bWriteBack) {
		mWriteBack = bWriteBack;
	}
	public boolean isCreateSubTable() {
		return mCreateSubTable;
	}
	public void setCreateSubTable(boolean bCreateSubTable) {
		mCreateSubTable = bCreateSubTable;
	}
	public boolean isCreateNewFile() {
		return mCreateNewFile;
	}
	public void setCreateNewFile(boolean bCreateNewFile) {
		mCreateNewFile = bCreateNewFile;
	}
       
        public boolean isTranslate() {
            return mTranslate;
        }

        public void setTranslate(boolean translate) {
            mTranslate = translate;
        }
}

public class SProgram{
	private STable mTable;
	private SPOIExcel mExcel;
	private SXml mXml;
	private SProgramOpt mOpt;
	
	static class SProgramInfo{
		private static String mFile_Excel = "stringTable.xls";
		private static String mFile_List = "./out/List.txt";
		private static String mFile_ListNew = "./out/List_new.txt";
		private static String mFile_Config = "SConfig.cfg";
		private static String mTag_LangBegin = "LANGUAGE_BEGIN";
		private static String mTag_LangEnd = "LANGUAGE_END";
		private static String mTag_TranslateIncludedTagsBegin = "TRANSLATE_INCLUDED_TAGS_BEGIN";
		private static String mTag_TranslateIncludedTagsEnd="TRANSLATE_INCLUDED_TAGS_END";

	}
	

	public static void main(String args []){
		SProgram prog = new SProgram();
		
		for (int i = 0; i < args.length; i++) {
			String arg = args[i];
            if (arg.charAt(0) == '-') {
                for (int j = 1; j < arg.length(); j++) {
                    switch (arg.charAt(j)) {
                    case 'm': prog.mOpt.setMerge(true); break;
                    case 'w': prog.mOpt.setWriteBack(true); break;
                    case 's': prog.mOpt.setCreateSubTable(true); break;
                    case 'c': prog.mOpt.setCreateNewFile(true); break;
                    case 't': prog.mOpt.setTranslate(true); break;
                    case 'u': break;
                    default:
                        System.out.println(String.format("%s: Unknown option '-%c'. Aborting.", "STool", arg.charAt(j)));
                        System.exit(1);
                    }
                }
            }
        }
		
		if(args.length >= 1){
			if(args[0].equals("x2e")){
				prog.x2e();
			}else if(args[0].equals("e2x")){
				prog.e2x();
			}else if(args[0].equals("createLang")){
				prog.mOpt.setMerge(false);
				prog.mOpt.setWriteBack(true);
				prog.CreateLang(args);
			}
		}else{
			System.out.println("Error: Parameter error, please execute \"./STool.sh help\" to get help!");
		}
	}
	public SProgram(){
		mTable = new STable();
		mOpt = new SProgramOpt();
		mExcel = new SPOIExcel(mTable, mOpt);
		mXml = new SXml(mTable, mOpt);
	}
	private void CreateLang(String args []){
		if(args.length < 4){
			System.out.println("Error: Parameter missing!");
			System.exit(1);
		}
		//Set Language map
		String langNew = args[1];
		String langTemplate = args[2];
		String langTrans = args[3];
		String langDef = "values";
		HashSet<String> langSet = new HashSet<String>();
		langSet.add(langTemplate);
		langSet.add(langNew);
		langSet.add(langTrans);
		langSet.add(langDef);
		mTable.SetLangs(langSet);

		mXml.Load(ReadList(SProgramInfo.mFile_ListNew));
		mTable.Translate(langNew, langTrans, langDef);
		mXml.Save(langNew);
	}
	private void e2x(){
		mExcel.Load(SProgramInfo.mFile_Excel);
		mXml.Save();
	}
	private ArrayList<String> ReadConfigList(String strFile,String beginTAG,String endTAG){
		ArrayList<String> list = new ArrayList<String>();

		do{
			try {
				File file = new File(strFile);
				if(!file.exists()){
					System.out.println("File missing: " + strFile);
					break;
				}
				FileReader fReader = new FileReader(strFile);
				BufferedReader bufReader = new BufferedReader(fReader);
				boolean isBegin = false;
				for (String item = bufReader.readLine(); null != item; item = bufReader.readLine()) {
					if(item.trim().equals(beginTAG)){
						isBegin = true;
						continue;
					}
					if(item.trim().equals(endTAG)){
						isBegin = false;
						break;
					}
					if(isBegin && !item.trim().isEmpty()){
						list.add(item.trim());
					}
				}
				bufReader.close();
				fReader.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}while(false);

		return list;
	}
	private ArrayList<String> ReadList(String strFile){
		ArrayList<String> list = new ArrayList<String>();

		do{
			try {
				File file = new File(strFile);
				if(!file.exists()){
					System.out.println("File missing: " + strFile);
					break;
				}
				FileReader fReader = new FileReader(strFile);
				BufferedReader bufReader = new BufferedReader(fReader);
				for (String item = bufReader.readLine(); null != item; item = bufReader.readLine()) {
					if((null != item) && (!item.trim().isEmpty())){
						list.add(item.trim());
					}
				}
				bufReader.close();
				fReader.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}while(false);

		return list;
	}
	private void x2e(){
		mTable.SetLangs(ReadConfigList(SProgramInfo.mFile_Config,SProgramInfo.mTag_LangBegin,SProgramInfo.mTag_LangEnd));
		mXml.Load(ReadList(SProgramInfo.mFile_List));
		mExcel.Save(SProgramInfo.mFile_Excel);
	}
}

class STable{
	private ArrayList<String> mLangs;
	private HashMap<String, SPathItem> mMap_Item;
	private ArrayList<String> mKeys;

	public STable() {
		mLangs = new ArrayList<String>();
		mMap_Item = new HashMap<String, SPathItem>();
		mKeys = new ArrayList<String>();
	}
	public void AddItem(String path){
		mKeys.add(path);
		mMap_Item.put(path, new SPathItem());
	}
	public void CheckAndAddItem(String path){
		if(!mKeys.contains(path)){
			AddItem(path);
		}
	}
	public boolean ContainsItem(String path){
		return mKeys.contains(path);
	}
	public ArrayList<String> GetLangs(){
		return mLangs;
	}

	public ArrayList<String> GetKeys(){
		return mKeys;
	}
	public SPathItem GetItem(String path){
		return mMap_Item.get(path);
	}
	public void SetLangs(ArrayList<String> langs){
		mLangs = langs;
	}
	public void SetLangs(HashSet<String> langs){
		Iterator<String> it = langs.iterator();
		for(int i = 0; it.hasNext(); i++){
			mLangs.add(it.next());
		}
	}

	public void Translate(String desLang, String srcLang, String defLang){
		for(int i = 0, count0 = mKeys.size(); i < count0; i++){
			String path = mKeys.get(i);
			SPathItem pathItem = GetItem(path);
			ArrayList<String> fileNameList = pathItem.GetKeys();
			for(int j = 0, count1 = fileNameList.size(); j < count1; j++){
				String fileName = fileNameList.get(j);
				SFileItem fileItem = pathItem.GetItem(fileName);
				ArrayList<String> keys = fileItem.GetKeys();
				HashMap<String, SItem> keyIDToItem = fileItem.GetItemMap();
				for(int k = 0, count2 = keys.size(); k < count2; k++){
					HashMap<String, String> translateMap = keyIDToItem.get(keys.get(k)).GetTrans();
					String strFromSrc = translateMap.get(srcLang);
					if(null != strFromSrc){
						translateMap.put(desLang, strFromSrc);
						continue;
					}
					String strFromDef = translateMap.get(defLang);
					if(null != strFromDef){
						translateMap.put(desLang, strFromDef);
						continue;
					}
				}
			}
		}
	}
}

class SXml{
	private STable mTable = null;
	private String mOutDir = "out";
	private SProgramOpt mOpt;
	
	public SXml() {
		super();
		mTable = new STable();
		mOpt = new SProgramOpt();
	}
	public SXml(STable table) {
		super();
		mTable = table;
		mOpt = new SProgramOpt();
	}
	public SXml(STable table, SProgramOpt opt) {
		super();
		mTable = table;
		mOpt = opt;
	}
	public STable GetTable() {
		return mTable;
	}
	public void SetTable(STable table) {
		mTable = table;
	}
	public String GetOutDir() {
		return mOutDir;
	}
	public void SetOutDir(String dir) {
		mOutDir = dir;
	}
	public SProgramOpt GetOpt() {
		return mOpt;
	}
	public void SetOpt(SProgramOpt opt) {
		mOpt = opt;
	}
	private void CheckAndSave(SPathItem pathItem, String fileName, String lang, String desXmlPath, String srcXmlPath){
		HashMap<String, SItem> items = pathItem.GetItem(fileName).GetItemMap();
		ArrayList<String> tempKeys = new ArrayList<String>(pathItem.GetItem(fileName).GetKeys());
		ArrayList<BookMark> tempBM = new ArrayList<BookMark>();
		for(int i = 0; i < tempKeys.size(); i++){
			tempBM.add(null);
		}

		do{
			boolean isChange = false;
			VTDGen vGen = new VTDGen();

			File srcFile = new File(srcXmlPath);
			if(!srcFile.exists()){
				if(mOpt.isCreateNewFile()){
					String defaultForamt = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
					String str0 = "<resources xmlns:android=\"http://schemas.android.com/apk/res/android\" xmlns:xliff=\"urn:oasis:names:tc:xliff:document:1.2\">";
					String str1 = "</resources>";
					defaultForamt = defaultForamt + "\n\n\n\n" + str0 + "\n" + str1;
					vGen.setDoc(defaultForamt.getBytes());
					try {
						vGen.parse(false);
					} catch (EncodingException e) {
						e.printStackTrace();
					} catch (EOFException e) {
						e.printStackTrace();
					} catch (EntityException e) {
						e.printStackTrace();
					} catch (ParseException e) {
						e.printStackTrace();
					}
				}else{
					break;
				}
			}else{
				if(!vGen.parseFile(srcXmlPath, true)){
					System.err.println(String.format("Paser file error: %s", srcXmlPath));
					break;
				}
			}
			
			VTDNav vNav = vGen.getNav();
			AutoPilot ap = new AutoPilot(vNav);
			try {
				XMLModifier vModifier = new XMLModifier(vNav);
				if(!vNav.toString(vNav.getRootIndex()).equals("resources")){
					break;
				}
				ap.selectXPath("/resources");
				for (boolean bResItem = vNav.toElement(VTDNav.FIRST_CHILD); false != bResItem; bResItem = vNav.toElement(VTDNav.NEXT_SIBLING)) {
					if (vNav.toString(vNav.getCurrentIndex()).equals("string")) {
						SItem item = new SItem();
						item.SetType("string");
						if (-1 != vNav.getAttrVal("name")) {
							String str = vNav.toString(vNav.getAttrVal("name"));
							item.SetName(str);
						}
						if (-1 != vNav.getAttrVal("product")) {
							String str = vNav.toString(vNav.getAttrVal("product"));
							item.SetProduct(str);
						}
						long l = vNav.getContentFragment();
						if (-1 != l) {
							String valueFromXml = vNav.toRawString((int) (l & 0xFFFFFF), (int) (l >> 32));
							SItem itemFromExcel = items.get(item.GetKey());
							if(null == itemFromExcel){
								continue;
							}
							String valueFromExcel = itemFromExcel.GetTrans().get(lang);
							if((null != valueFromExcel) && !valueFromXml.equals(valueFromExcel)){
								vModifier.removeContent((int) (l & 0xFFFFFF), (int) (l >> 32));
								vModifier.insertAfterHead(valueFromExcel);
								isChange = true;
							}
						}else{
							SItem itemFromExcel = items.get(item.GetKey());
							if(null == itemFromExcel){
								continue;
							}
							String valueFromExcel = itemFromExcel.GetTrans().get(lang);
							if((null != valueFromExcel) && !valueFromExcel.isEmpty()){
								vModifier.insertAfterHead(valueFromExcel);
								isChange = true;
							}
						}
						BookMark bm = new BookMark();
						bm.bind(vNav);
						bm.recordCursorPosition();
						tempBM.set(tempKeys.indexOf(item.GetKey()), bm);
					} else if (vNav.toString(vNav.getCurrentIndex()).equals("string-array")) {
						String nameValue = new String();
						if (-1 != vNav.getAttrVal("name")) {
							nameValue = vNav.toString(vNav.getAttrVal("name"));
						}
						String productValue = new String();
						if (-1 != vNav.getAttrVal("product")) {
							productValue = vNav.toString(vNav.getAttrVal("product"));
						}
						int idx = 0;
						vNav.push();
						for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item"), idx++) {
							SItem item = new SItem();
							item.SetType("string-array");
							item.SetName(nameValue);
							item.SetIndex(String.format("%d", idx));
							item.SetProduct(productValue);
							long l = vNav.getContentFragment();
							if (-1 != l) {
								String valueFromXml = vNav.toRawString((int) (l & 0xFFFFFF), (int) (l >> 32));
								SItem itemFromExcel = items.get(item.GetKey());
								if(null == itemFromExcel){
									continue;
								}
								String valueFromExcel = itemFromExcel.GetTrans().get(lang);
								if((null != valueFromExcel) && !valueFromXml.equals(valueFromExcel)){
									vModifier.removeContent((int) (l & 0xFFFFFF), (int) (l >> 32));
									vModifier.insertAfterHead(valueFromExcel);
									isChange = true;
								}
							}else{
								SItem itemFromExcel = items.get(item.GetKey());
								if(null == itemFromExcel){
									continue;
								}
								String valueFromExcel = itemFromExcel.GetTrans().get(lang);
								if((null != valueFromExcel) && !valueFromExcel.isEmpty()){
									vModifier.insertAfterHead(valueFromExcel);
									isChange = true;
								}
							}
							BookMark bm = new BookMark();
							bm.bind(vNav);
							bm.recordCursorPosition();
							tempBM.set(tempKeys.indexOf(item.GetKey()), bm);
						}
						vNav.pop();
					} else if (vNav.toString(vNav.getCurrentIndex()).equals("plurals")) {
						String nameValue = new String();
						if (-1 != vNav.getAttrVal("name")) {
							nameValue = vNav.toString(vNav.getAttrVal("name"));
						}
						String productValue = new String();
						if (-1 != vNav.getAttrVal("product")) {
							productValue = vNav.toString(vNav.getAttrVal("product"));
						}
						
						int idx = 0;
						vNav.push();
						for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item"), idx++) {
							String idxValue = new String();
							if (-1 != vNav.getAttrVal("quantity")) {
								idxValue = vNav.toString(vNav.getAttrVal("quantity"));
							}
							
							SItem item = new SItem();
							item.SetType("plurals");
							item.SetName(nameValue);
							item.SetIndex(idxValue);
							item.SetProduct(productValue);
							long l = vNav.getContentFragment();
							if (-1 != l) {
								String valueFromXml = vNav.toRawString((int) (l & 0xFFFFFF), (int) (l >> 32));
								SItem itemFromExcel = items.get(item.GetKey());
								if(null == itemFromExcel){
									continue;
								}
								String valueFromExcel = itemFromExcel.GetTrans().get(lang);
								if((null != valueFromExcel) && !valueFromXml.equals(valueFromExcel)){
									vModifier.removeContent((int) (l & 0xFFFFFF), (int) (l >> 32));
									vModifier.insertAfterHead(valueFromExcel);
									isChange = true;
								}
							}else{
								SItem itemFromExcel = items.get(item.GetKey());
								if(null == itemFromExcel){
									continue;
								}
								String valueFromExcel = itemFromExcel.GetTrans().get(lang);
								if((null != valueFromExcel) && !valueFromExcel.isEmpty()){
									vModifier.insertAfterHead(valueFromExcel);
									isChange = true;
								}
							}
							BookMark bm = new BookMark();
							bm.bind(vNav);
							bm.recordCursorPosition();
							tempBM.set(tempKeys.indexOf(item.GetKey()), bm);
						}
						vNav.pop();
					}
				}
				if (mOpt.isMerge()) {
					//Check string item that does not exist in current file
					for (int i = 0, count = tempBM.size(); i < count; i++) {
						if (tempBM.get(i) == null) {
							int prePos = i;
							while (prePos >= 0) {
								if (tempBM.get(prePos) != null) {
									break;
								}
								prePos--;
							}
							SItem preItem = null;
							if (-1 != prePos) {
								preItem = items.get(tempKeys.get(prePos));
							}
							SItem cntItem = items.get(tempKeys.get(i));
							if ((null != preItem)
								&& (preItem.GetType().equals("string-array")) 
								&& (cntItem.GetType().equals("string-array"))
								&& (preItem.GetName().equals(cntItem.GetName()))
								&& (preItem.GetProduct().equals(cntItem.GetProduct()))){
								String tempBuffer = "";
								for(; (i < count) && (tempBM.get(i) == null); i++){
									SItem tempItem = items.get(tempKeys.get(i));
									if((tempItem.GetType().equals("string-array"))
										&& (preItem.GetName().equals(tempItem.GetName()))
										&& (preItem.GetProduct().equals(tempItem.GetProduct()))){
										String tempValue = tempItem.GetTrans().get(lang);
										if((null != tempValue) && (!tempValue.trim().isEmpty())){
											tempBuffer += String.format("\n        <item>%s</item>", tempValue);
										}
									}else{
										break;
									}
								}
								if(!tempBuffer.isEmpty()){
									tempBM.get(prePos).setCursorPosition();
									vModifier.insertAfterElement(tempBuffer);
									isChange = true;
								}
							}else if ((null != preItem)
									&& (preItem.GetType().equals("plurals")) 
									&& (cntItem.GetType().equals("plurals"))
									&& (preItem.GetName().equals(cntItem.GetName()))
									&& (preItem.GetProduct().equals(cntItem.GetProduct()))){
									String tempBuffer = "";
									for(; (i < count) && (tempBM.get(i) == null); i++){
										SItem tempItem = items.get(tempKeys.get(i));
										if((tempItem.GetType().equals("plurals"))
											&& (preItem.GetName().equals(tempItem.GetName()))
											&& (preItem.GetProduct().equals(tempItem.GetProduct()))){
											String tempIdx = tempItem.GetIndex();
											String tempValue = tempItem.GetTrans().get(lang);
											if((null != tempValue) && (!tempValue.trim().isEmpty())){
												tempBuffer += String.format("\n        <item quantity=\"%s\">%s</item>", tempIdx, tempValue);
											}
										}else{
											break;
										}
									}
									if(!tempBuffer.isEmpty()){
										tempBM.get(prePos).setCursorPosition();
										vModifier.insertAfterElement(tempBuffer);
										isChange = true;
									}
							}else{
								String tempBuffer = "";
								for(; (i < count) && (tempBM.get(i) == null); i++){
									SItem tempItem = items.get(tempKeys.get(i));
									String tempName = tempItem.GetName();
									String tempType = tempItem.GetType();
									String tempProduct = tempItem.GetProduct();
									String tempValue = tempItem.GetTrans().get(lang);
									if(tempItem.GetType().equals("string-array")){
										String tempBuffer2 = "";
										for(; (i < count) && (tempBM.get(i) == null); i++){
											SItem tempItem2 = items.get(tempKeys.get(i));
											String tempName2 = tempItem2.GetName();
											String tempType2 = tempItem2.GetType();
											String tempProduct2 = tempItem2.GetProduct();
											String tempValue2 = tempItem2.GetTrans().get(lang);
											if((tempName.equals(tempName2))
													&& (tempProduct.equals(tempProduct2))
													&& (tempType.equals(tempType2))){
												if((null != tempValue2) && (!tempValue2.trim().isEmpty())){
													tempBuffer2 += String.format("\n        <item>%s</item>", tempValue2);
												}
											}else{
												i--;
												break;
											}
										}
										if(!tempBuffer2.isEmpty()){
											if(!tempProduct.trim().isEmpty()){
												tempBuffer += String.format("\n    <string-array name=\"%s\" product=\"%s\">", tempName, tempProduct);
											}else{
												tempBuffer += String.format("\n    <string-array name=\"%s\">", tempName);
											}
											tempBuffer += tempBuffer2;
											tempBuffer += "\n    </string-array>";
										}
									}else if(tempItem.GetType().equals("plurals")){
										String tempBuffer2 = "";
										for(; (i < count) && (tempBM.get(i) == null); i++){
											SItem tempItem2 = items.get(tempKeys.get(i));
											String tempName2 = tempItem2.GetName();
											String tempType2 = tempItem2.GetType();
											String tempProduct2 = tempItem2.GetProduct();
											String tempIdx2 = tempItem2.GetIndex();
											String tempValue2 = tempItem2.GetTrans().get(lang);
											if((tempName.equals(tempName2))
													&& (tempProduct.equals(tempProduct2))
													&& (tempType.equals(tempType2))){
												if((null != tempValue2) && (!tempValue2.trim().isEmpty())){
													tempBuffer2 += String.format("\n        <item quantity=\"%s\">%s</item>", tempIdx2, tempValue2);
												}
											}else{
												i--;
												break;
											}
										}
										if(!tempBuffer2.isEmpty()){
											if(!tempProduct.trim().isEmpty()){
												tempBuffer += String.format("\n    <plurals name=\"%s\" product=\"%s\">", tempName, tempProduct);
											}else{
												tempBuffer += String.format("\n    <plurals name=\"%s\">", tempName);
											}
											tempBuffer += tempBuffer2;
											tempBuffer += "\n    </plurals>";
										}
									}else if(tempItem.GetType().equals("string")){
										if((null != tempValue) && (!tempValue.trim().isEmpty())){
											if(!tempProduct.trim().isEmpty()){
												tempBuffer += String.format("\n    <string name=\"%s\" product=\"%s\">%s</string>", tempName, tempProduct, tempValue);
											}else{
												tempBuffer += String.format("\n    <string name=\"%s\">%s</string>", tempName, tempValue);
											}
										}
									}
								}
								try {
									if(!tempBuffer.isEmpty()){
										if(-1 == prePos){
											ap.selectXPath("/resources");
											if(-1 != ap.evalXPath()){
												vModifier.insertAfterHead(tempBuffer);
												isChange = true;
											}
										}else{
											if(preItem.GetType().equals("string-array")){
												tempBM.get(prePos).setCursorPosition();
												boolean bItem = vNav.toElement(VTDNav.PARENT);
												if(bItem){
													vModifier.insertAfterElement(tempBuffer);
												}
											}else if(preItem.GetType().equals("plurals")){
												tempBM.get(prePos).setCursorPosition();
												boolean bItem = vNav.toElement(VTDNav.PARENT);
												if(bItem){
													vModifier.insertAfterElement(tempBuffer);
												}
											}else if(preItem.GetType().equals("string")){
												tempBM.get(prePos).setCursorPosition();
												vModifier.insertAfterElement(tempBuffer);
											}
											isChange = true;
										}
									}
								} catch (XPathEvalException e) {
									e.printStackTrace();
								}
							}
						}
					}
				}
				if(isChange){
					//System.out.println(String.format("Update file: %s.", desXmlPath));
					File file = new File(desXmlPath);
					file.getParentFile().mkdirs();
					try {
						FileOutputStream fos = new FileOutputStream(desXmlPath);
						vModifier.output(fos);
						fos.close(); 
					} catch (IOException e) {
						e.printStackTrace();
					} catch (TranscodeException e) {
						e.printStackTrace();
					} 
				}
			} catch (NavException e) {
				e.printStackTrace();
			} catch (XPathParseException e) {
				e.printStackTrace();
			} catch (ModifyException e) {
				e.printStackTrace();
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			}
		}while(false);
	}
	public void Load(ArrayList<String> paths){
		for(int i = 0; i < paths.size(); i++){
			Load(paths.get(i));
		}
	}
	public void Load(String path){
		File file = new File(path);
		if(file.exists()){
			ParserFile(path);
		}
	}
	private void ParserFile(String strFile){
		System.out.println(String.format("Paser file %s", strFile));
		File file = new File(strFile);
		String fileName = file.getName();
		String lang = file.getParentFile().getName();
		String path = file.getParentFile().getParentFile().getPath();
		VTDGen vGen = new VTDGen();
		if (vGen.parseFile(strFile, false)) {
			VTDNav vNav = vGen.getNav();
			AutoPilot ap = new AutoPilot(vNav);
			try {
				if (vNav.toString(vNav.getRootIndex()).equals("resources")) {
					// Add a new pathItem
					mTable.CheckAndAddItem(path);
					SPathItem pathItem = mTable.GetItem(path);
					
					pathItem.CheckAndAddItem(fileName);
					SFileItem fileItem = pathItem.GetItem(fileName);
					fileItem.InsertItemWithSortBegin();

					ap.selectXPath("/resources");
					for (boolean bResItem = vNav.toElement(VTDNav.FIRST_CHILD); false != bResItem; bResItem = vNav.toElement(VTDNav.NEXT_SIBLING)) {
						if (vNav.toString(vNav.getCurrentIndex()).equals("string")) {
							SItem item = new SItem();
							item.SetPath(path);
							item.SetFile(fileName);
							item.SetType("string");
							if (-1 != vNav.getAttrVal("name")) {
								String str = vNav.toString(vNav.getAttrVal("name"));
								item.SetName(str);
							}
							if (-1 != vNav.getAttrVal("msgid")) {
								String str = vNav.toString(vNav.getAttrVal("msgid"));
								item.SetMsgID(str);
							}
							item.SetLength("0");
							item.SetIndex("0");
							if (-1 != vNav.getAttrVal("product")) {
								String str = vNav.toString(vNav.getAttrVal("product"));
								item.SetProduct(str);
							}
							long l = vNav.getContentFragment();
							if (-1 != l) {
								String content = vNav.toRawString((int) (l & 0xFFFFFF), (int) (l >> 32));
								item.GetTrans().put(lang, content);
							}
							fileItem.InsertItemWithSort(item);
						} else if (vNav.toString(vNav.getCurrentIndex()).equals("string-array")) {
							String nameValue = new String();
							if (-1 != vNav.getAttrVal("name")) {
								nameValue = vNav.toString(vNav.getAttrVal("name"));
							}
							String productValue = new String();
							if (-1 != vNav.getAttrVal("product")) {
								productValue = vNav.toString(vNav.getAttrVal("product"));
							}
							//
							int count = 0;
							vNav.push();
							for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item")) {
								count++;
							}
							vNav.pop();

							//
							int idx = 0;
							vNav.push();
							for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item"), idx++) {
								SItem item = new SItem();
								item.SetPath(path);
								item.SetFile(fileName);
								item.SetType("string-array");
								item.SetName(nameValue);
								item.SetLength(String.format("%d", count));
								item.SetIndex(String.format("%d", idx));
								item.SetProduct(productValue);
								long l = vNav.getContentFragment();
								if (-1 != l) {
									String content = vNav.toRawString((int) (l & 0xFFFFFF),(int) (l >> 32));
									item.GetTrans().put(lang, content);
								}
								fileItem.InsertItemWithSort(item);
							}
							vNav.pop();
						} else if (vNav.toString(vNav.getCurrentIndex()).equals("plurals")) {
							String nameValue = new String();
							if (-1 != vNav.getAttrVal("name")) {
								nameValue = vNav.toString(vNav.getAttrVal("name"));
							}
							String productValue = new String();
							if (-1 != vNav.getAttrVal("product")) {
								productValue = vNav.toString(vNav.getAttrVal("product"));
							}
							//
							int count = 0;
							vNav.push();
							for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item")) {
								count++;
							}
							vNav.pop();

							//
							int idx = 0;
							vNav.push();
							for (boolean bItem = vNav.toElement(VTDNav.FIRST_CHILD, "item"); false != bItem; bItem = vNav.toElement(VTDNav.NEXT_SIBLING, "item"), idx++) {
								String idxValue = new String();
								if (-1 != vNav.getAttrVal("quantity")) {
									idxValue = vNav.toString(vNav.getAttrVal("quantity"));
								}
								
								SItem item = new SItem();
								item.SetPath(path);
								item.SetFile(fileName);
								item.SetType("plurals");
								item.SetName(nameValue);
								item.SetLength(String.format("%d", count));
								item.SetIndex(idxValue);
								item.SetProduct(productValue);
								long l = vNav.getContentFragment();
								if (-1 != l) {
									String content = vNav.toRawString((int) (l & 0xFFFFFF),(int) (l >> 32));
									item.GetTrans().put(lang, content);
								}
								fileItem.InsertItemWithSort(item);
							}
							vNav.pop();
						}
					}
					fileItem.InsertItemWithSortEnd();
				}
			} catch (NavException e) {
				e.printStackTrace();
			} catch (XPathParseException e) {
				e.printStackTrace();
			}
		}else{
			System.err.println(String.format("Paser file error: %s", strFile));
		}
	}
	public void Save(){
		ArrayList<String> langs = mTable.GetLangs();
		Save(langs);
	}
	public void Save(ArrayList<String> saveLangs){
		ArrayList<String> langs = mTable.GetLangs();
		ArrayList<String> paths = mTable.GetKeys();
		for(int i = 0, count0 = paths.size(); i < count0; i++){
			String path = paths.get(i);
			SPathItem pathItem = mTable.GetItem(path);
			ArrayList<String> fileNameList = pathItem.GetKeys();
			for(int j = 0, count1 = fileNameList.size(); j < count1; j++){
				for(int k = 0, count2 = langs.size(); k < count2; k++){
					String langName = langs.get(k);
					if(!saveLangs.contains(langName)){
						continue;
					}
					String srcXmlPath = path + File.separator + langs.get(k) + File.separator + fileNameList.get(j);
					String cntDir = System.getProperty("user.dir");
					String desXmlPath = cntDir + File.separator + mOutDir + File.separator + "res" + File.separator + srcXmlPath.substring(3);
					if(mOpt.isWriteBack()){
						CheckAndSave(pathItem, fileNameList.get(j), langName, srcXmlPath, srcXmlPath);
					}else{
						CheckAndSave(pathItem, fileNameList.get(j), langName, desXmlPath, srcXmlPath);
					}
				}
			}
		}
		System.out.println("Successful update xml file!");
	}
	public void Save(String lang){
		ArrayList<String> langs = new ArrayList<String>();
		langs.add(lang);
		Save(langs);
	}
}
