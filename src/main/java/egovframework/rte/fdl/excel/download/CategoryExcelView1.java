package egovframework.rte.fdl.excel.download;

import java.awt.Color;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.record.CFRuleRecord.ComparisonOperator;
import org.apache.poi.hssf.record.cf.BorderFormatting;
import org.apache.poi.hssf.usermodel.EscherGraphics;
import org.apache.poi.hssf.usermodel.EscherGraphics2d;
import org.apache.poi.hssf.usermodel.HSSFBorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFConditionalFormattingRule;
import org.apache.poi.hssf.usermodel.HSSFFontFormatting;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSheetConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.servlet.view.document.AbstractExcelView;

import egovframework.rte.fdl.excel.util.AbstractPOIExcelView;

public class CategoryExcelView1 extends AbstractExcelView {
	 
	private static final Logger LOGGER  = LoggerFactory.getLogger(CategoryExcelView.class);
 
	@Override
	protected void buildExcelDocument(Map model, HSSFWorkbook wb, HttpServletRequest req, HttpServletResponse resp) throws Exception {
        
 
        LOGGER.debug("### buildExcelDocument start !!!");
        resp.setHeader("Content-disposition", "attachment;filename=test.xls");
		resp.setHeader("Content-Type",
				"application/vnd.ms-excel; charset=MS949");
		resp.setHeader("Content-Description", "JSP Generated Data");
		resp.setHeader("Content-Transfer-Encoding", "binary;");
		resp.setHeader("Pragma", "no-cache;");
		resp.setHeader("Expires", "-1;");
		
		HSSFCell cell = null;
        HSSFSheet sheet = wb.createSheet("User List");
        sheet.setDefaultColumnWidth(12);
 
        // put text in first cell
        cell = getCell(sheet, 0, 0);
        setText(cell, "User List");
 
        // set header information
        setText(getCell(sheet, 2, 0), "id");
        setText(getCell(sheet, 2, 1), "name");
        setText(getCell(sheet, 2, 2), "description");
        setText(getCell(sheet, 2, 3), "use_yn");
        setText(getCell(sheet, 2, 4), "reg_user");
 
        LOGGER.debug("### buildExcelDocument cast");
 
 
        Map<String, Object> map= (Map<String, Object>) model.get("categoryMap");
        List<Object> categories = (List<Object>) map.get("category");
 
        boolean isVO = false;
 
        if (categories.size() > 0) {
        	Object obj = categories.get(0);
        	isVO = obj instanceof UsersVO;
        }
 
        for (int i = 0; i < categories.size(); i++) {
 
        	if (isVO) {	// VO
 
        		LOGGER.debug("### buildExcelDocument VO : {} started!!", i);
 
        		UsersVO category = (UsersVO) categories.get(i);
 
	            cell = getCell(sheet, 3 + i, 0);
	            setText(cell, category.getId());
 
	            cell = getCell(sheet, 3 + i, 1);
	            setText(cell, category.getName());
 
	            cell = getCell(sheet, 3 + i, 2);
	            setText(cell, category.getDescription());
 
	            cell = getCell(sheet, 3 + i, 3);
	            setText(cell, category.getUseyn());
 
	            cell = getCell(sheet, 3 + i, 4);
	            setText(cell, category.getReguser());
 
	            LOGGER.debug("### buildExcelDocument VO : {} end!!", i);
 
        	 } else {	// Map
 
        		LOGGER.debug("### buildExcelDocument Map : {} started!!", i);
 
        		Map<String, String> category = (Map<String, String>) categories.get(i);
 
 	            cell = getCell(sheet, 3 + i, 0);
 	            setText(cell, category.get("id"));
 
 	            cell = getCell(sheet, 3 + i, 1);
 	            setText(cell, category.get("name"));
 
 	            cell = getCell(sheet, 3 + i, 2);
 	            setText(cell, category.get("description"));
 
 	            cell = getCell(sheet, 3 + i, 3);
 	            setText(cell, category.get("useyn"));
 
 	            cell = getCell(sheet, 3 + i, 4);
 	            setText(cell, category.get("reguser"));
 
 	            LOGGER.debug("### buildExcelDocument Map : {} end!!", i);
        	 }
        }
        
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        HSSFClientAnchor a = new HSSFClientAnchor(0, 0, 0, 0, (short)3, 2, (short)4, 3);
        //0, 0, 0, 0, 시작위치x, 시작위치y, 선x, 선y
        HSSFShapeGroup group = patriarch.createGroup(a);
        group.setCoordinates(0, 0, 320, 276);
        float verticalPointsPerPixel = a.getAnchorHeightInPoints(sheet) / Math.abs(group.getY2() - group.getY1());
        EscherGraphics g = new EscherGraphics(group, wb, Color.black, verticalPointsPerPixel);
        EscherGraphics2d g2d = new EscherGraphics2d(g);
        g2d.drawLine(0, 0, 320, 276);
    }
}
