package egovframework.rte.fdl.excel.download;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.POIXMLDocumentPart;
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

import egovframework.rte.fdl.excel.util.AbstractPOIExcelView;

public class CategoryExcelView extends AbstractPOIExcelView {

	private static final Logger LOGGER = LoggerFactory
			.getLogger(CategoryExcelView.class);

	@Override
	protected void buildExcelDocument(Map model, XSSFWorkbook wb,
			HttpServletRequest req, HttpServletResponse resp) throws Exception {

		resp.setHeader("Content-disposition", "attachment;filename=test.xls");
		resp.setHeader("Content-Type",
				"application/vnd.ms-excel; charset=MS949");
		resp.setHeader("Content-Description", "JSP Generated Data");
		resp.setHeader("Content-Transfer-Encoding", "binary;");
		resp.setHeader("Pragma", "no-cache;");
		resp.setHeader("Expires", "-1;");

		XSSFCell cell = null;

		LOGGER.debug("### buildExcelDocument start !!!");

		XSSFSheet sheet = wb.createSheet("User List");
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

		Map<String, Object> map = (Map<String, Object>) model
				.get("categoryMap");
		List<Object> categories = (List<Object>) map.get("category");

		boolean isVO = false;

		if (categories.size() > 0) {
			Object obj = categories.get(0);
			isVO = obj instanceof UsersVO;
		}

		for (int i = 0; i < categories.size(); i++) {

			if (isVO) { // VO

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

			} else { // Map

				LOGGER.debug("### buildExcelDocument Map : {} started!!", i);

				Map<String, String> category = (Map<String, String>) categories
						.get(i);

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

		Cell cell1 = getCell(sheet, 2, 0);
		cell1.setCellValue("             test1\ntest2");
		CellStyle cellStyle = wb.createCellStyle();

		XSSFCellStyle cs = (XSSFCellStyle) cellStyle;
		cs.setWrapText(true);
		StylesTable table = null;
		for (POIXMLDocumentPart part : ((XSSFWorkbook) wb).getRelations()) {
			if (part instanceof StylesTable) {
				table = (StylesTable) part;
				break;
			}
		}

		CTBorder ct = CTBorder.Factory.newInstance();

		ct.setDiagonalDown(true);

		CTBorderPr pr = ct.isSetDiagonal() ? ct.getDiagonal() : ct
				.addNewDiagonal();

		org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.Enum borderStyle = STBorderStyle.Enum
				.forString("thin");
		pr.setStyle(borderStyle);
		ct.setDiagonal(pr);

		int idx = table.putBorder(new XSSFCellBorder(ct, null));

		cs.getCoreXf().setBorderId(idx);
		cs.getCoreXf().setApplyBorder(true);

		cell1.setCellStyle(cellStyle);
	}
}
