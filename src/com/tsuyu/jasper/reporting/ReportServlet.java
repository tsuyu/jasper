package com.tsuyu.jasper.reporting;

import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.tsuyu.jasper.db.Db;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRExporter;
import net.sf.jasperreports.engine.JRExporterParameter;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.design.JRDesignQuery;
import net.sf.jasperreports.engine.design.JasperDesign;
import net.sf.jasperreports.engine.export.JRCsvExporter;
import net.sf.jasperreports.engine.export.JRCsvExporterParameter;
import net.sf.jasperreports.engine.export.JRHtmlExporter;
import net.sf.jasperreports.engine.export.JRPdfExporter;
import net.sf.jasperreports.engine.export.JRRtfExporter;
import net.sf.jasperreports.engine.export.JRTextExporter;
import net.sf.jasperreports.engine.export.JRTextExporterParameter;
import net.sf.jasperreports.engine.export.JRXlsExporter;
import net.sf.jasperreports.engine.export.JRXlsExporterParameter;
import net.sf.jasperreports.engine.export.oasis.JROdsExporter;
import net.sf.jasperreports.engine.export.oasis.JROdtExporter;
import net.sf.jasperreports.engine.export.ooxml.JRDocxExporter;
import net.sf.jasperreports.engine.xml.JRXmlLoader;

@SuppressWarnings("serial")
public class ReportServlet extends HttpServlet {

	@Override
	public void init() throws ServletException {

	}

	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {

		Connection conn = Db.connectMysql();

		JasperDesign jasperDesign;
		try {
			jasperDesign = JRXmlLoader
					.load("D:\\javworkspace\\strutsworkspace\\jasper\\src\\customer.jrxml");

			JRDesignQuery query = new JRDesignQuery();

			query.setText("SELECT `customer`.`first_name` AS customer_first_name, "
					+ "`customer`.`last_name` AS customer_last_name, "
					+ "`customer`.`email` AS customer_email "
					+ "FROM  `customer` customer");

			jasperDesign.setQuery(query);

			JasperReport report = JasperCompileManager
					.compileReport(jasperDesign);

			HashMap<String, Object> params = new HashMap<String, Object>();
			params.put("title", "Customer");

			JasperPrint jasperPrint = JasperFillManager.fillReport(report,
					params, conn);

			ReportServlet.format(request, response, jasperPrint);

			conn.close();
		} catch (JRException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void format(HttpServletRequest request,
			HttpServletResponse response, JasperPrint jasperPrint)
			throws JRException, IOException {

		String type = request.getParameter("format");
		OutputStream ouputStream = response.getOutputStream();

		JRExporter exporter = null;

		if (type.equals("xls")) {

			response.setContentType("application/vnd.ms-excel");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.xls\"");

			exporter = new JRXlsExporter();
			exporter.setParameter(JRXlsExporterParameter.IS_ONE_PAGE_PER_SHEET,
					true);
			exporter.setParameter(
					JRXlsExporterParameter.IS_WHITE_PAGE_BACKGROUND, false);
			exporter.setParameter(
					JRXlsExporterParameter.IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS,
					true);
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);

		} else if (type.equals("csv")) {

			response.setContentType("application/csv");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.csv\"");

			exporter = new JRCsvExporter();
			exporter.setParameter(JRCsvExporterParameter.FIELD_DELIMITER, ",");
			exporter.setParameter(JRCsvExporterParameter.RECORD_DELIMITER, "\n");
			exporter.setParameter(JRCsvExporterParameter.CHARACTER_ENCODING,
					"UTF-8");
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
		} else if (type.equals("docx")) {
			
			response.setContentType("application/vnd.ms-word");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.docx\"");
			
			exporter = new JRDocxExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
		} else if (type.equals("html")) {

			response.setContentType("application/html");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.html\"");

			exporter = new JRHtmlExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
		} else if (type.equals("pdf")) {

			response.setContentType("application/pdf");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.pdf\"");

			exporter = new JRPdfExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);

		} else if (type.equals("ods")) {
			
			response.setContentType("vnd.oasis.opendocument.spreadsheet");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.ods\"");

			exporter = new JROdsExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);

			exporter.exportReport();
		} else if (type.equals("odt")) {

			response.setContentType("vnd.oasis.opendocument.text");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.odt\"");
			
			exporter = new JROdtExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
		} else if (type.equals("txt")) {

			response.setContentType("application/txt");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.txt\"");

			exporter = new JRTextExporter();
			exporter.setParameter(JRTextExporterParameter.PAGE_WIDTH, 120);
			exporter.setParameter(JRTextExporterParameter.PAGE_HEIGHT, 60);
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
			exporter.exportReport();

		} else if (type.equals("rtf")) {

			response.setContentType("application/rtf");
			response.setHeader("Content-Disposition",
					"inline; filename=\"report.rtf\"");
			exporter = new JRRtfExporter();
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM,
					ouputStream);
		}

		exporter.exportReport();

	}

	public void destroy() {

	}
}