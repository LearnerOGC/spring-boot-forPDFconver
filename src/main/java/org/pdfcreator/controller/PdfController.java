package org.pdfcreator.controller;

import org.pdfcreator.service.PdfFileCreatorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

@RestController
@RequestMapping(value = "/api")
public class PdfController {
	@Autowired
	private PdfFileCreatorService pdfFileCreatorService;

	@RequestMapping(value = "/createpdf", method = RequestMethod.GET)
	public void createPdf(HttpServletResponse response) {
		pdfFileCreatorService.pdfCreator("C:\\D_files\\make_files\\word", "C:\\D_files\\pdf");
	}

	@RequestMapping(value = "/htmltopdf", method = RequestMethod.GET)
	public void createHtmlToPdf(HttpServletResponse response) {
		pdfFileCreatorService.htmlToPdf("C:\\D_files\\make_files\\word", "C:\\D_files\\pdf");
	}

	@RequestMapping(value = "/officetopdf", method = RequestMethod.GET)
	public void createOfficeToPdf(HttpServletResponse response) {
		pdfFileCreatorService.officeToPdf("C:\\D_files\\make_files\\word", "C:\\D_files\\pdf");
	}

	@RequestMapping(value = "/txttopdf", method = RequestMethod.GET)
	public void createTxtToPdf(HttpServletResponse response) {
		pdfFileCreatorService.txtToPdf("C:\\D_files\\make_files\\word", "C:\\D_files\\pdf");
	}
}
