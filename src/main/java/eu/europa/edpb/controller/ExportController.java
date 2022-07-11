package eu.europa.edpb.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import eu.europa.edpb.services.ExportService;

@Controller
@RequestMapping("/export")
public class ExportController {
	
	@Autowired
	private ExportService exportService;

	@GetMapping(value= "/docx")
	public String exportStatement(Model model) {

//		model.addAttribute("doc", model);
		
		exportService.exportStatement(null);
		
		
		
		return "view-docx";
	}
	
		
	
}
