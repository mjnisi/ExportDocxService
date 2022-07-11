package eu.europa.edpb.services;


public interface ExportService {
	
	public String exportStatement(String xmlDoc);
	
	public String exportLetter(String xmlDoc);
	


}
