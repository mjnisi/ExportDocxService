package eu.europa.edpb.services;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import javax.xml.datatype.DatatypeConfigurationException;
import javax.xml.datatype.DatatypeFactory;
import javax.xml.datatype.XMLGregorianCalendar;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.log4j.Logger;
import org.docx4j.Docx4J;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.CommentsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.Comments;
import org.docx4j.wml.Comments.Comment;
import org.docx4j.wml.Highlight;
import org.docx4j.wml.P;
import org.docx4j.wml.RStyle;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

import com.fasterxml.jackson.databind.ser.std.CalendarSerializer;
import com.jayway.jsonpath.DocumentContext;
import com.jayway.jsonpath.JsonPath;
import com.ximpleware.AutoPilot;
import com.ximpleware.NavException;
import com.ximpleware.ParseException;
import com.ximpleware.VTDGen;
import com.ximpleware.VTDNav;
import com.ximpleware.XPathEvalException;
import com.ximpleware.XPathParseException;

@Service
public class ExportServiceImpl implements ExportService {

	final static String STYLE_TITLE = "StyleTitle";
	final static String STYLE_TITLE_SPACE = "StyleTitleSpace";
	final static String STYLE_ADOPTED = "StyleAdopted";
	final static String STYLE_FORMULA = "StyleFormula";
	final static String STYLE_FORMULA_SPACE = "StyleFormulaSpace";
	final static String STYLE_CONTENT = "StyleContent";
	final static String STYLE_CONCLUTION = "StyleConclution";

	private ObjectFactory factory = Context.getWmlObjectFactory();
	private WordprocessingMLPackage mlPackage;
	
	final static Logger logger = Logger.getLogger(ExportServiceImpl.class);

	private java.math.BigInteger commentId = BigInteger.valueOf(0);
				

	@Override
	public String exportStatement(String xmlDoc) {

		if(logger.isDebugEnabled()){
			logger.debug("Init");
		}
		
		String EXPORT_TEMPLATE_BVD_FILE = "/templates/StatementEDPB.dotx";

		
		String jsonPath = "/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/media/annot_bill_cl5jc32710002q011gkfo1dto.xml.json";
		String xmlPath = "/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/bill_cl5jc32710002q011gkfo1dto.xml";

		try {
			
			mlPackage = createDocxFromTemplateVD(EXPORT_TEMPLATE_BVD_FILE);
			
 			loadTemplate();
			setMargins();

			InputStream isXml = getClass().getResourceAsStream(xmlPath);
			byte[] ba = IOUtils.toByteArray(isXml);
	         
						
			List<HashMap<String, String>> annotations = getAnnotationsFromFile(jsonPath);
	
			if (annotations != null ) {
				// Create CommentsPart
				CommentsPart cp = new CommentsPart();
				mlPackage.getMainDocumentPart().addTargetPart(cp);

				Comments comments = factory.createComments();
				cp.setJaxbElement(comments);
			}


			
			VTDGen vg = new VTDGen();
			vg.setDoc(ba);
			vg.parse(false);

			AutoPilot ap = new AutoPilot();
			AutoPilot ap2 = new AutoPilot();

			// Title
			ap.selectXPath("/akomaNtoso/bill/coverPage/longTitle/p//text()");

			VTDNav vn = vg.getNav();
			ap.bind(vn);
			ap2.bind(vn);


			int result = -1;
			String title = "";
			while ((result = ap.evalXPath()) != -1) {
				if (vn.toString(result).trim() != "") {
					title += vn.toString(result) + " ";
				}
			}
			
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_TITLE, title);
			if(logger.isDebugEnabled()){
				logger.debug("Title : " + title);
			}

			// Preface
			ap.resetXPath();
			ap.bind(vn);
			ap.selectXPath("//preamble/formula/p/text()");

			result = -1;
			result = ap.evalXPath();

			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_TITLE_SPACE, "");
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_TITLE_SPACE, "");

			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_FORMULA, vn.toString(result));
			if(logger.isDebugEnabled()){
				logger.debug("Formula : " + vn.toString(result));
			}

			// Body
			ap.resetXPath();
			ap.bind(vn);
			ap.declareXPathNameSpace("xml", "https://www.w3.org/2001/xml.xsd");
			ap2.declareXPathNameSpace("xml", "https://www.w3.org/2001/xml.xsd");

			
			ap.bind(vn);
			ap2.bind(vn);

			
			ap.selectXPath("//body/paragraph/content/p/@id");

			//buscar el id de p y buscarlo dentro del array de id comments para ver si hay q crear un commentario.
			
			
			result = -1;
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");

			
		    int count = 0;

		    // iterate over all file IDs
		    while ((result = ap.evalXPath()) != -1) {
		      int j;

		      // retrieve the value of the id attribute field
		      String attributeName = vn.toString(result);
		      int attributeId = vn.getAttrVal("id");
		      String attributeVal = vn.toString(attributeId);
			  
		      logger.debug(" attributeName ==> " + attributeName + " attributeId ==> " + attributeId + " attributeVal ==> " + attributeVal);

			  String path = "//body/paragraph/content/p[@id='"+ attributeVal +"']/text()";
		      ap2.selectXPath(path);
		      
		      while ((j = ap2.evalXPath()) != -1) {
		    	  
		    	String pId =  attributeVal;
		    	String p = vn.toString(j);
			    logger.debug("Paragraph num "+ ++count + " ID ==> " + pId);
		        
		        ArrayList<HashMap<String, String>> filteredAnn = filterAnnotations(pId, annotations);
		        addCommentToP(p, STYLE_CONTENT , filteredAnn);
		        
				mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");

		      }
		      ap2.resetXPath();
		    }
		    ap.resetXPath();
		    
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");

			// Conclusion
			ap.resetXPath();

			ap.bind(vn);
			ap.selectXPath("//signature//text()");

			result = -1;

			while ((result = ap.evalXPath()) != -1) {
				mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONCLUTION, vn.toString(result));
				
				if(logger.isDebugEnabled()){
					logger.debug("Conclution : " + vn.toString(result));
				}
			}

			System.out.println(mlPackage.getMainDocumentPart().getXML());

			mlPackage.save(new java.io.File("doc2.doc"), Docx4J.FLAG_SAVE_ZIP_FILE);

		} catch (Docx4JException | IOException | ParseException | XPathParseException | XPathEvalException
				| NavException e) {
			e.printStackTrace();
		}

		return null;
	}

	@Override
	public String exportLetter(String xmlDoc) {
		// TODO Auto-generated method stub
		return null;
	}

	
	private List<HashMap<String, String>> getAnnotationsFromFile(String jsonPath) {
		
		String json = getResource(jsonPath);
//
		JSONObject obj = new JSONObject(json);
		JSONArray rows = obj.getJSONArray("rows");

		List<HashMap<String, String>> jsonAnnotations = new ArrayList<HashMap<String, String>>();
		
		
		for (int i = 0; i < rows.length(); i++) {

			String jsonPathSelector = "$.rows["+i+"].target[*].selector[0]";
			String jsonPathAuthor = "$.rows["+i+"].user_info.display_name";
			String jsonPathCreationDate = "$.rows["+i+"].created";
			String jsonPathComment = "$.rows["+i+"].text";
			String jsonPathTag = "$.rows["+i+"].tags[0]";


			DocumentContext jsonContext = JsonPath.parse(json);

			List<HashMap<String, String>> selectors = jsonContext.read(jsonPathSelector);
			String author = jsonContext.read(jsonPathAuthor);
			String creationDate = jsonContext.read(jsonPathCreationDate);
			String comment = jsonContext.read(jsonPathComment);
			String tag = jsonContext.read(jsonPathTag);

			HashMap<String, String> selector = selectors.get(0);
			selector.put("author", author);
			selector.put("creationDate", creationDate);
			selector.put("comment", comment);
			selector.put("tag", tag);

			jsonAnnotations.add(selector);
		}
		
		logger.debug("jsonComments :" + jsonAnnotations);
		return jsonAnnotations;
		
	}

	
	
	private String getResource(String resource) {
        StringBuilder json = new StringBuilder();
        try {
            BufferedReader in = new BufferedReader(
                    new InputStreamReader(Objects.requireNonNull(getClass().getResourceAsStream(resource)),
                            StandardCharsets.UTF_8));
            String str;
            while ((str = in.readLine()) != null)
                json.append(str);
            in.close();
        } catch (IOException e) {
            throw new RuntimeException("Caught exception reading resource " + resource, e);
        }
        return json.toString();
    }

	
	
	private void setMargins() {
		mlPackage.getMainDocumentPart().getJaxbElement().getBody().getSectPr().getPgMar().setTop(new BigInteger("0"));
		//make those margins work only for the first page
	}
	
	
	private WordprocessingMLPackage createDocxFromTemplateVD(String template) {
		WordprocessingMLPackage wordPackage = null;
		try {
			InputStream is = getTemplate(template);
			wordPackage = WordprocessingMLPackage.load(is);
			
		} catch (Docx4JException e) {
			logger.error("Error: " + e.getMessage());
		}
		
		if(logger.isDebugEnabled()){
			logger.debug("Template loaded.");
		}
		return wordPackage;
	}

	private InputStream getTemplate(String key) {
		InputStream is = getClass().getResourceAsStream(key);
		if (is != null) {
			return is;
		}
		return null;
	}

	private MainDocumentPart loadTemplate() throws Docx4JException {

		MainDocumentPart mainDocumentPart = mlPackage.getMainDocumentPart();

		try {
			VariablePrepare.prepare(mlPackage);

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return mainDocumentPart;

	}	
	
	private ArrayList<HashMap<String, String>> filterAnnotations(String pId, List<HashMap<String, String>> annotations) {
		ArrayList<HashMap<String, String>> pAnn = new ArrayList<HashMap<String, String>>();
		
		annotations.forEach(i -> {
			if(pId.equals(i.get("id"))) {
				pAnn.add(i);
			}
		});
		
		return pAnn;
	}
	
	private void addCommentToP(String text, String style, ArrayList<HashMap<String, String>> pAnnotations ) {
		
		P p = mlPackage.getMainDocumentPart().createStyledParagraphOfText(style, "");
		CTLanguage lang = new CTLanguage();
		lang.setVal("fr-BE");

		pAnnotations.forEach(System.out::println);
			
		int i = 0;
		for (HashMap<String, String> hashMap : pAnnotations) {
		
			String pId = hashMap.get("id");
			int startPosition =(Integer) ((Object)hashMap.get("start"));  
			int endPosition = (Integer) ((Object)hashMap.get("end"));  

			String author = hashMap.get("author");
			String creationDate = hashMap.get("creationDate");
			String textComment = (hashMap.get("comment")).replaceAll("<.*?>" , " ");
			String tag = hashMap.get("tag");//"protected";
//			String tag = "protected";

			

			commentId = commentId.add(java.math.BigInteger.ONE);
			
			org.docx4j.wml.R wR = factory.createR();
			org.docx4j.wml.RPr wRPr = factory.createRPr();
			org.docx4j.wml.Text wT = factory.createText();
			

			wT.setValue(text.substring(i, startPosition));
			wT.setSpace("preserve");

			wRPr.setLang(lang);
			wR.setRPr(wRPr);
			wR.getContent().add(wT);
			p.getContent().add(wR);
		
			// Create object for commentRangeStart
			CommentRangeStart commentrangestart = factory.createCommentRangeStart();
			commentrangestart.setId(commentId); 

			p.getContent().add(commentrangestart);

			org.docx4j.wml.R wR1 = factory.createR();
			org.docx4j.wml.RPr wRPr1 = factory.createRPr();
			org.docx4j.wml.Text wT1 = factory.createText();
			
			if(tag.equals("suggestion")) {//protected
				wT1.setValue("[Depersonalised]");
				Highlight highlight = factory.createHighlight();
				highlight.setVal("lightGray");
				wRPr1.setHighlight(highlight);
				
			} else {
				wT1.setValue(text.substring(startPosition, endPosition));				

			}

			wT1.setSpace("preserve");
			
			wRPr1.setLang(lang);
			wR1.setRPr(wRPr1);
			wR1.getContent().add(wT1);
			
			p.getContent().add(wR1);
			
			
			// Create object for commentRangeEnd
			CommentRangeEnd commentrangeend = factory.createCommentRangeEnd();
			commentrangeend.setId(commentId);

			p.getContent().add(commentrangeend);
			
			i = endPosition;
			
			if(pAnnotations.lastIndexOf(hashMap)==pAnnotations.size()-1) {

				org.docx4j.wml.R wR2 = factory.createR();
				org.docx4j.wml.RPr wRPr2 = factory.createRPr();
				org.docx4j.wml.Text wT2 = factory.createText();
				
				wT2.setValue(text.substring(endPosition));
				wT2.setSpace("preserve");
	
				wRPr2.setLang(lang);
				wR2.setRPr(wRPr2);
				wR2.getContent().add(wT2);
				
				p.getContent().add(wR2);

			}

			p.getContent().add(createRunCommentReference(commentId));

			Comment theComment = createComment(commentId, author, creationDate, textComment);

			mlPackage.getMainDocumentPart().getCommentsPart().getJaxbElement().getComment().add(theComment);

		}
		mlPackage.getMainDocumentPart().getContent().add(p);
	

	}

	
	
	private org.docx4j.wml.Comments.Comment createComment(java.math.BigInteger commentId, String author, String date,
			String message) {

		org.docx4j.wml.Comments.Comment comment = factory.createCommentsComment();
		comment.setId(commentId);
		if (author != null) {
			comment.setAuthor(author);
		}
		if (date != null) {
			try {
				XMLGregorianCalendar result = DatatypeFactory.newInstance().newXMLGregorianCalendar(date);
				comment.setDate(result);

			} catch (DatatypeConfigurationException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		org.docx4j.wml.P commentP = factory.createP();
		comment.getEGBlockLevelElts().add(commentP);
		org.docx4j.wml.R commentR = factory.createR();
		commentP.getContent().add(commentR);
		org.docx4j.wml.Text commentText = factory.createText();
		commentR.getContent().add(commentText);

		commentText.setValue(message);

		return comment;
	}

	private org.docx4j.wml.R createRunCommentReference(java.math.BigInteger commentId) {

		org.docx4j.wml.R run = factory.createR();
		org.docx4j.wml.R.CommentReference commentRef = factory.createRCommentReference();
		run.getContent().add(commentRef);
		commentRef.setId(commentId);

		return run;

	}

}
