package eu.europa.edpb.services;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

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
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.Comments;
import org.docx4j.wml.Comments.Comment;
import org.docx4j.wml.P;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

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
//		String EXPORT_TEMPLATE_BVD_FILE = "/templates/StatementEDPBcomment3.dotx";

		
		String jsonPath = "/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/media/annot_bill_cl5jc32710002q011gkfo1dto.xml.json";
		String xmlPath = "/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/bill_cl5jc32710002q011gkfo1dto.xml";

		try {
			
			mlPackage = createDocxFromTemplateVD(EXPORT_TEMPLATE_BVD_FILE);
			
 			loadTemplate();
			setMargins();

			InputStream isXml = getClass().getResourceAsStream(xmlPath);
			byte[] ba = IOUtils.toByteArray(isXml);
	         
						
			List<HashMap<String, String>> annotations = getAnnotationsFromFile(jsonPath);
	
			
			// Comments
			CommentsPart cp = new CommentsPart();
			mlPackage.getMainDocumentPart().addTargetPart(cp);

			Comments comments = factory.createComments();
			cp.setJaxbElement(comments);

			
			VTDGen vg = new VTDGen();
			vg.setDoc(ba);
			vg.parse(false);

			AutoPilot ap = new AutoPilot();

			// Title
			ap.selectXPath("/akomaNtoso/bill/coverPage/longTitle/p//text()");

			VTDNav vn = vg.getNav();
			ap.bind(vn);

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
			ap.selectXPath("//body/paragraph/content/p/text()");
			//buscar el id de p y buscarlo dentro del array de id comments para ver si hay q crear un commentario.
			
			
			result = -1;
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");

			while ((result = ap.evalXPath()) != -1) {
				
				String[] positions = {"10:15"};
				addCommentToP(vn.toString(result), STYLE_CONTENT , positions);
//				addCommentToP1(id, vn.toString(result), STYLE_CONTENT , annotations);

				if(logger.isDebugEnabled()){
					logger.debug("Paragraph : " + vn.toString(result));
				}
			}

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
		logger.debug("json :"+json);

		JSONObject obj = new JSONObject(json);
		JSONArray rows = obj.getJSONArray("rows");
		BigInteger total = obj.getBigInteger("total");
		JSONArray replies = obj.getJSONArray("replies");

		logger.debug("rows :"+rows);
		logger.debug("total :"+total);
		logger.debug("replies :"+replies);

		String jsonPathComments = "$.rows[*].target[*].selector[0]";
		//buscar tb los start y end, hacer un metodo q reciba parrafo, inicio y fin y cree un comentario. Devuelva un id. 
	
		DocumentContext jsonContext = JsonPath.parse(json);
		List<HashMap<String, String>> jsonAnnotations = jsonContext.read(jsonPathComments);
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

	
	
	private void addCommentToP(String text, String style, String[] positions ) {
		
		P p = mlPackage.getMainDocumentPart().createStyledParagraphOfText(style, "");
		CTLanguage lang = new CTLanguage();
		lang.setVal("fr-BE");

		for (String pos : positions) {
			int startPosition = Integer.parseInt(pos.split(":")[0]);
			int endPosition = Integer.parseInt(pos.split(":")[1]);

			commentId = commentId.add(java.math.BigInteger.ONE);
			
			org.docx4j.wml.R wR = factory.createR();
			org.docx4j.wml.RPr wRPr = factory.createRPr();
			org.docx4j.wml.Text wT = factory.createText();
			

			wT.setValue(text.substring(0, startPosition));
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
			

			wT1.setValue(text.substring(startPosition, endPosition));
			wT1.setSpace("preserve");
			
			wRPr1.setLang(lang);
			wR1.setRPr(wRPr1);
			wR1.getContent().add(wT1);

			p.getContent().add(wR1);
			
			
			// Create object for commentRangeEnd
			CommentRangeEnd commentrangeend = factory.createCommentRangeEnd();
			commentrangeend.setId(commentId); // substitute your comment id

			p.getContent().add(commentrangeend);

			org.docx4j.wml.R wR2 = factory.createR();
			org.docx4j.wml.RPr wRPr2 = factory.createRPr();
			org.docx4j.wml.Text wT2 = factory.createText();
			
			wT2.setValue(text.substring(endPosition));
			wT2.setSpace("preserve");

			wRPr2.setLang(lang);
			wR2.setRPr(wRPr2);
			wR2.getContent().add(wT2);
			
			p.getContent().add(wR2);


			p.getContent().add(createRunCommentReference(commentId));

			Comment theComment = createComment(commentId, "MARIA NISI", null, "my first comment");

			mlPackage.getMainDocumentPart().getCommentsPart().getJaxbElement().getComment().add(theComment);

		}
		mlPackage.getMainDocumentPart().getContent().add(p);
	

	}
	
	
	
	
	private org.docx4j.wml.Comments.Comment createComment(java.math.BigInteger commentId, String author, Calendar date,
			String message) {

		org.docx4j.wml.Comments.Comment comment = factory.createCommentsComment();
		comment.setId(commentId);
		if (author != null) {
			comment.setAuthor(author);
		}
		if (date != null) {
//			String dateString = RFC3339_FORMAT.format(date.getTime()) ;	
//			comment.setDate(value)
			// TODO - at present this is XMLGregorianCalendar
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
