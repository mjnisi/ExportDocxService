package eu.europa.edpb.services;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.Calendar;

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
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.Comments;
import org.docx4j.wml.Comments.Comment;
import org.docx4j.wml.P;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

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


	@Override
	public String exportStatement(String xmlDoc) {

		if(logger.isDebugEnabled()){
			logger.debug("Init");
		}
		
		String EXPORT_TEMPLATE_BVD_FILE = "/templates/StatementEDPB.dotx";

		try {
			
			mlPackage = createDocxFromTemplateVD(EXPORT_TEMPLATE_BVD_FILE);
			
			loadTemplate();
			setMargins();
			
			InputStream isXml = getClass().getResourceAsStream(
					"/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/bill_cl5jc32710002q011gkfo1dto.xml");
			byte[] ba = IOUtils.toByteArray(isXml);

			InputStream isJson = getClass().getResourceAsStream(
					"/documents/Proposal_197_2447391798399561462/PROP_ACT_6020660205072156884/media/annot_bill_cl5jc32710002q011gkfo1dto.xml.json");

			
			JSONObject obj = new JSONObject(isJson);
			JSONArray rows = obj.getJSONArray("rows");
			JSONObject total = obj.getJSONObject("total");
			JSONArray replies = obj.getJSONArray("replies");
			
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

			result = -1;
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");
			mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, "");

			while ((result = ap.evalXPath()) != -1) {
				mlPackage.getMainDocumentPart().addStyledParagraphOfText(STYLE_CONTENT, vn.toString(result));
				
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


			// Comments
			CommentsPart cp = new CommentsPart();
			mlPackage.getMainDocumentPart().addTargetPart(cp);

			Comments comments = factory.createComments();
			cp.setJaxbElement(comments);

			// Add a comment to the comments part
			java.math.BigInteger commentId = BigInteger.valueOf(0);
			Comment theComment = createComment(commentId, "MARIA NISI", null, "my first comment");
			comments.getComment().add(theComment);

			// Add comment reference to document
			// P paraToCommentOn =
			// wordMLPackage.getMainDocumentPart().addParagraphOfText("here is some
			// content");
			P p = new P();

			mlPackage.getMainDocumentPart().getContent().add(p);

			// Create object for commentRangeStart
			CommentRangeStart commentrangestart = factory.createCommentRangeStart();
			commentrangestart.setId(commentId); // substitute your comment id

			// The actual content, in the middle
			p.getContent().add(commentrangestart);

			org.docx4j.wml.Text t = factory.createText();
			t.setValue("hello");

			org.docx4j.wml.R run = factory.createR();
			run.getContent().add(t);

			p.getContent().add(run);

			// Create object for commentRangeEnd
			CommentRangeEnd commentrangeend = factory.createCommentRangeEnd();
			commentrangeend.setId(commentId); // substitute your comment id

			p.getContent().add(commentrangeend);

			p.getContent().add(createRunCommentReference(commentId));

//			System.out.println(mlPackage.getMainDocumentPart().getXML());
//
//			// ++, for next comment ...
//			commentId = commentId.add(java.math.BigInteger.ONE);

			System.out.println(mlPackage.getMainDocumentPart().getContent());

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

	
	private void setMargins() {
		mlPackage.getMainDocumentPart().getJaxbElement().getBody().getSectPr().getPgMar().setTop(new BigInteger("0"));

	}
	
	
//	Instanciación del documento:
//	template de word donde se definen los estilos

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
