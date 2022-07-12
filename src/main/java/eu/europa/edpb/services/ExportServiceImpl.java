package eu.europa.edpb.services;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;

import org.apache.commons.compress.utils.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.ObjectFactory;
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

	@Override
	public String exportStatement(String xmlDoc) {
		
		String EXPORT_TEMPLATE_BVD_FILE = "/templates/Statement.dotx";


		WordprocessingMLPackage mlPackage = createDocxFromTemplateVD(EXPORT_TEMPLATE_BVD_FILE);

		try {
//			Con el objeto devuelto (MainDocumentPart), ya empezaríamos a realizar las inserciones en el 
//			futuro documento docx con el siguiente método: 			mainDocumentPart.addStyledParagraphOfText("LEOSCONCLUSIONPRES", conclusionPresidente);
//			Aquí tener en cuenta que el primer parámetro es un estilo que tendremos creado en la plantilla de Word a usar. 
//			El según parámetro es el String a insertar en el documento docx con el estilo comentado.

			MainDocumentPart mainDocumentPart = loadTemplate(mlPackage);
			mainDocumentPart.getJaxbElement().getBody().getSectPr().getPgMar().setTop(new BigInteger("0"));
			
			
			InputStream is = getClass().getResourceAsStream("/documents/bill_cl59nigaz00023o113tmadl7l.xml");
			byte[] ba = IOUtils.toByteArray(is);

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
			System.out.println("Title: " + title);
			mainDocumentPart.addStyledParagraphOfText(STYLE_TITLE, title);

			// Preface
			ap.resetXPath();
			ap.bind(vn);
			ap.selectXPath("//preamble/formula/p/text()");

			result = -1;
			result = ap.evalXPath();
			if (result != -1) {
				System.out.println("formula: " + vn.toString(result));
			}
			mainDocumentPart.addStyledParagraphOfText(STYLE_TITLE_SPACE, "");
			mainDocumentPart.addStyledParagraphOfText(STYLE_TITLE_SPACE, "");

			mainDocumentPart.addStyledParagraphOfText(STYLE_FORMULA, vn.toString(result));

			// Body
			ap.resetXPath();
			ap.bind(vn);
			ap.selectXPath("//body/paragraph/content/p/text()");

			result = -1;

			mainDocumentPart.addStyledParagraphOfText(STYLE_CONTENT, "");
			mainDocumentPart.addStyledParagraphOfText(STYLE_CONTENT, "");

			while ((result = ap.evalXPath()) != -1) {
				System.out.println("Paragraph: " + vn.toString(result));
				mainDocumentPart.addStyledParagraphOfText(STYLE_CONTENT, vn.toString(result));
			}

			mainDocumentPart.addStyledParagraphOfText(STYLE_CONTENT, "");
			mainDocumentPart.addStyledParagraphOfText(STYLE_CONTENT, "");

			// Conclusion
			ap.resetXPath();

			ap.bind(vn);
			ap.selectXPath("//signature//text()");

			result = -1;

			while ((result = ap.evalXPath()) != -1) {
				System.out.println("Conclution: " + vn.toString(result));
				mainDocumentPart.addStyledParagraphOfText(STYLE_CONCLUTION, vn.toString(result));
			}

			System.out.println(mainDocumentPart.getXML());
			
			
			// .. content type
			mlPackage.getContentTypeManager().addDefaultContentType("html", "text/html");

			mlPackage.save(new java.io.File("helloworld0.docx"), Docx4J.FLAG_SAVE_ZIP_FILE);

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


//	Instanciación del documento:
//	template de word donde se definen los estilos

	private WordprocessingMLPackage createDocxFromTemplateVD(String template) {
		WordprocessingMLPackage wordPackage = null;
		try {
			InputStream is = getTemplate(template);
			wordPackage = WordprocessingMLPackage.load(is);
		} catch (Docx4JException e) {
			System.out.println(e.getMessage());
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

	// Una vez que tenemos el objeto WordProcessingMLPackage realizaríamos la
	// siguiente instrucción:
	private MainDocumentPart loadTemplate(WordprocessingMLPackage createWordProcessing) throws Docx4JException {

		MainDocumentPart mainDocumentPart = createWordProcessing.getMainDocumentPart();

		try {
			VariablePrepare.prepare(createWordProcessing);
		} catch (Exception e) {
//				LOG.error(e.getMessage(),e);
			System.out.println(e.getMessage());
		}
		return mainDocumentPart;

	}

//		Con el objeto devuelto (MainDocumentPart), ya empezaríamos a realizar las inserciones en el 
//		futuro documento docx con el siguiente método:

//		mainDocumentPart.addStyledParagraphOfText("LEOSCONCLUSIONPRES", conclusionPresidente);

//		Aquí tener en cuenta que el primer parámetro es un estilo que tendremos creado en la plantilla de Word a usar. 
//		El según parámetro es el String a insertar en el documento docx con el estilo comentado.




}
