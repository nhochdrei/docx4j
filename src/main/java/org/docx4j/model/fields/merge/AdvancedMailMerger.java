package org.docx4j.model.fields.merge;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.ComplexFieldLocator;
import org.docx4j.model.fields.FieldRef;
import org.docx4j.model.fields.FieldsPreprocessor;
import org.docx4j.model.fields.merge.parsers.FieldParser;
import org.docx4j.model.fields.merge.parsers.IfExpressionParser;
import org.docx4j.model.fields.merge.parsers.MergefieldParser;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.Docx4JRuntimeException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.vml.CTTextbox;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;
import org.jvnet.jaxb2_commons.ppp.Child;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AdvancedMailMerger {
	private static Logger log = LoggerFactory.getLogger(AdvancedMailMerger.class);

	public static void performMerge(WordprocessingMLPackage input,
	                                Map<DataFieldName, String> data,
	                                boolean processHeadersAndFooters) throws Docx4JRuntimeException {
		try {
			FieldsPreprocessor.complexifyFields(input.getMainDocumentPart());
		} catch (Docx4JException e) {
			throw Docx4JRuntimeException.wrap(e);
		}
		List<Object> mdpResults = performOnInstance(input, input.getMainDocumentPart().getContent(), data);
		input.getMainDocumentPart().getContent().clear();
		input.getMainDocumentPart().getContent().addAll(mdpResults);

		if (processHeadersAndFooters) {
			RelationshipsPart rp = input.getMainDocumentPart().getRelationshipsPart();
			for (Relationship r : rp.getJaxbElement().getRelationship()) {

				if (r.getType().equals(Namespaces.HEADER) || r.getType().equals(Namespaces.FOOTER)) {
					JaxbXmlPart part = (JaxbXmlPart) rp.getPart(r);

					try {
						FieldsPreprocessor.complexifyFields(part);
					} catch (Docx4JException e) {
						throw Docx4JRuntimeException.wrap(e);
					}

					List<Object> results = performOnInstance(input, ((ContentAccessor) part).getContent(), data);
					((ContentAccessor) part).getContent().clear();
					((ContentAccessor) part).getContent().addAll(results);
				}
			}
		}
	}

	private static List<Object> performOnInstance(WordprocessingMLPackage input,
	                                              List<Object> contentList,
	                                              Map<DataFieldName, String> datamap) throws Docx4JRuntimeException {
		Body shell = Context.getWmlObjectFactory().createBody();
		shell.getContent().addAll(contentList);
		Body shellClone = XmlUtils.deepCopy(shell);

		ComplexFieldLocator fl = new ComplexFieldLocator();
		new TraversalUtil(shellClone, fl);
		log.debug("Found " + fl.getStarts().size() + " fields ");

		List<FieldRef> fieldRefs = new ArrayList<>();
		canonicaliseStarts(fl, fieldRefs);

		// Populate
		for (FieldRef fr : fieldRefs) {
			String val = getFieldValue(input, datamap, fr);

			if (val != null) {
				fr.setResult(val);
			}
		}

		return shellClone.getContent();

	}

	private static String getFieldValue(WordprocessingMLPackage input, Map<DataFieldName, String> datamap, FieldRef fr) {
		final String fn = fr.getFldName();
		final String instr = extractInstr(fr.getInstructions(), input, datamap, fr);
		final String language = extractLang(fr.getResultsSlot());

		FieldParser parser = null;

		if (fn.equals("MERGEFIELD")) {
			parser = new MergefieldParser(input, language, datamap);
		} else if (fn.equals("IF")) {
			parser = new IfExpressionParser();
		}

		return parser == null ? null : parser.value(instr);
	}

	protected static String extractInstr(List<Object> instructions, WordprocessingMLPackage input, Map<DataFieldName, String> datamap, FieldRef fr) {
		String instr = instructions.stream()
		                           .map(XmlUtils::unwrap)
		                           .map(o -> {
			                           if (o instanceof Text) {
				                           return ((Text) o).getValue();
			                           } else if (o instanceof FieldRef) {
				                           String f = getFieldValue(input, datamap, (FieldRef) o);
				                           if (f == null)
					                           return "__MAILMERGER_QUOTE__\"\"__MAILMERGER_QUOTE__";
				                           else
					                           return "__MAILMERGER_QUOTE__\"" +
							                           f.replace("\\", "\\\\")
							                            .replace("\"", "\\\"") +
							                           "\"__MAILMERGER_QUOTE__";
			                           } else if (o instanceof Child) {
				                           return "";
			                           } else {
				                           throw new RuntimeException("Invalid type: " + o.getClass().getCanonicalName());
			                           }
		                           })
		                           .collect(Collectors.joining());
		instr = instr.replace("\"__MAILMERGER_QUOTE__\"", "\"") // Eliminate double quotations
		             .replace("__MAILMERGER_QUOTE__", ""); // Strip remaining quote-markers
		return instr;
	}

	private static String extractLang(R resultsSlot) {
		RPr rPr = resultsSlot.getRPr();
		if (rPr != null) {
			CTLanguage lang = rPr.getLang();
			if (lang != null) {
				return lang.getVal();
			}
		}
		return null;
	}

	protected static void canonicaliseStarts(ComplexFieldLocator fl,
	                                         List<FieldRef> fieldRefs) throws Docx4JRuntimeException {
		for (P p : fl.getStarts()) {
			int index;
			if (p.getParent() instanceof ContentAccessor) {
				// 2.8.1
				index = ((ContentAccessor) p.getParent()).getContent().indexOf(p);
				P newP = FieldsPreprocessor.canonicalise(p, fieldRefs);
				((ContentAccessor) p.getParent()).getContent().set(index, newP);
			} else if (p.getParent() instanceof java.util.List) {
				// 3.0
				index = ((java.util.List) p.getParent()).indexOf(p);
				P newP = FieldsPreprocessor.canonicalise(p, fieldRefs);
				log.debug("NewP length: " + newP.getContent().size());
				((java.util.List) p.getParent()).set(index, newP);
			} else if (p.getParent() instanceof CTTextbox) {
				// 3.0.1
				index = ((CTTextbox) p.getParent()).getTxbxContent().getContent().indexOf(p);
				P newP = FieldsPreprocessor.canonicalise(p, fieldRefs);
				((CTTextbox) p.getParent()).getTxbxContent().getContent().set(index, newP);
			} else {
				throw new Docx4JRuntimeException("Unexpected parent: " + p.getParent().getClass().getName());
			}
		}
	}
}
