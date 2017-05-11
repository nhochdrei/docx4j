package org.docx4j.model.fields.merge.parsers;

import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.transform.TransformerException;

import org.apache.commons.lang3.StringUtils;
import org.docx4j.model.fields.FldSimpleModel;
import org.docx4j.model.fields.FormattingSwitchHelper;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.Docx4JRuntimeException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class MergefieldParser extends AbstractFieldParser {
	private static final Logger log = LoggerFactory.getLogger(MergefieldParser.class);

	private static Pattern MERGEFIELD_PREFIX_REGEX = Pattern.compile("\\\\b\\s+(?:\"((?:[^\"\\\\]|\\\\.)*)\"|([^ ]+))");
	private static Pattern MERGEFIELD_SUFFIX_REGEX = Pattern.compile("\\\\f\\s+(?:\"((?:[^\"\\\\]|\\\\.)*)\"|([^ ]+))");

	public MergefieldParser(WordprocessingMLPackage input, String language, Map<DataFieldName, String> data) {
		super(input, language, data);
	}

	private static String getDatafieldNameFromInstr(String instr) {
		if (instr == null)
			return null;

		String tmp = instr.substring(instr.indexOf("MERGEFIELD") + 10);
		tmp = tmp.trim();
		String datafieldName = null;
		// A data field name will be quoted if it contains spaces
		if (tmp.startsWith("\"")) {
			if (tmp.indexOf("\"", 1) > -1) {
				datafieldName = tmp.substring(1, tmp.indexOf("\"", 1));
			} else {
				log.warn("Quote mismatch in " + instr);
				// hope for the best
				datafieldName = tmp.indexOf(" ") > -1 ? tmp.substring(1, tmp.indexOf(" ")) : tmp.substring(1);
			}
		} else {
			datafieldName = tmp.indexOf(" ") > -1 ? tmp.substring(0, tmp.indexOf(" ")) : tmp;
		}
		log.debug("Key: '" + datafieldName + "'");

		return datafieldName;

	}

	private static String extractPrefix(String val, String instr) {
		Matcher pm = MERGEFIELD_PREFIX_REGEX.matcher(instr);
		if (pm.find()) {
			if (!StringUtils.isEmpty(pm.group(1)))
				val = pm.group(1).concat(val);
			else if (!StringUtils.isBlank(pm.group(2)))
				val = pm.group(2).concat(val);
		}

		return val;
	}

	private static String extractSuffix(String val, String instr) {
		Matcher sm = MERGEFIELD_SUFFIX_REGEX.matcher(instr);
		if (sm.find()) {
			if (!StringUtils.isEmpty(sm.group(1)))
				val = val.concat(sm.group(1));
			else if (!StringUtils.isBlank(sm.group(2)))
				val = val.concat(sm.group(2));
		}

		return val;
	}

	@Override
	public String value(String instr) {
		String datafieldName = getDatafieldNameFromInstr(instr);
		String val = getData().getOrDefault(new DataFieldName(datafieldName), "");

		if (!StringUtils.isEmpty(val)) {
			FldSimpleModel fsm = new FldSimpleModel();
			try {
				fsm.build(instr);
				try {
					val = FormattingSwitchHelper.applyFormattingSwitch(getInput(), fsm, val, getLanguage());
				} catch (Docx4JException e) {
					throw Docx4JRuntimeException.wrap(e);
				}
			} catch (TransformerException e) {
				log.warn("Can't format the field", e);
			}

			val = extractPrefix(val, instr);
			val = extractSuffix(val, instr);
		}

		return val;
	}
}
