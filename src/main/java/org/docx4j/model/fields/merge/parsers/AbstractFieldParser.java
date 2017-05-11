package org.docx4j.model.fields.merge.parsers;

import java.util.Map;

import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public abstract class AbstractFieldParser implements FieldParser {
	private WordprocessingMLPackage input;
	private String language;
	private Map<DataFieldName, String> data;

	public AbstractFieldParser(WordprocessingMLPackage input, String language, Map<DataFieldName, String> data) {
		this.input = input;
		this.language = language;
		this.data = data;
	}

	protected WordprocessingMLPackage getInput() {
		return input;
	}

	protected String getLanguage() {
		return language;
	}

	protected Map<DataFieldName, String> getData() {
		return data;
	}
}
