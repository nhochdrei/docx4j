package org.docx4j.model.fields.merge.parsers;

@FunctionalInterface
public interface FieldParser {
	String value(String instruction);
}
