package org.docx4j.model.fields.merge.parsers;

import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class IfExpressionParser implements FieldParser {
	private static final Logger log = LoggerFactory.getLogger(IfExpressionParser.class);
	private Lexer lexer;

	@Override
	public String value(String instruction) {
		try {
			return parse(instruction);
		} catch (ParserException ex) {
			log.warn("Error parsing if expression", ex);
			return null;
		}
	}

	private String parse(String instruction) {
		try {
			lexer = new Lexer(instruction.trim());
			return parseIf();
		} catch (ParserException ex) {
			throw ex;
		} catch (Exception ex) {
			throw new ParserException("Error in top-level parsing", ex);
		}
	}

	private String parseIf() {
		TokenType curr = lexer.next();
		if (curr != TokenType.IF)
			throw new ParserException("Expected beginning IF, found " + curr.name());

		String a = parseValue();

		TokenType op = lexer.next();
		if (op != TokenType.OP_EQ && op != TokenType.OP_NEQ)
			throw new ParserException("Expected OP_EQ[=] or OP_NEQ[<>], found " + op.name());

		String b = parseValue();

		String trueCase = parseValue();
		String falseCase = parseValue();

		if (op == TokenType.OP_EQ)
			return Objects.equals(a, b) ? trueCase : falseCase;
		else // op == OP_NEQ
			return !Objects.equals(a, b) ? trueCase : falseCase;
	}

	private String parseValue() {
		TokenType curr = lexer.next();
		if (curr != TokenType.PLAIN && curr != TokenType.STRING)
			throw new ParserException("Expected PLAIN or STRING, found " + curr.name());
		return lexer.value();
	}

	public enum TokenType {
		INVALID,
		STRING, // "foo "
		PLAIN, // foo
		IF, // IF
		OP_EQ, // =
		OP_NEQ // <>
	}

	private class Lexer {
		private String content;
		private int pos = 0;
		private int lastpos = 0;
		private String lastval = null;

		public Lexer(String content) {
			this.content = content;
		}

		public TokenType next() {
			try {
				skipWhiteSpace();

				char c = nextChar();

				switch (c) {
					case 'I':
						if (nextChar() == 'F')
							return endToken(TokenType.IF);
						back(); // Put back for further inspection
						break;
					case '=':
						return endToken(TokenType.OP_EQ);
					case '<':
						if (nextChar() == '>')
							return endToken(TokenType.OP_NEQ);
						back(); // Put back for further inspection
						break;
					case '"':
						boolean escape = false;
						char curr;
						StringBuilder val = new StringBuilder(32);
						while ((curr = nextChar()) != '"' || escape) {
							if (curr == '\\') {
								if (escape)
									val.append(curr);
								escape = !escape;
							} else {
								escape = false;
								val.append(curr);
							}
						}
						return endToken(TokenType.STRING, val.toString());
				}

				// Assume plain
				StringBuilder builder = new StringBuilder(32);
				while (!Character.isWhitespace(c) || pos == content.length()) {
					builder.append(c);
					c = nextChar();
				}
				return endToken(TokenType.PLAIN, builder.toString());
			} catch (Exception ex) {
				throw new ParserException("Error lexing near '" + content.substring(lastpos) + "'", ex);
			}
		}

		public String value() {
			return lastval;
		}

		private char nextChar() {
			return content.charAt(pos++);
		}

		private void back() {
			--pos;
		}

		private void skipWhiteSpace() {
			while (Character.isWhitespace(content.charAt(pos))) {
				++pos;
			}
		}

		private TokenType endToken(TokenType token) {
			return endToken(token, null);
		}

		private TokenType endToken(TokenType token, String val) {
			lastval = val;
			lastpos = pos;
			return token;
		}
	}

	private class ParserException extends RuntimeException {
		public ParserException(String message) {
			super(message);
		}

		public ParserException(String message, Throwable cause) {
			super(message, cause);
		}
	}
}
