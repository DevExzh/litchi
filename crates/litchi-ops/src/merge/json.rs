//! A bounded JSON reader that keeps number spellings (RFC 8259).
//!
//! Merge data renders a number as it is written (`5200.00` stays
//! `5200.00`), as OfficeCLI does; `serde_json` would normalize it. Objects
//! keep their members in source order.

use super::Error;

/// The deepest nesting a merge data source may use.
pub(super) const MAX_DEPTH: usize = 128;

/// One JSON value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Value {
    Null,
    Bool(bool),
    /// The number's exact spelling.
    Number(String),
    String(String),
    Array(Vec<Value>),
    Object(Vec<(String, Value)>),
}

/// Parse one complete JSON text.
pub(super) fn parse(source: &[u8]) -> Result<Value, Error> {
    let text = std::str::from_utf8(source).map_err(|error| Error::Json {
        line: line_of(source, error.valid_up_to()),
        column: column_of(source, error.valid_up_to()),
    })?;
    let text = text.strip_prefix('\u{FEFF}').unwrap_or(text);
    let mut parser = Parser {
        bytes: text.as_bytes(),
        position: 0,
    };
    parser.skip_whitespace();
    let value = parser.value(0)?;
    parser.skip_whitespace();
    if parser.position != parser.bytes.len() {
        return Err(parser.error());
    }
    Ok(value)
}

fn line_of(bytes: &[u8], position: usize) -> usize {
    bytes[..position.min(bytes.len())]
        .iter()
        .filter(|byte| **byte == b'\n')
        .count()
        + 1
}

fn column_of(bytes: &[u8], position: usize) -> usize {
    let position = position.min(bytes.len());
    let start = bytes[..position]
        .iter()
        .rposition(|byte| *byte == b'\n')
        .map_or(0, |newline| newline + 1);
    String::from_utf8_lossy(&bytes[start..position])
        .chars()
        .count()
        + 1
}

struct Parser<'a> {
    bytes: &'a [u8],
    position: usize,
}

impl Parser<'_> {
    fn error(&self) -> Error {
        Error::Json {
            line: line_of(self.bytes, self.position),
            column: column_of(self.bytes, self.position),
        }
    }

    fn peek(&self) -> Option<u8> {
        self.bytes.get(self.position).copied()
    }

    fn skip_whitespace(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\t' | b'\n' | b'\r')) {
            self.position += 1;
        }
    }

    fn expect(&mut self, literal: &[u8]) -> Result<(), Error> {
        if self.bytes[self.position..].starts_with(literal) {
            self.position += literal.len();
            Ok(())
        } else {
            Err(self.error())
        }
    }

    fn value(&mut self, depth: usize) -> Result<Value, Error> {
        if depth > MAX_DEPTH {
            return Err(Error::TooDeep);
        }
        match self.peek() {
            Some(b'{') => self.object(depth),
            Some(b'[') => self.array(depth),
            Some(b'"') => self.string().map(Value::String),
            Some(b't') => self.expect(b"true").map(|()| Value::Bool(true)),
            Some(b'f') => self.expect(b"false").map(|()| Value::Bool(false)),
            Some(b'n') => self.expect(b"null").map(|()| Value::Null),
            Some(b'-' | b'0'..=b'9') => self.number(),
            _ => Err(self.error()),
        }
    }

    fn object(&mut self, depth: usize) -> Result<Value, Error> {
        self.position += 1;
        let mut members = Vec::new();
        self.skip_whitespace();
        if self.peek() == Some(b'}') {
            self.position += 1;
            return Ok(Value::Object(members));
        }
        loop {
            self.skip_whitespace();
            if self.peek() != Some(b'"') {
                return Err(self.error());
            }
            let key = self.string()?;
            self.skip_whitespace();
            self.expect(b":")?;
            self.skip_whitespace();
            let value = self.value(depth + 1)?;
            members.push((key, value));
            self.skip_whitespace();
            match self.peek() {
                Some(b',') => self.position += 1,
                Some(b'}') => {
                    self.position += 1;
                    return Ok(Value::Object(members));
                },
                _ => return Err(self.error()),
            }
        }
    }

    fn array(&mut self, depth: usize) -> Result<Value, Error> {
        self.position += 1;
        let mut items = Vec::new();
        self.skip_whitespace();
        if self.peek() == Some(b']') {
            self.position += 1;
            return Ok(Value::Array(items));
        }
        loop {
            self.skip_whitespace();
            items.push(self.value(depth + 1)?);
            self.skip_whitespace();
            match self.peek() {
                Some(b',') => self.position += 1,
                Some(b']') => {
                    self.position += 1;
                    return Ok(Value::Array(items));
                },
                _ => return Err(self.error()),
            }
        }
    }

    fn number(&mut self) -> Result<Value, Error> {
        let start = self.position;
        if self.peek() == Some(b'-') {
            self.position += 1;
        }
        match self.peek() {
            Some(b'0') => self.position += 1,
            Some(b'1'..=b'9') => self.digits(),
            _ => return Err(self.error()),
        }
        if self.peek() == Some(b'.') {
            self.position += 1;
            if !matches!(self.peek(), Some(b'0'..=b'9')) {
                return Err(self.error());
            }
            self.digits();
        }
        if matches!(self.peek(), Some(b'e' | b'E')) {
            self.position += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.position += 1;
            }
            if !matches!(self.peek(), Some(b'0'..=b'9')) {
                return Err(self.error());
            }
            self.digits();
        }
        let lexical = std::str::from_utf8(&self.bytes[start..self.position])
            .map_err(|_| self.error())?
            .to_owned();
        Ok(Value::Number(lexical))
    }

    fn digits(&mut self) {
        while matches!(self.peek(), Some(b'0'..=b'9')) {
            self.position += 1;
        }
    }

    fn string(&mut self) -> Result<String, Error> {
        self.position += 1;
        let mut value = String::new();
        loop {
            let start = self.position;
            while let Some(byte) = self.peek() {
                if byte == b'"' || byte == b'\\' || byte < 0x20 {
                    break;
                }
                self.position += 1;
            }
            value.push_str(
                std::str::from_utf8(&self.bytes[start..self.position]).map_err(|_| self.error())?,
            );
            match self.peek() {
                Some(b'"') => {
                    self.position += 1;
                    return Ok(value);
                },
                Some(b'\\') => {
                    self.position += 1;
                    let escaped = self.peek().ok_or_else(|| self.error())?;
                    self.position += 1;
                    match escaped {
                        b'"' => value.push('"'),
                        b'\\' => value.push('\\'),
                        b'/' => value.push('/'),
                        b'b' => value.push('\u{8}'),
                        b'f' => value.push('\u{C}'),
                        b'n' => value.push('\n'),
                        b'r' => value.push('\r'),
                        b't' => value.push('\t'),
                        b'u' => value.push(self.unicode_escape()?),
                        _ => {
                            self.position -= 1;
                            return Err(self.error());
                        },
                    }
                },
                _ => return Err(self.error()),
            }
        }
    }

    fn hex4(&mut self) -> Result<u32, Error> {
        let digits = self
            .bytes
            .get(self.position..self.position + 4)
            .and_then(|digits| std::str::from_utf8(digits).ok())
            .and_then(|digits| u32::from_str_radix(digits, 16).ok())
            .ok_or_else(|| self.error())?;
        self.position += 4;
        Ok(digits)
    }

    fn unicode_escape(&mut self) -> Result<char, Error> {
        let first = self.hex4()?;
        let code = if (0xD800..0xDC00).contains(&first) {
            if !self.bytes[self.position..].starts_with(b"\\u") {
                return Err(self.error());
            }
            self.position += 2;
            let second = self.hex4()?;
            if !(0xDC00..0xE000).contains(&second) {
                return Err(self.error());
            }
            0x10000 + ((first - 0xD800) << 10) + (second - 0xDC00)
        } else {
            first
        };
        char::from_u32(code).ok_or_else(|| self.error())
    }
}

/// The text of a value as OfficeCLI renders it: a string as itself, a
/// number as written, `true`/`false`, `null` as nothing, and an object or
/// array as indented JSON.
pub(super) fn render(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::Bool(flag) => flag.to_string(),
        Value::Number(lexical) => lexical.clone(),
        Value::String(text) => text.clone(),
        Value::Array(_) | Value::Object(_) => {
            let mut output = String::new();
            write_indented(&mut output, value, 0);
            output
        },
    }
}

fn write_indented(output: &mut String, value: &Value, indent: usize) {
    let pad = |output: &mut String, level: usize| {
        for _ in 0..level {
            output.push_str("  ");
        }
    };
    match value {
        Value::Null => output.push_str("null"),
        Value::Bool(flag) => output.push_str(if *flag { "true" } else { "false" }),
        Value::Number(lexical) => output.push_str(lexical),
        Value::String(text) => write_string(output, text),
        Value::Array(items) => {
            if items.is_empty() {
                output.push_str("[]");
                return;
            }
            output.push('[');
            for (index, item) in items.iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                output.push('\n');
                pad(output, indent + 1);
                write_indented(output, item, indent + 1);
            }
            output.push('\n');
            pad(output, indent);
            output.push(']');
        },
        Value::Object(members) => {
            if members.is_empty() {
                output.push_str("{}");
                return;
            }
            output.push('{');
            for (index, (key, item)) in members.iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                output.push('\n');
                pad(output, indent + 1);
                write_string(output, key);
                output.push_str(": ");
                write_indented(output, item, indent + 1);
            }
            output.push('\n');
            pad(output, indent);
            output.push('}');
        },
    }
}

fn write_string(output: &mut String, text: &str) {
    output.push('"');
    for character in text.chars() {
        match character {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            control if u32::from(control) < 0x20 => {
                output.push_str(&format!("\\u{:04x}", u32::from(control)));
            },
            other => output.push(other),
        }
    }
    output.push('"');
}
