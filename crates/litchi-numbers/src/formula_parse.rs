//! A bounded parser from Numbers formula text to an [`Expression`]
//! (IWORK_API_REQUESTS.md request 17).
//!
//! The accepted language is the local, A1-addressed subset the authoring
//! model represents: finite numbers, double-quoted text (`""` escapes a
//! quote), `TRUE`/`FALSE`, cell references with optional `$` anchors, cell
//! ranges (`A1:B3`), whole-column (`A:C`) and whole-row (`1:4`) ranges, the
//! authorable functions, the binary operators `^ * / + - & = <> < <= > >=`,
//! unary `-`/`+`, postfix `%` and parentheses. A leading `=` is optional.
//! Cross-table references need source-bound table handles and are built
//! with [`Expression::table_cell`] instead.

use super::{AxisReference, BinaryOperator, CellReference, Error, Expression};

/// Nesting depth the recursive-descent parser admits before refusing.
const MAX_PARSE_DEPTH: usize = 128;
/// Bytes of formula text the parser admits.
const MAX_TEXT_BYTES: usize = 64 * 1024;

impl Expression {
    /// Parse formula text such as `=SUM(A1:A3)*2` into an expression whose
    /// references are local to the formula's own table.
    ///
    /// ```
    /// use litchi_numbers::formula::Expression;
    ///
    /// # fn main() -> Result<(), litchi_numbers::formula::Error> {
    /// let sum = Expression::parse("=SUM(A1:A3)*2")?;
    /// let same = Expression::parse("SUM(A1:A3) * 2")?;
    /// assert!(sum == same);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Error::Syntax`] for text outside the accepted language (including
    /// cross-table references), [`Error::UnsupportedFunction`] or
    /// [`Error::InvalidArity`] for a function outside the authorable subset,
    /// and the construction errors of the individual expression builders.
    pub fn parse(text: &str) -> Result<Self, Error> {
        if text.len() > MAX_TEXT_BYTES {
            return Err(Error::Limit {
                kind: super::LimitKind::OwnedBytes,
                observed: text.len(),
                maximum: MAX_TEXT_BYTES,
            });
        }
        let mut parser = Parser {
            source: text.as_bytes(),
            offset: 0,
            depth: 0,
        };
        parser.skip_space();
        if parser.peek() == Some(b'=') {
            parser.offset += 1;
        }
        let expression = parser.comparison()?;
        parser.skip_space();
        if parser.offset != parser.source.len() {
            return Err(parser.syntax());
        }
        Ok(expression)
    }
}

struct Parser<'a> {
    source: &'a [u8],
    offset: usize,
    depth: usize,
}

impl Parser<'_> {
    const fn syntax(&self) -> Error {
        Error::Syntax {
            offset: self.offset,
        }
    }

    fn peek(&self) -> Option<u8> {
        self.source.get(self.offset).copied()
    }

    fn skip_space(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\t' | b'\r' | b'\n')) {
            self.offset += 1;
        }
    }

    fn eat(&mut self, token: &[u8]) -> bool {
        self.skip_space();
        if self.source[self.offset..].starts_with(token) {
            self.offset += token.len();
            true
        } else {
            false
        }
    }

    fn enter(&mut self) -> Result<(), Error> {
        self.depth += 1;
        if self.depth > MAX_PARSE_DEPTH {
            return Err(Error::Limit {
                kind: super::LimitKind::Depth,
                observed: self.depth,
                maximum: MAX_PARSE_DEPTH,
            });
        }
        Ok(())
    }

    fn comparison(&mut self) -> Result<Expression, Error> {
        self.enter()?;
        let mut left = self.concatenation()?;
        loop {
            let operator = if self.eat(b"<>") {
                BinaryOperator::NotEqual
            } else if self.eat(b"<=") {
                BinaryOperator::LessThanOrEqual
            } else if self.eat(b">=") {
                BinaryOperator::GreaterThanOrEqual
            } else if self.eat(b"<") {
                BinaryOperator::LessThan
            } else if self.eat(b">") {
                BinaryOperator::GreaterThan
            } else if self.eat(b"=") {
                BinaryOperator::Equal
            } else {
                break;
            };
            let right = self.concatenation()?;
            left = Expression::binary(operator, left, right)?;
        }
        self.depth -= 1;
        Ok(left)
    }

    fn concatenation(&mut self) -> Result<Expression, Error> {
        let mut left = self.additive()?;
        while self.eat(b"&") {
            let right = self.additive()?;
            left = Expression::binary(BinaryOperator::Concatenate, left, right)?;
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expression, Error> {
        let mut left = self.multiplicative()?;
        loop {
            let operator = if self.eat(b"+") {
                BinaryOperator::Add
            } else if self.eat(b"-") {
                BinaryOperator::Subtract
            } else {
                break;
            };
            let right = self.multiplicative()?;
            left = Expression::binary(operator, left, right)?;
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> Result<Expression, Error> {
        let mut left = self.power()?;
        loop {
            let operator = if self.eat(b"*") {
                BinaryOperator::Multiply
            } else if self.eat(b"/") {
                BinaryOperator::Divide
            } else {
                break;
            };
            let right = self.power()?;
            left = Expression::binary(operator, left, right)?;
        }
        Ok(left)
    }

    fn power(&mut self) -> Result<Expression, Error> {
        let mut left = self.unary()?;
        while self.eat(b"^") {
            let right = self.unary()?;
            left = Expression::binary(BinaryOperator::Power, left, right)?;
        }
        Ok(left)
    }

    fn unary(&mut self) -> Result<Expression, Error> {
        self.enter()?;
        let expression = if self.eat(b"-") {
            Expression::negate(self.unary()?)?
        } else if self.eat(b"+") {
            self.unary()?
        } else {
            self.postfix()?
        };
        self.depth -= 1;
        Ok(expression)
    }

    fn postfix(&mut self) -> Result<Expression, Error> {
        let mut value = self.primary()?;
        while self.eat(b"%") {
            value = Expression::percent(value)?;
        }
        Ok(value)
    }

    fn primary(&mut self) -> Result<Expression, Error> {
        self.skip_space();
        match self.peek() {
            Some(b'(') => {
                self.offset += 1;
                let inner = self.comparison()?;
                if !self.eat(b")") {
                    return Err(self.syntax());
                }
                Ok(inner)
            },
            Some(b'"') => self.text(),
            Some(byte) if byte.is_ascii_digit() || byte == b'.' => {
                // `1:4` is a row range; any other digit run is a number.
                if let Some(rows) = self.row_range()? {
                    return Ok(rows);
                }
                self.number()
            },
            Some(byte) if byte.is_ascii_alphabetic() || byte == b'$' => self.word(),
            _ => Err(self.syntax()),
        }
    }

    fn number(&mut self) -> Result<Expression, Error> {
        let start = self.offset;
        while matches!(self.peek(), Some(byte) if byte.is_ascii_digit()) {
            self.offset += 1;
        }
        if self.peek() == Some(b'.') {
            self.offset += 1;
            while matches!(self.peek(), Some(byte) if byte.is_ascii_digit()) {
                self.offset += 1;
            }
        }
        if matches!(self.peek(), Some(b'e' | b'E')) {
            let mark = self.offset;
            self.offset += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.offset += 1;
            }
            if matches!(self.peek(), Some(byte) if byte.is_ascii_digit()) {
                while matches!(self.peek(), Some(byte) if byte.is_ascii_digit()) {
                    self.offset += 1;
                }
            } else {
                self.offset = mark;
            }
        }
        let literal = std::str::from_utf8(&self.source[start..self.offset])
            .map_err(|_| Error::Syntax { offset: start })?;
        let value: f64 = literal
            .parse()
            .map_err(|_| Error::Syntax { offset: start })?;
        Expression::number(value)
    }

    fn text(&mut self) -> Result<Expression, Error> {
        let start = self.offset;
        self.offset += 1;
        let mut value = Vec::new();
        loop {
            match self.peek() {
                None => return Err(Error::Syntax { offset: start }),
                Some(b'"') => {
                    self.offset += 1;
                    if self.peek() == Some(b'"') {
                        value.push(b'"');
                        self.offset += 1;
                    } else {
                        break;
                    }
                },
                Some(byte) => {
                    value.push(byte);
                    self.offset += 1;
                },
            }
        }
        let value = String::from_utf8(value).map_err(|_| Error::Syntax { offset: start })?;
        Expression::text(value)
    }

    /// A function call, a Boolean literal, or a cell/column reference.
    fn word(&mut self) -> Result<Expression, Error> {
        let start = self.offset;
        let mut end = start;
        while matches!(self.source.get(end), Some(byte) if byte.is_ascii_alphanumeric() || *byte == b'.' || *byte == b'_')
        {
            end += 1;
        }
        let word = &self.source[start..end];
        let after = {
            let mut cursor = end;
            while matches!(self.source.get(cursor), Some(b' ' | b'\t')) {
                cursor += 1;
            }
            self.source.get(cursor).copied()
        };
        if after == Some(b'(') && word.first().is_some_and(u8::is_ascii_alphabetic) {
            self.offset = end;
            self.eat(b"(");
            let name = std::str::from_utf8(word).map_err(|_| Error::Syntax { offset: start })?;
            let mut arguments = Vec::new();
            if !self.eat(b")") {
                loop {
                    arguments.push(self.comparison()?);
                    if self.eat(b")") {
                        break;
                    }
                    if !self.eat(b",") {
                        return Err(self.syntax());
                    }
                }
            }
            return Expression::function(name, arguments);
        }
        if word.eq_ignore_ascii_case(b"TRUE") {
            self.offset = end;
            return Ok(Expression::boolean(true));
        }
        if word.eq_ignore_ascii_case(b"FALSE") {
            self.offset = end;
            return Ok(Expression::boolean(false));
        }
        self.reference()
    }

    /// `[$]COL[$]ROW` optionally followed by `:` and another reference, or
    /// a whole-column range `[$]COL:[$]COL`.
    fn reference(&mut self) -> Result<Expression, Error> {
        let start = self.offset;
        let first = self.endpoint().ok_or(Error::Syntax { offset: start })?;
        if self.peek() == Some(b':') {
            self.offset += 1;
            let second_start = self.offset;
            let second = self.endpoint().ok_or(Error::Syntax {
                offset: second_start,
            })?;
            return match (first, second) {
                (Endpoint::Cell(start), Endpoint::Cell(end)) => Expression::range(start, end),
                (Endpoint::Column(start), Endpoint::Column(end)) => Expression::columns(start, end),
                _ => Err(Error::Syntax {
                    offset: second_start,
                }),
            };
        }
        match first {
            Endpoint::Cell(reference) => Ok(Expression::cell(reference)),
            Endpoint::Column(_) => Err(Error::Syntax { offset: start }),
        }
    }

    /// A whole-row range `[$]ROW:[$]ROW`, when the input has that shape.
    fn row_range(&mut self) -> Result<Option<Expression>, Error> {
        let start = self.offset;
        let Some(first) = self.row_endpoint() else {
            self.offset = start;
            return Ok(None);
        };
        if self.peek() != Some(b':') {
            self.offset = start;
            return Ok(None);
        }
        self.offset += 1;
        let Some(second) = self.row_endpoint() else {
            return Err(self.syntax());
        };
        Expression::rows(first, second).map(Some)
    }

    fn row_endpoint(&mut self) -> Option<AxisReference> {
        let absolute = self.peek() == Some(b'$');
        if absolute {
            self.offset += 1;
        }
        let row = self.digits()?;
        Some(if absolute {
            AxisReference::absolute(row.checked_sub(1)?)
        } else {
            AxisReference::relative(row.checked_sub(1)?)
        })
    }

    fn endpoint(&mut self) -> Option<Endpoint> {
        let absolute_column = self.peek() == Some(b'$');
        if absolute_column {
            self.offset += 1;
        }
        let mut column = 0_usize;
        let mut letters = 0_usize;
        while let Some(byte) = self.peek().filter(u8::is_ascii_alphabetic) {
            column = column
                .checked_mul(26)?
                .checked_add(usize::from(byte.to_ascii_uppercase() - b'A') + 1)?;
            letters += 1;
            self.offset += 1;
        }
        if letters == 0 || letters > 4 {
            return None;
        }
        let column = column - 1;
        let absolute_row = self.peek() == Some(b'$');
        if absolute_row {
            self.offset += 1;
        }
        match self.digits() {
            Some(row) => Some(Endpoint::Cell(CellReference::mixed(
                row.checked_sub(1)?,
                column,
                absolute_row,
                absolute_column,
            ))),
            None if !absolute_row => Some(Endpoint::Column(if absolute_column {
                AxisReference::absolute(column)
            } else {
                AxisReference::relative(column)
            })),
            None => None,
        }
    }

    fn digits(&mut self) -> Option<usize> {
        let start = self.offset;
        let mut value = 0_usize;
        while let Some(byte) = self.peek().filter(u8::is_ascii_digit) {
            value = value
                .checked_mul(10)?
                .checked_add(usize::from(byte - b'0'))?;
            self.offset += 1;
        }
        (self.offset > start).then_some(value)
    }
}

enum Endpoint {
    Cell(CellReference),
    Column(AxisReference),
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(text: &str) -> Expression {
        Expression::parse(text).unwrap_or_else(|error| panic!("{text}: {error}"))
    }

    #[test]
    fn literals_references_and_ranges() {
        assert!(parse("=1.5") == Expression::number(1.5).unwrap());
        assert!(parse("\"a\"\"b\"") == Expression::text("a\"b").unwrap());
        assert!(parse("TRUE") == Expression::boolean(true));
        assert!(parse("B3") == Expression::cell(CellReference::relative(2, 1)));
        assert!(parse("$B$3") == Expression::cell(CellReference::absolute(2, 1)));
        assert!(
            parse("A1:B2")
                == Expression::range(CellReference::relative(0, 0), CellReference::relative(1, 1))
                    .unwrap()
        );
        assert!(
            parse("A:C")
                == Expression::columns(AxisReference::relative(0), AxisReference::relative(2))
                    .unwrap()
        );
        assert!(
            parse("2:4")
                == Expression::rows(AxisReference::relative(1), AxisReference::relative(3))
                    .unwrap()
        );
    }

    #[test]
    fn precedence_functions_and_unary_operators() {
        let expected = Expression::binary(
            BinaryOperator::Add,
            Expression::number(1.0).unwrap(),
            Expression::binary(
                BinaryOperator::Multiply,
                Expression::number(2.0).unwrap(),
                Expression::number(3.0).unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
        assert!(parse("1+2*3") == expected);
        assert!(
            parse("sum(A1, 2)")
                == Expression::function(
                    "SUM",
                    [
                        Expression::cell(CellReference::relative(0, 0)),
                        Expression::number(2.0).unwrap(),
                    ]
                )
                .unwrap()
        );
        assert!(
            parse("-A1%")
                == Expression::negate(
                    Expression::percent(Expression::cell(CellReference::relative(0, 0))).unwrap()
                )
                .unwrap()
        );
        assert!(
            parse("A1<>2")
                == Expression::binary(
                    BinaryOperator::NotEqual,
                    Expression::cell(CellReference::relative(0, 0)),
                    Expression::number(2.0).unwrap()
                )
                .unwrap()
        );
    }

    #[test]
    fn malformed_text_and_unknown_functions_are_typed_errors() {
        assert!(matches!(Expression::parse("1+"), Err(Error::Syntax { .. })));
        assert!(matches!(Expression::parse("(1"), Err(Error::Syntax { .. })));
        assert!(matches!(
            Expression::parse("Table 2::A1"),
            Err(Error::Syntax { .. })
        ));
        assert!(matches!(
            Expression::parse("VLOOKUP(A1)"),
            Err(Error::UnsupportedFunction | Error::Limit { .. })
        ));
        assert!(matches!(
            Expression::parse("IF(A1)"),
            Err(Error::InvalidArity)
        ));
        let deep = "(".repeat(200) + "1" + &")".repeat(200);
        assert!(matches!(Expression::parse(&deep), Err(Error::Limit { .. })));
    }
}
