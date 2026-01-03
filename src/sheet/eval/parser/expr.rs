//! Expression parsing (tokens + recursive descent parser).

use super::ast::{BinaryOp, Expr};
use super::literal::parse_literal;
use super::reference::{parse_range_reference, parse_single_cell_reference};

#[derive(Debug, Clone)]
enum Token {
    Atom(String),
    Plus,
    Minus,
    Star,
    Slash,
    Comma,
    LParen,
    RParen,
    Eq,
    Lt,
    Gt,
    Le,
    Ge,
    Ne,
}

/// Parse a full expression string into an AST.
///
/// Supported grammar (minimal subset):
///
/// expr  := term (("+" | "-") term)*
/// term  := factor (("*" | "/") factor)*
/// factor:= "-" factor
///        | primary
/// primary := ATOM        // literal, reference, or function name
///          | "(" expr ")"
pub fn parse_expression(current_sheet: &str, input: &str) -> Option<Expr> {
    let tokens = tokenize(input)?;
    if tokens.is_empty() {
        return None;
    }

    let mut parser = ExprParser {
        tokens,
        pos: 0,
        current_sheet,
    };

    let expr = parser.parse_expr()?;

    // All tokens must be consumed for a valid expression.
    if parser.peek().is_some() {
        return None;
    }

    Some(expr)
}

fn tokenize(input: &str) -> Option<Vec<Token>> {
    let mut tokens = Vec::new();
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.peek().copied() {
        if ch.is_whitespace() {
            chars.next();
            continue;
        }

        // String literal starting with '"' – consume until the matching
        // closing quote, allowing for doubled quotes inside (Excel-style
        // escaping). The whole literal, including quotes, becomes a single
        // Atom token so that parse_literal can handle it.
        if ch == '"' {
            let mut buf = String::new();
            buf.push('"');
            chars.next();

            while let Some(c) = chars.next() {
                buf.push(c);
                if c == '"' {
                    // Escaped quote "" inside the string – keep both
                    // characters and continue.
                    if chars.peek() == Some(&'"') {
                        buf.push('"');
                        chars.next();
                        continue;
                    }
                    // Closing quote.
                    break;
                }
            }

            if !buf.ends_with('"') {
                // Unterminated string literal.
                return None;
            }

            tokens.push(Token::Atom(buf));
            continue;
        }

        match ch {
            '+' => {
                tokens.push(Token::Plus);
                chars.next();
            },
            '-' => {
                tokens.push(Token::Minus);
                chars.next();
            },
            '*' => {
                tokens.push(Token::Star);
                chars.next();
            },
            '/' => {
                tokens.push(Token::Slash);
                chars.next();
            },
            '=' => {
                tokens.push(Token::Eq);
                chars.next();
            },
            '<' => {
                chars.next();
                match chars.peek().copied() {
                    Some('=') => {
                        chars.next();
                        tokens.push(Token::Le);
                    },
                    Some('>') => {
                        chars.next();
                        tokens.push(Token::Ne);
                    },
                    _ => tokens.push(Token::Lt),
                }
            },
            '>' => {
                chars.next();
                match chars.peek().copied() {
                    Some('=') => {
                        chars.next();
                        tokens.push(Token::Ge);
                    },
                    _ => tokens.push(Token::Gt),
                }
            },
            ',' => {
                tokens.push(Token::Comma);
                chars.next();
            },
            '(' => {
                tokens.push(Token::LParen);
                chars.next();
            },
            ')' => {
                tokens.push(Token::RParen);
                chars.next();
            },
            _ => {
                // Collect an atom until we hit whitespace or an operator/paren.
                let mut buf = String::new();
                while let Some(c) = chars.peek().copied() {
                    if c.is_whitespace()
                        || matches!(c, '+' | '-' | '*' | '/' | '(' | ')' | ',' | '<' | '>' | '=')
                    {
                        break;
                    }
                    buf.push(c);
                    chars.next();
                }

                if buf.is_empty() {
                    return None;
                }

                tokens.push(Token::Atom(buf));
            },
        }
    }

    Some(tokens)
}

struct ExprParser<'a> {
    tokens: Vec<Token>,
    pos: usize,
    current_sheet: &'a str,
}

impl<'a> ExprParser<'a> {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn next(&mut self) -> Option<&Token> {
        let tok = self.tokens.get(self.pos);
        if tok.is_some() {
            self.pos += 1;
        }
        tok
    }

    fn parse_expr(&mut self) -> Option<Expr> {
        // Lowest-precedence level: comparison operators. These operate on the
        // results of additive/subtractive expressions.
        let mut node = self.parse_add_sub()?;

        loop {
            match self.peek() {
                Some(Token::Eq) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Eq,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Ne) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Ne,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Gt) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Gt,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Ge) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Ge,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Lt) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Lt,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Le) => {
                    self.next();
                    let rhs = self.parse_add_sub()?;
                    node = Expr::Binary {
                        op: BinaryOp::Le,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                _ => break,
            }
        }

        Some(node)
    }

    fn parse_add_sub(&mut self) -> Option<Expr> {
        let mut node = self.parse_term()?;

        loop {
            match self.peek() {
                Some(Token::Plus) => {
                    self.next();
                    let rhs = self.parse_term()?;
                    node = Expr::Binary {
                        op: BinaryOp::Add,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Minus) => {
                    self.next();
                    let rhs = self.parse_term()?;
                    node = Expr::Binary {
                        op: BinaryOp::Sub,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                _ => break,
            }
        }

        Some(node)
    }

    fn parse_term(&mut self) -> Option<Expr> {
        let mut node = self.parse_factor()?;

        loop {
            match self.peek() {
                Some(Token::Star) => {
                    self.next();
                    let rhs = self.parse_factor()?;
                    node = Expr::Binary {
                        op: BinaryOp::Mul,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                Some(Token::Slash) => {
                    self.next();
                    let rhs = self.parse_factor()?;
                    node = Expr::Binary {
                        op: BinaryOp::Div,
                        left: Box::new(node),
                        right: Box::new(rhs),
                    };
                },
                _ => break,
            }
        }

        Some(node)
    }

    fn parse_factor(&mut self) -> Option<Expr> {
        match self.peek() {
            Some(Token::Minus) => {
                // Unary minus
                self.next();
                let inner = self.parse_factor()?;
                Some(Expr::UnaryMinus(Box::new(inner)))
            },
            Some(Token::LParen) => {
                self.next();
                let expr = self.parse_expr()?;
                match self.next() {
                    Some(Token::RParen) => Some(expr),
                    _ => None,
                }
            },
            Some(Token::Atom(_)) => self.parse_atom(),
            _ => None,
        }
    }

    fn parse_atom(&mut self) -> Option<Expr> {
        let atom = match self.next() {
            Some(Token::Atom(s)) => s.clone(),
            _ => return None,
        };

        // Literal value
        if let Some(lit) = parse_literal(&atom) {
            return Some(Expr::Literal(lit));
        }

        // Function call: NAME(expr, expr, ...)
        if matches!(self.peek(), Some(Token::LParen)) {
            let name = atom.to_uppercase();
            // Consume '('
            self.next();

            let mut args = Vec::new();
            // Handle empty argument list: NAME()
            if !matches!(self.peek(), Some(Token::RParen)) {
                loop {
                    let expr = self.parse_expr()?;
                    args.push(expr);
                    match self.peek() {
                        Some(Token::Comma) => {
                            self.next();
                        },
                        Some(Token::RParen) => {
                            self.next();
                            break;
                        },
                        _ => return None,
                    }
                }
            } else {
                // Consume ')'
                self.next();
            }

            return Some(Expr::FunctionCall { name, args });
        }

        // Range reference
        if let Some(range) = parse_range_reference(self.current_sheet, &atom) {
            return Some(Expr::Range(range));
        }

        // Single-cell reference
        if let Some((sheet, row, col)) = parse_single_cell_reference(self.current_sheet, &atom) {
            return Some(Expr::Reference { sheet, row, col });
        }

        Some(Expr::Name(atom))
    }
}

#[cfg(test)]
mod tests {
    use super::{parse_expression, tokenize};

    #[test]
    fn parses_function_with_range_arguments() {
        let expr = parse_expression("Sheet1", "SUMXMY2(A1:A2,B1:B2)");
        assert!(
            expr.is_some(),
            "parser failed on SUMXMY2 with ranges; tokens: {:?}",
            tokenize("SUMXMY2(A1:A2,B1:B2)")
        );
    }
}
