//! Admission for programmatically constructed queries, before recursive compilation.
use std::fmt::{self, Write as _};

use crate::{Code, Error, Result};
use litchi_core::selector::{Complex, Condition, Expr, Predicate, Pseudo, Query, Scope, limits};

fn limit() -> Error {
    Error::new(
        Code::LimitExceeded,
        "query exceeds a structural or byte limit",
    )
}
fn depth(depth: usize) -> Result<()> {
    if depth > limits::MAX_NESTING {
        return Err(Error::new(
            Code::MaxDepthExceeded,
            "query exceeds the nesting limit",
        ));
    }
    Ok(())
}
struct Admission {
    nodes: usize,
    bytes: usize,
}
impl Admission {
    fn count(&mut self, bytes: usize) -> Result<()> {
        self.nodes = self.nodes.saturating_add(1);
        self.bytes = self.bytes.saturating_add(bytes);
        if self.nodes > 4096 || self.bytes > limits::MAX_INPUT_BYTES {
            return Err(limit());
        }
        Ok(())
    }
    fn expression(&mut self, expr: &Expr<'_>, nesting: usize, terms: &mut usize) -> Result<()> {
        depth(nesting)?;
        self.count(0)?;
        match expr {
            Expr::Condition(condition) => {
                *terms += 1;
                if *terms > limits::MAX_TERMS {
                    return Err(limit());
                }
                self.count(condition.key().name().len())?;
                if let Condition::Compare { value, .. } = condition {
                    self.count(value.as_str().len())?;
                }
            },
            Expr::And(parts) | Expr::Or(parts) => {
                if parts.is_empty() || parts.len() > limits::MAX_TERMS {
                    return Err(limit());
                }
                for part in parts {
                    self.expression(part, nesting + 1, terms)?;
                }
            },
            Expr::Not(inner) => self.expression(inner, nesting + 1, terms)?,
            _ => return Err(limit()),
        }
        Ok(())
    }
    fn list(&mut self, list: &[Complex<'_>], nesting: usize) -> Result<()> {
        depth(nesting)?;
        if list.is_empty() || list.len() > limits::MAX_ALTERNATIVES {
            return Err(limit());
        }
        for complex in list {
            if complex.rest().len() >= limits::MAX_COMPOUNDS {
                return Err(limit());
            }
            for compound in complex.compounds() {
                self.count(0)?;
                if let litchi_core::selector::Element::Named(name) = compound.element() {
                    self.count(name.len())?;
                }
                if compound
                    .predicates()
                    .len()
                    .saturating_add(compound.pseudos().len())
                    > limits::MAX_PREDICATES
                {
                    return Err(limit());
                }
                for predicate in compound.predicates() {
                    match predicate {
                        Predicate::Filter(expr) => self.expression(expr, nesting, &mut 0)?,
                        Predicate::Index(position) if position.one_based().is_none() => {
                            return Err(limit());
                        },
                        _ => {},
                    }
                }
                for pseudo in compound.pseudos() {
                    match pseudo {
                        Pseudo::Has(list) | Pseudo::Not(list) => self.list(list, nesting + 1)?,
                        Pseudo::Contains(text) => self.count(text.len())?,
                        _ => self.count(0)?,
                    }
                }
            }
        }
        Ok(())
    }
}

pub(super) fn validate(query: &Query<'_>) -> Result<()> {
    let mut admission = Admission { nodes: 0, bytes: 0 };
    if let Some(scope) = query.scope() {
        match scope {
            Scope::Path(path) => {
                if !path.is_absolute() {
                    return Err(Error::new(
                        Code::InvalidSelector,
                        "query scope paths must be absolute",
                    ));
                }
                if path.steps().len() > limits::MAX_STEPS {
                    return Err(Error::new(
                        Code::MaxDepthExceeded,
                        "query scope exceeds the path depth limit",
                    ));
                }
                crate::path::arg::validate(&crate::path::PathArg::from(&**path))?;
            },
            Scope::Sheet(name) => {
                if name.is_empty() || name.chars().any(char::is_control) {
                    return Err(Error::new(
                        Code::InvalidSelector,
                        "invalid query worksheet scope",
                    ));
                }
                admission.count(name.len())?;
            },
            _ => return Err(limit()),
        }
    }
    admission.list(query.alternatives(), 0)?;
    struct Bound(usize);
    impl fmt::Write for Bound {
        fn write_str(&mut self, text: &str) -> fmt::Result {
            self.0 = self.0.saturating_add(text.len());
            if self.0 > limits::MAX_INPUT_BYTES {
                Err(fmt::Error)
            } else {
                Ok(())
            }
        }
    }
    write!(Bound(0), "{query}").map_err(|_| limit())
}
