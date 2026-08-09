//! Typed `chart:class` values.

use crate::namespace::CHARTNS;
use litchi_core::{Error, Result};

const MAX_LEXICAL_BYTES: usize = 255;

/// The standard ODF chart class vocabulary from ODF 1.4 part 3, §19.15.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChartClassKind {
    Area,
    Bar,
    Bubble,
    Circle,
    FilledRadar,
    Gantt,
    Line,
    Radar,
    Ring,
    Scatter,
    Stock,
    Surface,
    /// A caller-selected, non-standard namespaced class.
    Extension,
    /// A syntactically valid value retained from an input producer.
    Unknown,
}

/// A bounded, namespace-resolved `chart:class` token retaining its `QName` spelling.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the public `ChartClass` name distinguishes the ODF chart vocabulary type"
)]
pub struct ChartClass {
    kind: ChartClassKind,
    lexical: String,
    namespace_uri: Option<String>,
}

impl ChartClass {
    /// Parse a lexical `QName` with the namespace URI bound to its prefix.
    ///
    /// # Errors
    ///
    /// Returns an error if the `QName` is malformed or has no resolved
    /// namespace URI.
    pub fn parse(
        lexical_input: impl Into<String>,
        resolved_namespace_uri: Option<&str>,
    ) -> Result<Self> {
        let lexical = lexical_input.into();
        validate_qname(&lexical)?;
        let Some(resolved_uri) = resolved_namespace_uri.filter(|uri| !uri.is_empty()) else {
            return invalid("chart class QName has no resolved namespace URI");
        };
        let namespace_uri = Some(resolved_uri.to_owned());
        let local = lexical.rsplit_once(':').map_or("", |(_, local)| local);
        let kind = if namespace_uri.as_deref() == Some(CHARTNS) {
            standard_kind(local).unwrap_or(ChartClassKind::Unknown)
        } else {
            ChartClassKind::Extension
        };
        Ok(Self {
            kind,
            lexical,
            namespace_uri,
        })
    }
    /// Construct a caller-selected extension chart class.
    ///
    /// # Errors
    ///
    /// Returns an error if the `QName` is malformed, its namespace is absent,
    /// or it uses the standard chart namespace.
    pub fn extension(lexical: impl Into<String>, namespace_uri: impl AsRef<str>) -> Result<Self> {
        let mut value = Self::parse(lexical, Some(namespace_uri.as_ref()))?;
        if value.namespace_uri.as_deref() == Some(CHARTNS) {
            return invalid("extension chart class must not use the standard chart namespace");
        }
        value.kind = ChartClassKind::Extension;
        Ok(value)
    }
    /// Construct an explicitly unknown, inert chart class for preservation.
    ///
    /// # Errors
    ///
    /// Returns an error if the `QName` is malformed or its namespace is absent.
    pub fn unknown(lexical: impl Into<String>, namespace_uri: impl AsRef<str>) -> Result<Self> {
        let mut value = Self::parse(lexical, Some(namespace_uri.as_ref()))?;
        value.kind = ChartClassKind::Unknown;
        Ok(value)
    }
    #[must_use]
    pub fn area() -> Self {
        Self::standard(ChartClassKind::Area, "area")
    }
    #[must_use]
    pub fn bar() -> Self {
        Self::standard(ChartClassKind::Bar, "bar")
    }
    #[must_use]
    pub fn bubble() -> Self {
        Self::standard(ChartClassKind::Bubble, "bubble")
    }
    #[must_use]
    pub fn circle() -> Self {
        Self::standard(ChartClassKind::Circle, "circle")
    }
    #[must_use]
    pub fn filled_radar() -> Self {
        Self::standard(ChartClassKind::FilledRadar, "filled-radar")
    }
    #[must_use]
    pub fn gantt() -> Self {
        Self::standard(ChartClassKind::Gantt, "gantt")
    }
    #[must_use]
    pub fn line() -> Self {
        Self::standard(ChartClassKind::Line, "line")
    }
    #[must_use]
    pub fn radar() -> Self {
        Self::standard(ChartClassKind::Radar, "radar")
    }
    #[must_use]
    pub fn ring() -> Self {
        Self::standard(ChartClassKind::Ring, "ring")
    }
    #[must_use]
    pub fn scatter() -> Self {
        Self::standard(ChartClassKind::Scatter, "scatter")
    }
    #[must_use]
    pub fn stock() -> Self {
        Self::standard(ChartClassKind::Stock, "stock")
    }
    #[must_use]
    pub fn surface() -> Self {
        Self::standard(ChartClassKind::Surface, "surface")
    }
    #[must_use]
    pub fn kind(&self) -> ChartClassKind {
        self.kind
    }
    /// The exact producer or caller supplied `QName` spelling.
    #[must_use]
    pub fn lexical(&self) -> &str {
        &self.lexical
    }
    /// The namespace URI resolved for the `QName` prefix, if available.
    #[must_use]
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }
    #[must_use]
    pub(crate) fn namespace_alias(&self) -> Option<(&str, &str)> {
        let (prefix, _) = self.lexical.split_once(':')?;
        self.namespace_uri.as_deref().map(|uri| (prefix, uri))
    }
    fn standard(kind: ChartClassKind, local: &str) -> Self {
        Self {
            kind,
            lexical: format!("chart:{local}"),
            namespace_uri: Some(CHARTNS.to_owned()),
        }
    }
}

fn standard_kind(local: &str) -> Option<ChartClassKind> {
    Some(match local {
        "area" => ChartClassKind::Area,
        "bar" => ChartClassKind::Bar,
        "bubble" => ChartClassKind::Bubble,
        "circle" => ChartClassKind::Circle,
        "filled-radar" => ChartClassKind::FilledRadar,
        "gantt" => ChartClassKind::Gantt,
        "line" => ChartClassKind::Line,
        "radar" => ChartClassKind::Radar,
        "ring" => ChartClassKind::Ring,
        "scatter" => ChartClassKind::Scatter,
        "stock" => ChartClassKind::Stock,
        "surface" => ChartClassKind::Surface,
        _ => return None,
    })
}
fn validate_qname(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_LEXICAL_BYTES || value.matches(':').count() != 1 {
        return invalid(format!("invalid chart class QName '{value}'"));
    }
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid("chart class QName is missing a prefix");
    };
    if !valid_name(prefix) || !valid_name(local) {
        return invalid(format!("invalid chart class QName '{value}'"));
    }
    Ok(())
}
fn valid_name(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::{ChartClass, ChartClassKind};
    use crate::namespace::CHARTNS;
    use litchi_core::Result;

    #[test]
    fn recognizes_every_standard_class_without_normalizing_an_alias() -> Result<()> {
        let classes = [
            ("area", ChartClassKind::Area),
            ("bar", ChartClassKind::Bar),
            ("bubble", ChartClassKind::Bubble),
            ("circle", ChartClassKind::Circle),
            ("filled-radar", ChartClassKind::FilledRadar),
            ("gantt", ChartClassKind::Gantt),
            ("line", ChartClassKind::Line),
            ("radar", ChartClassKind::Radar),
            ("ring", ChartClassKind::Ring),
            ("scatter", ChartClassKind::Scatter),
            ("stock", ChartClassKind::Stock),
            ("surface", ChartClassKind::Surface),
        ];
        for (local, kind) in classes {
            let value = ChartClass::parse(format!("c:{local}"), Some(CHARTNS))?;
            assert_eq!(value.kind(), kind);
            assert_eq!(value.lexical(), format!("c:{local}"));
        }
        Ok(())
    }

    #[test]
    fn extension_and_unknown_are_bounded_and_explicit() -> Result<()> {
        assert_eq!(
            ChartClass::extension("vendor:heat", "urn:vendor")?.kind(),
            ChartClassKind::Extension
        );
        assert_eq!(
            ChartClass::parse("chart:future", Some(CHARTNS))?.kind(),
            ChartClassKind::Unknown
        );
        assert!(ChartClass::parse("chart:bad value", Some(CHARTNS)).is_err());
        Ok(())
    }
}
