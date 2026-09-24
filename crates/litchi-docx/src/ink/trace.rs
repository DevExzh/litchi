//! Bounded InkML trace-format and trace-data validation for the DOCX profile.
//!
//! The shared DrawingML reader intentionally keeps trace data as source spans.
//! This module validates the imported InkML lexical grammar in place over a
//! stream of decoded character-data bytes. It does not decode points or retain
//! a point vector.

use crate::{Error, Result};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum ChannelType {
    Integer,
    Decimal,
    Double,
    Boolean,
}

#[derive(Debug)]
pub(crate) struct Channel {
    pub(crate) name: String,
    pub(crate) kind: ChannelType,
}

#[derive(Debug, Default)]
pub(crate) struct TraceFormat {
    pub(crate) regular: Vec<Channel>,
    pub(crate) opaque_intermittent: bool,
    pub(crate) context_definition: Option<usize>,
}

impl TraceFormat {
    pub(crate) fn default_xy() -> Self {
        Self {
            regular: vec![
                Channel {
                    name: "X".into(),
                    kind: ChannelType::Decimal,
                },
                Channel {
                    name: "Y".into(),
                    kind: ChannelType::Decimal,
                },
            ],
            opaque_intermittent: false,
            context_definition: None,
        }
    }

    pub(crate) fn total_channels(&self) -> usize {
        self.regular.len()
    }

    pub(crate) fn push(&mut self, channel: Channel) {
        self.regular.push(channel);
    }
}

pub(crate) fn channel_type(value: Option<&str>) -> Result<ChannelType> {
    match value.unwrap_or("decimal") {
        "integer" => Ok(ChannelType::Integer),
        "decimal" => Ok(ChannelType::Decimal),
        "double" => Ok(ChannelType::Double),
        "boolean" => Ok(ChannelType::Boolean),
        value => Err(Error::Invalid(format!(
            "DOCX Ink traceFormat channel type is unsupported: {value}"
        ))),
    }
}

/// Incremental validator used by the DOCX package scanner when XML character
/// data is split across text, CDATA, and entity events.
pub(crate) struct Validator<'a> {
    format: &'a TraceFormat,
    position: usize,
    point: usize,
    values: usize,
    previous: Vec<DifferenceOrder>,
    current: Option<CurrentValue>,
    after_comma: bool,
}

impl<'a> Validator<'a> {
    pub(crate) fn new(format: &'a TraceFormat) -> Result<Self> {
        if format.regular.is_empty() {
            return Err(Error::Invalid(
                "DOCX Ink traceFormat must declare at least one channel".into(),
            ));
        }
        let mut previous = Vec::new();
        previous
            .try_reserve(format.regular.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink trace difference state",
                source,
            })?;
        previous.resize(format.regular.len(), DifferenceOrder::Explicit);
        Ok(Self {
            format,
            position: 0,
            point: 0,
            values: 0,
            previous,
            current: None,
            after_comma: false,
        })
    }

    pub(crate) fn feed(&mut self, bytes: &[u8]) -> Result<()> {
        for &byte in bytes {
            self.feed_byte(byte)?;
            self.position = self.position.checked_add(1).ok_or_else(|| {
                Error::Invalid("DOCX Ink trace lexical position overflowed".into())
            })?;
        }
        Ok(())
    }

    pub(crate) fn finish(&mut self) -> Result<()> {
        if self.current.is_some() {
            self.finish_value()?;
        }
        if self.values != 0 {
            self.finish_point()?;
        } else if self.point == 0 {
            return self.fail("trace data must contain at least one point");
        } else if !self.after_comma {
            return self.fail("trace data ended without a point");
        }
        Ok(())
    }

    fn feed_byte(&mut self, byte: u8) -> Result<()> {
        let mut byte = Some(byte);
        while let Some(next) = byte.take() {
            if self.current.is_none() {
                self.start_value_or_separator(next)?;
                continue;
            }
            if self.consume_current(next)? {
                byte = Some(next);
            }
        }
        Ok(())
    }

    fn start_value_or_separator(&mut self, byte: u8) -> Result<()> {
        if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
            return Ok(());
        }
        if byte == b',' {
            if self.values == 0 {
                return self.fail("trace point is empty or starts with a comma");
            }
            self.finish_point()?;
            self.after_comma = true;
            return Ok(());
        }
        if self.values == self.format.total_channels() && !self.format.opaque_intermittent {
            return self.fail("trace point has more values than its channel format");
        }
        self.after_comma = false;
        let regular = self.values < self.format.regular.len();
        let prefix = if regular && matches!(byte, b'!' | b'\'' | b'"') {
            Some(byte)
        } else {
            None
        };
        if prefix.is_some() {
            self.current = Some(CurrentValue::awaiting(prefix));
            return Ok(());
        }
        if !regular && !self.format.opaque_intermittent && matches!(byte, b'!' | b'\'' | b'"') {
            return self.fail("difference order is valid only for regular channels");
        }
        self.current = Some(CurrentValue::new(byte, None)?);
        Ok(())
    }

    /// Consume one byte for the current value. Returns true when the caller
    /// must reprocess the byte as the first byte of the next value/separator.
    fn consume_current(&mut self, byte: u8) -> Result<bool> {
        let mut current = self.current.take().expect("current value is present");
        let reprocess = match &mut current.kind {
            CurrentKind::Awaiting { sign } => {
                if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
                    false
                } else if *sign && byte == b'-' {
                    return self.fail("numeric trace value has two signs");
                } else if byte == b'-' && !*sign {
                    *sign = true;
                    false
                } else {
                    let next = start_number(byte)?;
                    current.kind = CurrentKind::Number(next);
                    false
                }
            },
            CurrentKind::Number(number) => self.consume_number(number, byte)?,
            CurrentKind::Boolean => true,
            CurrentKind::Wildcard => true,
            CurrentKind::IntermittentWildcard => true,
        };
        if reprocess {
            self.current = Some(current);
            self.finish_value()?;
        } else {
            self.current = Some(current);
        }
        Ok(reprocess)
    }

    fn consume_number(&mut self, number: &mut NumberState, byte: u8) -> Result<bool> {
        match number {
            NumberState::Hex { digits } => {
                if is_hex_digit(byte) {
                    *digits = digits.saturating_add(1);
                    Ok(false)
                } else {
                    if *digits == 0 {
                        return self.fail("hexadecimal trace value has no digits");
                    }
                    Ok(true)
                }
            },
            NumberState::Digits {
                before,
                dot,
                after,
                exponent,
                exponent_digits,
                exponent_sign,
            } => {
                if byte.is_ascii_digit() {
                    if *exponent {
                        *exponent_digits = exponent_digits.saturating_add(1);
                        *exponent_sign = false;
                    } else if *dot {
                        *after = after.saturating_add(1);
                    } else {
                        *before = before.saturating_add(1);
                    }
                    return Ok(false);
                }
                if *exponent {
                    if *exponent_sign && matches!(byte, b'+' | b'-') {
                        *exponent_sign = false;
                        return Ok(false);
                    }
                    if *exponent_digits == 0 {
                        return self.fail("double trace value has no exponent digits");
                    }
                    return Ok(true);
                }
                if byte == b'.' {
                    if *dot {
                        return Ok(true);
                    }
                    *dot = true;
                    return Ok(false);
                }
                if matches!(byte, b'e' | b'E') {
                    *exponent = true;
                    *exponent_sign = true;
                    return Ok(false);
                }
                Ok(true)
            },
            NumberState::LeadingDot {
                digits,
                exponent,
                exponent_digits,
                exponent_sign,
            } => {
                if byte.is_ascii_digit() {
                    if *exponent {
                        *exponent_digits = exponent_digits.saturating_add(1);
                        *exponent_sign = false;
                    } else {
                        *digits = digits.saturating_add(1);
                    }
                    return Ok(false);
                }
                if *exponent {
                    if *exponent_sign && matches!(byte, b'+' | b'-') {
                        *exponent_sign = false;
                        return Ok(false);
                    }
                    if *exponent_digits == 0 {
                        return self.fail("double trace value has no exponent digits");
                    }
                    return Ok(true);
                }
                if matches!(byte, b'e' | b'E') {
                    if *digits == 0 {
                        return self.fail("decimal trace value has no digits");
                    }
                    *exponent = true;
                    *exponent_sign = true;
                    return Ok(false);
                }
                if *digits == 0 {
                    return self.fail("decimal trace value has no digits");
                }
                Ok(true)
            },
        }
    }

    fn finish_value(&mut self) -> Result<()> {
        let current = self
            .current
            .take()
            .ok_or_else(|| Error::Invalid("DOCX Ink trace value is missing".into()))?;
        let channel_index = self.values;
        let regular = channel_index < self.format.regular.len();
        let Some(channel) = self.format.regular.get(channel_index) else {
            if !self.format.opaque_intermittent {
                return self.fail("trace value channel is out of range");
            }
            match current.kind {
                CurrentKind::Awaiting { .. } => {
                    return self.fail("trace value is incomplete");
                },
                CurrentKind::Number(number) if !number.valid() => {
                    return self.fail("trace numeric value is lexically incomplete");
                },
                CurrentKind::Number(_)
                | CurrentKind::Boolean
                | CurrentKind::Wildcard
                | CurrentKind::IntermittentWildcard => {},
            }
            self.values = self
                .values
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("DOCX Ink trace value count overflowed".into()))?;
            return Ok(());
        };
        match current.kind {
            CurrentKind::Awaiting { .. } => {
                return self.fail("trace value is incomplete");
            },
            CurrentKind::Boolean => {
                if channel.name == "T" {
                    return self.fail("the T channel requires integer millisecond values");
                }
                if channel.kind != ChannelType::Boolean {
                    return self.fail("boolean trace values require a boolean channel");
                }
            },
            CurrentKind::Wildcard => {
                if regular {
                    if self.point == 0 {
                        return self
                            .fail("the first trace point cannot wildcard a regular channel");
                    }
                }
            },
            CurrentKind::IntermittentWildcard => {
                if regular {
                    return self
                        .fail("the intermittent wildcard is not valid for a regular channel");
                }
            },
            CurrentKind::Number(number) => {
                if !number.valid() {
                    return self.fail("trace numeric value is lexically incomplete");
                }
                if channel.kind == ChannelType::Boolean {
                    return self.fail("a boolean channel requires T or F values");
                }
                let (integer, exponent) = number.kind();
                if channel.name == "T" {
                    if !integer {
                        return self.fail("the T channel requires integer millisecond values");
                    }
                } else if channel.kind == ChannelType::Integer && !integer {
                    return self.fail("an integer channel requires integer values");
                }
                if channel.kind == ChannelType::Decimal && exponent {
                    return self.fail("a decimal channel cannot use an exponent");
                }
                if regular {
                    let prefix = current.prefix;
                    if self.point == 0 && prefix.is_some() && prefix != Some(b'!') {
                        return self.fail("the first trace point must use explicit regular values");
                    }
                    if self.point == 0 {
                        self.previous[channel_index] = DifferenceOrder::Explicit;
                    } else {
                        match prefix {
                            None => {},
                            Some(b'!') => {
                                self.previous[channel_index] = DifferenceOrder::Explicit;
                            },
                            Some(b'\'') => {
                                if self.previous[channel_index] != DifferenceOrder::Explicit {
                                    return self.fail(
                                        "first-difference values require a preceding explicit value",
                                    );
                                }
                                self.previous[channel_index] = DifferenceOrder::First;
                            },
                            Some(b'"') => {
                                if self.previous[channel_index] != DifferenceOrder::First {
                                    return self.fail(
                                        "second-difference values require a preceding first difference",
                                    );
                                }
                                self.previous[channel_index] = DifferenceOrder::Second;
                            },
                            Some(_) => unreachable!("prefix is restricted above"),
                        }
                    }
                } else if current.prefix.is_some() {
                    return self.fail("difference order is valid only for regular channels");
                }
            },
        }
        self.values = self
            .values
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("DOCX Ink trace value count overflowed".into()))?;
        Ok(())
    }

    fn finish_point(&mut self) -> Result<()> {
        if self.values == 0 {
            return self.fail("trace point is empty");
        }
        if self.values < self.format.regular.len() {
            return self.fail("trace point does not provide every regular channel");
        }
        self.point = self
            .point
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("DOCX Ink trace point count overflowed".into()))?;
        self.values = 0;
        Ok(())
    }

    fn fail<T>(&self, message: &str) -> Result<T> {
        Err(Error::Invalid(format!(
            "DOCX Ink trace data at byte {}: {message}",
            self.position
        )))
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DifferenceOrder {
    Explicit,
    First,
    Second,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct CurrentValue {
    prefix: Option<u8>,
    kind: CurrentKind,
}

impl CurrentValue {
    const fn awaiting(prefix: Option<u8>) -> Self {
        Self {
            prefix,
            kind: CurrentKind::Awaiting { sign: false },
        }
    }

    fn new(byte: u8, prefix: Option<u8>) -> Result<Self> {
        let kind = match byte {
            b'T' | b'F' => CurrentKind::Boolean,
            b'*' => CurrentKind::Wildcard,
            b'?' => CurrentKind::IntermittentWildcard,
            b'-' => CurrentKind::Awaiting { sign: true },
            _ => CurrentKind::Number(start_number(byte)?),
        };
        Ok(Self { prefix, kind })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CurrentKind {
    Awaiting { sign: bool },
    Number(NumberState),
    Boolean,
    Wildcard,
    IntermittentWildcard,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NumberState {
    Hex {
        digits: usize,
    },
    Digits {
        before: usize,
        dot: bool,
        after: usize,
        exponent: bool,
        exponent_digits: usize,
        exponent_sign: bool,
    },
    LeadingDot {
        digits: usize,
        exponent: bool,
        exponent_digits: usize,
        exponent_sign: bool,
    },
}

impl NumberState {
    fn valid(self) -> bool {
        match self {
            Self::Hex { digits } => digits > 0,
            Self::Digits {
                before,
                exponent,
                exponent_digits,
                ..
            } => before > 0 && (!exponent || exponent_digits > 0),
            Self::LeadingDot {
                digits,
                exponent,
                exponent_digits,
                ..
            } => digits > 0 && (!exponent || exponent_digits > 0),
        }
    }

    fn kind(self) -> (bool, bool) {
        match self {
            Self::Hex { .. } => (true, false),
            Self::Digits {
                before,
                dot,
                exponent,
                ..
            } => (!dot && !exponent && before > 0, exponent),
            Self::LeadingDot { exponent, .. } => (false, exponent),
        }
    }
}

fn start_number(byte: u8) -> Result<NumberState> {
    if byte == b'#' {
        return Ok(NumberState::Hex { digits: 0 });
    }
    if byte == b'.' {
        return Ok(NumberState::LeadingDot {
            digits: 0,
            exponent: false,
            exponent_digits: 0,
            exponent_sign: false,
        });
    }
    if byte.is_ascii_digit() {
        return Ok(NumberState::Digits {
            before: 1,
            dot: false,
            after: 0,
            exponent: false,
            exponent_digits: 0,
            exponent_sign: false,
        });
    }
    Err(Error::Invalid(
        "DOCX Ink trace value is not a decimal, double, or hexadecimal number".into(),
    ))
}

fn is_hex_digit(byte: u8) -> bool {
    byte.is_ascii_digit() || matches!(byte, b'A'..=b'F')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn t_boolean_value_is_rejected_even_if_the_declared_kind_is_boolean() {
        let format = TraceFormat {
            regular: vec![Channel {
                name: "T".into(),
                kind: ChannelType::Boolean,
            }],
            opaque_intermittent: false,
            context_definition: None,
        };
        let mut validator = Validator::new(&format).unwrap();
        validator.feed(b"T").unwrap();
        assert!(validator.finish().is_err());
    }
}
