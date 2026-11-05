//! Shared validation-prefix grammar for streaming reads and source edits.
//! Identification metadata stays opaque; ACDVALLIST supplies a list override.

use super::{Error, Result, parse_dval_list};
use crate::raw::{Kind, Limits, kind};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Step {
    Other,
    Begin,
    Content,
    End { list: bool },
    BareList,
}

#[derive(Default)]
pub(crate) struct Prelude {
    alternate: bool,
    alternate_content: Option<bool>,
    alternate_count: u8,
    pending: bool,
    list: Option<String>,
}

fn invalid(message: &'static str) -> Error {
    Error::Unrecognized {
        typ: "data-validation prelude",
        val: message,
    }
}

impl Prelude {
    pub(crate) fn feed(&mut self, typ: Kind, data: &[u8], limits: Limits) -> Result<Step> {
        if self.alternate {
            match typ {
                kind::AC_END => {
                    if !data.is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: data.len(),
                        });
                    }
                    let list = self
                        .alternate_content
                        .take()
                        .ok_or_else(|| invalid("empty alternate block"))?;
                    self.alternate = false;
                    Ok(Step::End { list })
                },
                kind::AC_BEGIN
                | kind::FRT_BEGIN
                | kind::FRT_END
                | kind::D_VAL
                | kind::D_VAL14
                | kind::BEGIN_D_VALS
                | kind::END_D_VALS
                | kind::BEGIN_D_VALS14
                | kind::END_D_VALS14 => {
                    Err(invalid("alternate block interrupts the validation grammar"))
                },
                _ => {
                    if self.alternate_content == Some(true)
                        || typ == kind::D_VAL_LIST && self.alternate_content.is_some()
                    {
                        return Err(invalid(
                            "DValList must be the only record in its alternate block",
                        ));
                    }
                    if typ == kind::D_VAL_LIST {
                        self.set_list(data, limits)?;
                    }
                    self.alternate_content = Some(typ == kind::D_VAL_LIST);
                    Ok(Step::Content)
                },
            }
        } else {
            match typ {
                kind::AC_BEGIN => {
                    if data.len() < 2 {
                        return Err(Error::InvalidLength {
                            expected: 2,
                            found: data.len(),
                        });
                    }
                    let count = usize::from(u16::from_le_bytes([data[0], data[1]]));
                    if count == 0 {
                        return Err(invalid("ACBegin has no product version"));
                    }
                    let expected = 2 + count * 4;
                    if data.len() != expected {
                        return Err(Error::InvalidLength {
                            expected,
                            found: data.len(),
                        });
                    }
                    // The worksheet grammar permits [ACUID] [ACDVALLIST].
                    if self.alternate_count == 2 {
                        return Err(invalid("too many alternate blocks before one validation"));
                    }
                    self.alternate_count += 1;
                    self.alternate = true;
                    self.pending = true;
                    Ok(Step::Begin)
                },
                kind::D_VAL_LIST => {
                    self.set_list(data, limits)?;
                    self.pending = true;
                    Ok(Step::BareList)
                },
                _ => Ok(Step::Other),
            }
        }
    }

    fn set_list(&mut self, data: &[u8], limits: Limits) -> Result<()> {
        if self.list.is_some() {
            return Err(invalid("duplicate list override"));
        }
        self.list = Some(parse_dval_list(data, limits)?);
        Ok(())
    }

    pub(crate) fn take_list(&mut self) -> Result<Option<String>> {
        if self.alternate {
            return Err(invalid("unterminated alternate block"));
        }
        self.pending = false;
        self.alternate_count = 0;
        Ok(self.list.take())
    }

    pub(crate) fn finish(&self) -> Result<()> {
        if self.pending || self.alternate {
            return Err(invalid("prelude is not followed by a validation"));
        }
        Ok(())
    }
}
