//! The raw layer's Python arguments (record 0336).

use litchi::raw::NewPart;

use crate::errors::Failure;
use litchi::Code;

/// `File.add_part`'s part: a typed part, or a custom one with its content
/// type, relationship, payload and name.
pub(crate) fn new_part(
    kind: Option<&str>,
    content_type: Option<String>,
    relationship: Option<String>,
    data: Option<Vec<u8>>,
    name: Option<String>,
) -> Result<NewPart, Failure> {
    match kind {
        Some(kind) if !kind.eq_ignore_ascii_case("custom") => {
            let part = NewPart::of_type(kind).map_err(|error| Failure::read(&error))?;
            Ok(match name {
                Some(name) => part.named(name),
                None => part,
            })
        },
        _ => {
            let missing = |what: &'static str| Failure::new(Code::MissingArgument, what);
            let part = NewPart::new(
                content_type.ok_or_else(|| missing("a custom part needs content_type"))?,
                relationship.ok_or_else(|| missing("a custom part needs relationship"))?,
                data.ok_or_else(|| missing("a custom part needs data"))?,
            );
            Ok(part.named(name.ok_or_else(|| missing("a custom part needs name"))?))
        },
    }
}
