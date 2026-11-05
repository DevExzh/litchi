//! Text encoding and start-tag rewriting for scene splices.

use litchi_core::{Error, Result};

/// Inline ODF content for one paragraph's text.
pub(super) struct Encoded {
    /// The escaped content.
    pub(super) markup: String,
    /// Whether the content is character data only (no `text:s`,
    /// `text:tab` or `text:line-break`).
    pub(super) plain: bool,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Escape an attribute value for double quotes.
pub(super) fn escape_attribute(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#9;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            _ => out.push(character),
        }
    }
    out
}

/// Whether `character` may appear in XML 1.0 character data.
const fn xml_char(character: char) -> bool {
    matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

/// Encode `text` as paragraph content that projects back to exactly
/// `text` under the slide parser's whitespace rules: a space run becomes a
/// literal space plus `text:s` where collapsing would drop it, `\t` is
/// `text:tab` and `\n` (or `\r\n`) is `text:line-break`.
pub(super) fn encode(prefix: &str, text: &str) -> Result<Encoded> {
    let mut markup = String::with_capacity(text.len() + 8);
    let mut plain = true;
    let characters: Vec<char> = text.chars().collect();
    let mut index = 0;
    // Whether the projection's last character is whitespace (or nothing),
    // so a literal space here would be collapsed away.
    let mut after_space = true;
    while index < characters.len() {
        let character = characters[index];
        if !xml_char(character) {
            return Err(invalid(format!(
                "ODP text cannot contain the control character U+{:04X}",
                u32::from(character)
            )));
        }
        match character {
            ' ' => {
                let mut count = 0;
                while index < characters.len() && characters[index] == ' ' {
                    count += 1;
                    index += 1;
                }
                let at_end = index == characters.len();
                let literal = !after_space && !at_end;
                if literal {
                    markup.push(' ');
                    count -= 1;
                }
                if count > 0 {
                    plain = false;
                    if count == 1 {
                        markup.push_str(&format!("<{prefix}:s/>"));
                    } else {
                        markup.push_str(&format!("<{prefix}:s {prefix}:c=\"{count}\"/>"));
                    }
                }
                after_space = true;
                continue;
            },
            '\t' => {
                plain = false;
                markup.push_str(&format!("<{prefix}:tab/>"));
                after_space = true;
            },
            '\r' if characters.get(index + 1) == Some(&'\n') => {},
            '\n' | '\r' => {
                plain = false;
                markup.push_str(&format!("<{prefix}:line-break/>"));
                after_space = true;
            },
            character if character.is_whitespace() => {
                // Other white space (no-break space and the like) is kept
                // as written; the projection treats it as a collapsible
                // space, as the slide parser does.
                markup.push(character);
                after_space = true;
            },
            '&' => {
                markup.push_str("&amp;");
                after_space = false;
            },
            '<' => {
                markup.push_str("&lt;");
                after_space = false;
            },
            '>' => {
                markup.push_str("&gt;");
                after_space = false;
            },
            character => {
                markup.push(character);
                after_space = false;
            },
        }
        index += 1;
    }
    Ok(Encoded { markup, plain })
}

/// One attribute of a start tag, as written.
struct RawAttribute<'a> {
    key: &'a str,
    value: &'a str,
    quote: char,
}

/// A parsed start tag.
pub(super) struct StartTag<'a> {
    pub(super) name: &'a str,
    attributes: Vec<RawAttribute<'a>>,
}

impl<'a> StartTag<'a> {
    /// Read a start tag (`<name a="v">` or `<name a="v"/>`) as written.
    pub(super) fn parse(raw: &'a str) -> Result<Self> {
        let malformed = || invalid("ODP scene start tag is malformed");
        let body = raw.strip_prefix('<').ok_or_else(malformed)?;
        let body = body.strip_suffix('>').ok_or_else(malformed)?;
        let body = body.strip_suffix('/').unwrap_or(body);
        let name_end = body
            .find(|character: char| character.is_ascii_whitespace())
            .unwrap_or(body.len());
        let name = &body[..name_end];
        if name.is_empty() {
            return Err(malformed());
        }
        let mut rest = &body[name_end..];
        let mut attributes = Vec::new();
        loop {
            rest = rest.trim_start_matches(|character: char| character.is_ascii_whitespace());
            if rest.is_empty() {
                break;
            }
            let equals = rest.find('=').ok_or_else(malformed)?;
            let key = rest[..equals].trim_end();
            rest = rest[equals + 1..].trim_start();
            let quote = rest.chars().next().ok_or_else(malformed)?;
            if quote != '"' && quote != '\'' {
                return Err(malformed());
            }
            let value_end = rest[1..].find(quote).ok_or_else(malformed)? + 1;
            let value = &rest[1..value_end];
            rest = &rest[value_end + 1..];
            attributes.push(RawAttribute { key, value, quote });
        }
        Ok(Self { name, attributes })
    }

    /// The qualified key of the first attribute whose local name is
    /// `local`.
    pub(super) fn key_of_local(&self, local: &str) -> Option<&'a str> {
        self.attributes
            .iter()
            .map(|attribute| attribute.key)
            .find(|key| key.rsplit(':').next() == Some(local))
    }

    /// The prefix an `xmlns:` declaration of the tag binds to `uri`.
    pub(super) fn prefix_of(&self, uri: &str) -> Option<&'a str> {
        self.attributes.iter().find_map(|attribute| {
            attribute
                .key
                .strip_prefix("xmlns:")
                .filter(|_| attribute.value == uri)
        })
    }

    /// Write the tag with `changes` applied (`None` removes an attribute;
    /// an absent attribute is appended), as an open tag or a self-closing
    /// one, dropping identity attributes when `strip_ids` is set.
    pub(super) fn write(
        &self,
        changes: &[(String, Option<String>)],
        self_closing: bool,
        strip_ids: bool,
    ) -> String {
        let mut out = String::with_capacity(64);
        out.push('<');
        out.push_str(self.name);
        for attribute in &self.attributes {
            if strip_ids && is_identity(attribute.key) {
                continue;
            }
            if let Some((_, value)) = changes.iter().find(|(key, _)| key == attribute.key) {
                if let Some(value) = value {
                    push_attribute(&mut out, attribute.key, &escape_attribute(value));
                }
                continue;
            }
            if attribute.quote == '\'' && attribute.value.contains('"') {
                push_attribute(
                    &mut out,
                    attribute.key,
                    &attribute.value.replace('"', "&quot;"),
                );
            } else {
                push_attribute(&mut out, attribute.key, attribute.value);
            }
        }
        for (key, value) in changes {
            if self.attributes.iter().any(|attribute| attribute.key == key) {
                continue;
            }
            if let Some(value) = value {
                push_attribute(&mut out, key, &escape_attribute(value));
            }
        }
        out.push_str(if self_closing { "/>" } else { ">" });
        out
    }
}

fn push_attribute(out: &mut String, key: &str, value: &str) {
    out.push(' ');
    out.push_str(key);
    out.push_str("=\"");
    out.push_str(value);
    out.push('"');
}

/// Identity attributes a copied template must not duplicate
/// (`xml:id`, `draw:id`, `text:id`).
fn is_identity(key: &str) -> bool {
    key.ends_with(":id")
}

/// The open-tag form of a start tag, rewritten only when it was
/// self-closing.
pub(super) fn open_tag(raw: &str) -> Result<String> {
    if raw.ends_with("/>") {
        Ok(StartTag::parse(raw)?.write(&[], false, false))
    } else {
        Ok(raw.to_owned())
    }
}

/// A copy of a start tag for a new sibling: open form, identities dropped.
pub(super) fn template(raw: &str) -> Result<String> {
    Ok(StartTag::parse(raw)?.write(&[], false, true))
}
