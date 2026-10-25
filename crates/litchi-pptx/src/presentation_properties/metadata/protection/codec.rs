//! Bounded protection XML codec.

use super::super::{BoundedString, Limits, process_str};
use super::model::{Algorithm, Settings, Verifier};
use crate::{Error, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64_ENGINE;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use rand::TryRng;
use rand::rngs::SysRng;
use sha2::digest::Output;
use sha2::{Digest, Sha512};
use zeroize::{Zeroize, Zeroizing};

const SPIN_COUNT: u32 = 100_000;

struct Sha512Output(Output<Sha512>);

impl Zeroize for Sha512Output {
    fn zeroize(&mut self) {
        self.0.as_mut_slice().fill(0);
    }
}

fn password_hash(password: &str, salt: &[u8], spin_count: u32) -> Zeroizing<Sha512Output> {
    let mut hasher = Sha512::new();
    for unit in password.encode_utf16() {
        hasher.update(unit.to_le_bytes());
    }
    hasher.update(salt);
    let mut hash = Zeroizing::new(Sha512Output(Output::<Sha512>::default()));
    hasher.finalize_into(&mut hash.0);

    for iteration in 0..spin_count {
        let mut hasher = Sha512::new();
        hasher.update(hash.0.as_slice());
        hasher.update(iteration.to_le_bytes());
        hasher.finalize_into(&mut hash.0);
    }
    hash
}

pub(crate) fn generate_verifier(password: &str, limits: Limits) -> Result<Verifier> {
    if password.len() > limits.max_string_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PasswordBytes),
            observed: password.len(),
            maximum: limits.max_string_bytes(),
        });
    }

    let mut salt = [0u8; 16];
    let mut rng = SysRng;
    rng.try_fill_bytes(&mut salt).map_err(|error| {
        Error::Invalid(format!(
            "failed to generate random salt for modify password: {error}"
        ))
    })?;
    let hash = password_hash(password, &salt, SPIN_COUNT);
    Ok(Verifier {
        algorithm: Algorithm::Sha512,
        spin_count: SPIN_COUNT,
        hash: BASE64_ENGINE.encode(hash.0.as_slice()),
        salt: BASE64_ENGINE.encode(salt),
    })
}

impl Settings {
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_xml(xml: &str, limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                observed: xml.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        let processed = process_str(xml, limits)?;
        if processed.len() > limits.max_output_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: processed.len(),
                maximum: limits.max_output_bytes(),
            });
        }
        let mut reader = Reader::from_str(processed.as_ref());
        reader.config_mut().trim_text(true);
        let mut settings = Self::new();
        let mut seen = false;
        let mut nodes = 0usize;
        let mut depth = 0usize;
        loop {
            match reader.read_event()? {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: usize::MAX,
                        maximum: limits.max_depth(),
                    })?;
                    if depth > limits.max_depth() {
                        return Err(Error::Limit {
                            kind: crate::opened::LimitKind::Codec(
                                crate::opened::CodecLimitKind::XmlDepth,
                            ),
                            observed: depth,
                            maximum: limits.max_depth(),
                        });
                    }
                    nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: usize::MAX,
                        maximum: limits.max_nodes(),
                    })?;
                    if nodes > limits.max_nodes() {
                        return Err(Error::Limit {
                            kind: crate::opened::LimitKind::Codec(
                                crate::opened::CodecLimitKind::XmlNodes,
                            ),
                            observed: nodes,
                            maximum: limits.max_nodes(),
                        });
                    }
                    check_attributes(&element, reader.decoder(), limits)?;
                    if element.local_name().as_ref() == b"modifyVerifier" {
                        if seen {
                            return Err(Error::Invalid(
                                "duplicate presentation modifyVerifier".into(),
                            ));
                        }
                        seen = true;
                        settings.modify = Some(parse_verifier(&element, reader.decoder(), limits)?);
                    }
                },
                Event::Empty(element) => {
                    check_empty_depth(depth, limits)?;
                    nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: usize::MAX,
                        maximum: limits.max_nodes(),
                    })?;
                    if nodes > limits.max_nodes() {
                        return Err(Error::Limit {
                            kind: crate::opened::LimitKind::Codec(
                                crate::opened::CodecLimitKind::XmlNodes,
                            ),
                            observed: nodes,
                            maximum: limits.max_nodes(),
                        });
                    }
                    check_attributes(&element, reader.decoder(), limits)?;
                    if element.local_name().as_ref() == b"modifyVerifier" {
                        if seen {
                            return Err(Error::Invalid(
                                "duplicate presentation modifyVerifier".into(),
                            ));
                        }
                        seen = true;
                        settings.modify = Some(parse_verifier(&element, reader.decoder(), limits)?);
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("protection XML has an unmatched end tag".into())
                    })?;
                },
                Event::Eof => break,
                Event::DocType(_) | Event::PI(_) => {
                    return Err(Error::Invalid(
                        "protection XML cannot contain DTDs or processing instructions".into(),
                    ));
                },
                _ => {},
            }
        }
        if depth != 0 {
            return Err(Error::Invalid(
                "protection XML has an unterminated element".into(),
            ));
        }
        Ok(settings)
    }

    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        let Some(verifier) = &self.modify else {
            return Ok(String::new());
        };
        let mut value = BoundedString::new(limits, "serialized protection XML");
        value.write_fmt(format_args!(
            r#"<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="{}" spinCount="{}" saltData="{}" hashData="{}"/>"#,
            verifier.algorithm.sid(),
            verifier.spin_count,
            verifier.salt,
            verifier.hash,
        ))?;
        Ok(value.finish())
    }

    pub fn to_pres_props_xml(&self, limits: Limits) -> Result<String> {
        if !self.read_only_recommended {
            return Ok(String::new());
        }
        let mut value =
            BoundedString::new(limits, "serialized protection presentation-properties XML");
        value.write_fmt(format_args!(
            r#"<p:extLst><p:ext uri="{}"><p1710:readonlyRecommended xmlns:p1710="{}" val="1"/></p:ext></p:extLst>"#,
            crate::presentation_properties::READONLY_RECOMMENDED_URI,
            crate::presentation_properties::P1710_NS,
        ))?;
        Ok(value.finish())
    }
}

fn parse_verifier(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Verifier> {
    let mut hash = None;
    let mut salt = None;
    let mut spins = None;
    let mut algorithm = None;
    let mut attributes = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        attributes = attributes.checked_add(1).ok_or_else(|| Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
            observed: usize::MAX,
            maximum: limits.max_attributes(),
        })?;
        if attributes > limits.max_attributes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
                observed: attributes,
                maximum: limits.max_attributes(),
            });
        }
        if attribute.key.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.key.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.value.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::StringBytes),
                observed: value.len(),
                maximum: limits.max_string_bytes(),
            });
        }
        let value = value.into_owned();
        match attribute.key.as_ref() {
            b"hashValue" | b"hashData" => set_once(&mut hash, value, "hash")?,
            b"saltValue" | b"saltData" => set_once(&mut salt, value, "salt")?,
            b"spinCount" | b"spinValue" => {
                let value = value
                    .parse::<u32>()
                    .map_err(|_err| Error::Invalid("protection spin count is not a u32".into()))?;
                if !(1..=10_000_000).contains(&value) {
                    return Err(Error::Invalid(
                        "protection spin count must be between 1 and 10000000".into(),
                    ));
                }
                set_once(&mut spins, value, "spin count")?;
            },
            b"algorithmName" => {
                set_once(&mut algorithm, Algorithm::from_uri(&value)?, "algorithm")?;
            },
            b"cryptAlgorithmSid" => {
                let sid = value.parse::<u32>().map_err(|_err| {
                    Error::Invalid("protection algorithm SID is not a u32".into())
                })?;
                set_once(&mut algorithm, Algorithm::from_sid(sid)?, "algorithm")?;
            },
            b"algIdExt" => {
                return Err(Error::Invalid(
                    "extended CryptoAPI protection algorithms are unsupported".into(),
                ));
            },
            _ => {},
        }
    }
    let hash = hash.ok_or_else(|| Error::Invalid("modifyVerifier is missing its hash".into()))?;
    let salt = salt.ok_or_else(|| Error::Invalid("modifyVerifier is missing its salt".into()))?;
    let spin_count =
        spins.ok_or_else(|| Error::Invalid("modifyVerifier is missing its spin count".into()))?;
    let algorithm = algorithm
        .ok_or_else(|| Error::Invalid("modifyVerifier is missing its algorithm".into()))?;
    if hash.len() > 128 || salt.len() > 1_368 {
        return Err(Error::Invalid(
            "protection Base64 field exceeds its bound".into(),
        ));
    }
    let decoded_hash = decode_base64(&hash)?;
    if decoded_hash.len() != algorithm.output_bytes() {
        return Err(Error::Invalid(format!(
            "protection hash has {} bytes, expected {}",
            decoded_hash.len(),
            algorithm.output_bytes()
        )));
    }
    let decoded_salt = decode_base64(&salt)?;
    if decoded_salt.is_empty() || decoded_salt.len() > 1_024 {
        return Err(Error::Invalid(
            "protection salt must contain 1 to 1024 bytes".into(),
        ));
    }
    Ok(Verifier {
        algorithm,
        spin_count,
        hash,
        salt,
    })
}

fn check_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    let mut attributes = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        attributes = attributes.checked_add(1).ok_or_else(|| Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
            observed: usize::MAX,
            maximum: limits.max_attributes(),
        })?;
        if attributes > limits.max_attributes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
                observed: attributes,
                maximum: limits.max_attributes(),
            });
        }
        if attribute.key.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.key.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.value.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::StringBytes),
                observed: value.len(),
                maximum: limits.max_string_bytes(),
            });
        }
    }
    Ok(())
}

fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
    let observed = depth.checked_add(1).unwrap_or(usize::MAX);
    if observed > limits.max_depth() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
            observed,
            maximum: limits.max_depth(),
        });
    }
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T, field: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::Invalid(format!("duplicate protection {field}")));
    }
    Ok(())
}

fn decode_base64(value: &str) -> Result<Vec<u8>> {
    let bytes = value.as_bytes();
    if bytes.is_empty() || !bytes.len().is_multiple_of(4) {
        return Err(Error::Invalid(
            "protection field is not valid Base64".into(),
        ));
    }
    let mut output = Vec::with_capacity(bytes.len() / 4 * 3);
    for chunk in bytes.as_chunks::<4>().0 {
        let a = sextet(chunk[0])?;
        let b = sextet(chunk[1])?;
        let c = if chunk[2] == b'=' {
            0
        } else {
            sextet(chunk[2])?
        };
        let d = if chunk[3] == b'=' {
            0
        } else {
            sextet(chunk[3])?
        };
        output.push((a << 2) | (b >> 4));
        if chunk[2] != b'=' {
            output.push((b << 4) | (c >> 2));
        }
        if chunk[3] != b'=' {
            output.push((c << 6) | d);
        }
    }
    Ok(output)
}

fn sextet(value: u8) -> Result<u8> {
    let value = match value {
        b'A'..=b'Z' => value - b'A',
        b'a'..=b'z' => value - b'a' + 26,
        b'0'..=b'9' => value - b'0' + 52,
        b'+' => 62,
        b'/' => 63,
        _ => {
            return Err(Error::Invalid(
                "protection field is not valid Base64".into(),
            ));
        },
    };
    Ok(value)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_republishes_a_legacy_verifier() {
        let xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:modifyVerifier cryptAlgorithmSid="14" spinCount="1000" saltData="AA==" hashData="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="/></p:presentation>"#;
        let settings = Settings::parse_xml(xml, Limits::DEFAULT).expect("verifier");
        assert_eq!(
            settings.modify().expect("modifier").algorithm(),
            Algorithm::Sha512
        );
        assert!(
            settings
                .to_xml(Limits::DEFAULT)
                .expect("serialize verifier")
                .contains("cryptAlgorithmSid=\"14\"")
        );
    }

    #[test]
    fn password_generation_is_explicitly_dependency_bound() {
        let mut settings = Settings::new();
        settings
            .set_modify_password("secret", Limits::DEFAULT)
            .unwrap();
        let verifier = settings.modify().unwrap();
        assert_eq!(verifier.algorithm(), Algorithm::Sha512);
        assert_eq!(verifier.spins(), SPIN_COUNT);
        assert_eq!(BASE64_ENGINE.decode(verifier.hash()).unwrap().len(), 64);
        assert_eq!(BASE64_ENGINE.decode(verifier.salt()).unwrap().len(), 16);
        assert!(!format!("{settings:?}").contains("secret"));
    }

    #[test]
    fn password_hash_matches_office_password_then_salt_order() {
        let hash = password_hash("Päss😀", &(0..16).collect::<Vec<_>>(), 2);
        assert_eq!(
            BASE64_ENGINE.encode(hash.0.as_slice()),
            "3ACFcYR0/M+PsEwOXR4/mcgYsTN1VMXMunIrbpt1lY+1Kal3nCkZOJjIEw+LWRlQzI3rL5HZnVIoL87I6R8tNw=="
        );
    }

    #[test]
    fn empty_protection_depth_is_checked_before_attributes() {
        let limits = Limits::DEFAULT.with_depth(1).expect("depth limit");
        let exact = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        Settings::parse_xml(std::str::from_utf8(exact).expect("UTF-8"), limits)
            .expect("depth-one empty root");

        let one_over = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:holder/></p:presentation>"#;
        assert_xml_depth_limit(
            Settings::parse_xml(std::str::from_utf8(one_over).expect("UTF-8"), limits)
                .expect_err("depth-two empty element"),
            2,
            1,
        );
        assert!(matches!(
            Settings::parse_xml(
                "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                Limits::DEFAULT,
            ),
            Err(Error::Invalid(_))
        ));
    }

    fn assert_xml_depth_limit(error: Error, observed: usize, maximum: usize) {
        match error {
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: actual,
                maximum: configured,
            } => {
                assert_eq!(actual, observed);
                assert_eq!(configured, maximum);
            },
            other => panic!("expected XML depth limit, got {other:?}"),
        }
    }
}
