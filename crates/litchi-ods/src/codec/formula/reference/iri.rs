//! Allocation-free lexical validation for RFC 3987 IRI references.
//!
//! This module deliberately stops at the generic IRI-reference grammar. It
//! does not resolve a reference, normalize it, apply scheme-specific rules, or
//! access any external resource. The caller supplies an already-decoded IRI;
//! format-specific escaping such as doubled ODF apostrophes is handled before
//! this validator is called.

use std::net::Ipv6Addr;

/// Validate an IRI reference using RFC 3987 section 2.2's generic grammar.
///
/// The implementation only borrows `value`: all scans use indices into the
/// input, a fixed-size scalar state, or the standard library's lexical IPv6
/// parser. In particular, it never allocates while checking a caller-owned
/// reference.
pub(super) fn is_valid_iri_reference(value: &str) -> bool {
    let (without_fragment, fragment) = split_first(value, b'#');
    if let Some(fragment) = fragment {
        if !validate_component(fragment, Component::Fragment) {
            return false;
        }
    }

    let (hier_part, query) = split_first(without_fragment, b'?');
    if let Some(query) = query {
        if !validate_component(query, Component::Query) {
            return false;
        }
    }

    validate_reference_hier_part(hier_part)
}

#[derive(Clone, Copy)]
enum Component {
    UserInfo,
    RegName,
    Path,
    Query,
    Fragment,
    SegmentNoColon,
    IpvFutureTail,
}

fn split_first(value: &str, delimiter: u8) -> (&str, Option<&str>) {
    if let Some(index) = find_byte(value, delimiter) {
        (&value[..index], Some(&value[index + 1..]))
    } else {
        (value, None)
    }
}

fn find_byte(value: &str, needle: u8) -> Option<usize> {
    value.as_bytes().iter().position(|byte| *byte == needle)
}

fn validate_reference_hier_part(value: &str) -> bool {
    if let Some(colon) = find_byte(value, b':') {
        let slash = find_byte(value, b'/');
        if slash.is_none_or(|slash| colon < slash) {
            if !is_valid_scheme(&value[..colon]) {
                // A colon in the first relative path segment is forbidden by
                // `ipath-noscheme`; it cannot be reinterpreted as a relative
                // reference when its prospective scheme is invalid.
                return false;
            }
            return validate_hier_part(&value[colon + 1..]);
        }
    }

    validate_relative_part(value)
}

fn is_valid_scheme(value: &str) -> bool {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return false;
    };
    if !first.is_ascii_alphabetic() {
        return false;
    }

    bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
}

fn validate_hier_part(value: &str) -> bool {
    if let Some(network_path) = value.strip_prefix("//") {
        return validate_network_path(network_path);
    }
    if value.starts_with('/') {
        // The authority alternative was already selected above for `//`.
        return validate_component(value, Component::Path);
    }
    if value.is_empty() {
        return true;
    }

    // `ipath-rootless = isegment-nz *( "/" isegment )`.
    validate_component(value, Component::Path)
}

fn validate_relative_part(value: &str) -> bool {
    if let Some(network_path) = value.strip_prefix("//") {
        return validate_network_path(network_path);
    }
    if value.starts_with('/') {
        // `ipath-absolute` begins with one slash. A leading `//` was handled
        // as the authority form above.
        return validate_component(value, Component::Path);
    }
    if value.is_empty() {
        return true;
    }

    // `ipath-noscheme = isegment-nz-nc *( "/" isegment )`.
    let first_slash = find_byte(value, b'/').unwrap_or(value.len());
    if first_slash == 0 || !validate_component(&value[..first_slash], Component::SegmentNoColon) {
        return false;
    }
    validate_component(&value[first_slash..], Component::Path)
}

fn validate_network_path(value: &str) -> bool {
    let slash = find_byte(value, b'/').unwrap_or(value.len());
    let authority = &value[..slash];
    let path = &value[slash..];
    validate_authority(authority) && validate_component(path, Component::Path)
}

fn validate_authority(value: &str) -> bool {
    let (userinfo, host_port) = if let Some(at) = find_byte(value, b'@') {
        (Some(&value[..at]), &value[at + 1..])
    } else {
        (None, value)
    };

    if let Some(userinfo) = userinfo {
        if !validate_component(userinfo, Component::UserInfo) {
            return false;
        }
    }

    if host_port.starts_with('[') {
        let Some(close) = find_byte(host_port, b']') else {
            return false;
        };
        if !is_valid_ip_literal(&host_port[1..close]) {
            return false;
        }
        return validate_port_suffix(&host_port[close + 1..]);
    }

    let colon = find_byte(host_port, b':').unwrap_or(host_port.len());
    let host = &host_port[..colon];
    if !validate_host(host) {
        return false;
    }
    if colon == host_port.len() {
        return true;
    }

    validate_port(&host_port[colon + 1..])
}

fn validate_port_suffix(value: &str) -> bool {
    if value.is_empty() {
        return true;
    }
    let Some(port) = value.strip_prefix(':') else {
        return false;
    };
    validate_port(port)
}

fn validate_port(value: &str) -> bool {
    // RFC 3986/3987 define `port = *DIGIT`, so an empty port after the colon
    // is syntactically valid and scheme policy is outside this validator.
    value.bytes().all(|byte| byte.is_ascii_digit())
}

fn validate_host(value: &str) -> bool {
    // The host alternatives are ordered IPv4address / ireg-name. If the
    // dotted-decimal production does not match, the generic reg-name remains
    // a legal fallback under RFC 3986's first-match-wins rule.
    is_valid_ipv4(value) || validate_component(value, Component::RegName)
}

fn is_valid_ip_literal(value: &str) -> bool {
    if value.starts_with('v') || value.starts_with('V') {
        let Some(dot) = find_byte(value, b'.') else {
            return false;
        };
        if dot <= 1
            || !value.as_bytes()[1..dot]
                .iter()
                .all(|byte| byte.is_ascii_hexdigit())
        {
            return false;
        }
        let tail = &value[dot + 1..];
        return !tail.is_empty() && validate_component(tail, Component::IpvFutureTail);
    }

    is_valid_ipv6(value)
}

fn is_valid_ipv6(value: &str) -> bool {
    if value.is_empty() {
        return false;
    }

    if !value
        .bytes()
        .all(|byte| byte.is_ascii_hexdigit() || matches!(byte, b':' | b'.'))
    {
        return false;
    }

    // `Ipv6Addr` performs the RFC 3986 IPv6 production check. Validate an
    // embedded dotted-decimal ls32 ourselves as well, because the ABNF
    // forbids leading-zero octets even if a platform parser accepts them.
    if let Some(dot) = find_byte(value, b'.') {
        let Some(colon) = value.as_bytes().iter().rposition(|byte| *byte == b':') else {
            return false;
        };
        if dot < colon || !is_valid_ipv4(&value[colon + 1..]) {
            return false;
        }
    }

    value.parse::<Ipv6Addr>().is_ok()
}

fn is_valid_ipv4(value: &str) -> bool {
    let mut octet_count = 0_u8;
    let mut digit_count = 0_u8;
    let mut number = 0_u16;
    let mut leading_zero = false;

    for byte in value.bytes() {
        if byte.is_ascii_digit() {
            if digit_count == 0 {
                leading_zero = byte == b'0';
            } else if leading_zero {
                return false;
            }
            digit_count += 1;
            if digit_count > 3 {
                return false;
            }
            number = number * 10 + u16::from(byte - b'0');
            if number > 255 {
                return false;
            }
        } else if byte == b'.' {
            if digit_count == 0 || octet_count == 3 {
                return false;
            }
            octet_count += 1;
            digit_count = 0;
            number = 0;
            leading_zero = false;
        } else {
            return false;
        }
    }

    digit_count != 0 && octet_count == 3
}

fn validate_component(value: &str, component: Component) -> bool {
    let bytes = value.as_bytes();
    let mut index = 0;

    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            if matches!(component, Component::IpvFutureTail)
                || index + 2 >= bytes.len()
                || !bytes[index + 1].is_ascii_hexdigit()
                || !bytes[index + 2].is_ascii_hexdigit()
            {
                return false;
            }
            index += 3;
            continue;
        }

        if byte.is_ascii() {
            if !is_allowed_ascii(byte, component) {
                return false;
            }
            index += 1;
            continue;
        }

        let Some(character) = value[index..].chars().next() else {
            return false;
        };
        if !is_allowed_unicode(character, component) {
            return false;
        }
        index += character.len_utf8();
    }

    true
}

fn is_allowed_ascii(byte: u8, component: Component) -> bool {
    let iunreserved = byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~');
    let sub_delim = matches!(
        byte,
        b'!' | b'$' | b'&' | b'\'' | b'(' | b')' | b'*' | b'+' | b',' | b';' | b'='
    );

    match component {
        Component::UserInfo => iunreserved || sub_delim || byte == b':',
        Component::RegName => iunreserved || sub_delim,
        Component::Path => iunreserved || sub_delim || matches!(byte, b':' | b'@' | b'/'),
        Component::Query | Component::Fragment => {
            iunreserved || sub_delim || matches!(byte, b':' | b'@' | b'/' | b'?')
        },
        Component::SegmentNoColon => iunreserved || sub_delim || byte == b'@',
        Component::IpvFutureTail => iunreserved || sub_delim || byte == b':',
    }
}

fn is_allowed_unicode(character: char, component: Component) -> bool {
    match component {
        Component::Query => is_ucschar(character) || is_iprivate(character),
        Component::IpvFutureTail => false,
        Component::UserInfo
        | Component::RegName
        | Component::Path
        | Component::Fragment
        | Component::SegmentNoColon => is_ucschar(character),
    }
}

fn is_ucschar(character: char) -> bool {
    let value = character as u32;
    (0x00A0..=0xD7FF).contains(&value)
        || (0xF900..=0xFDCF).contains(&value)
        || (0xFDF0..=0xFFEF).contains(&value)
        || (0x10000..=0x1FFFD).contains(&value)
        || (0x20000..=0x2FFFD).contains(&value)
        || (0x30000..=0x3FFFD).contains(&value)
        || (0x40000..=0x4FFFD).contains(&value)
        || (0x50000..=0x5FFFD).contains(&value)
        || (0x60000..=0x6FFFD).contains(&value)
        || (0x70000..=0x7FFFD).contains(&value)
        || (0x80000..=0x8FFFD).contains(&value)
        || (0x90000..=0x9FFFD).contains(&value)
        || (0xA0000..=0xAFFFD).contains(&value)
        || (0xB0000..=0xBFFFD).contains(&value)
        || (0xC0000..=0xCFFFD).contains(&value)
        || (0xD0000..=0xDFFFD).contains(&value)
        || (0xE1000..=0xEFFFD).contains(&value)
}

fn is_iprivate(character: char) -> bool {
    let value = character as u32;
    (0xE000..=0xF8FF).contains(&value)
        || (0xF0000..=0xFFFFD).contains(&value)
        || (0x100000..=0x10FFFD).contains(&value)
}

#[cfg(test)]
mod tests {
    use super::is_valid_iri_reference;

    #[test]
    fn accepts_empty_relative_query_and_fragment_references() {
        for value in [
            "",
            "?",
            "#",
            "?#",
            "?query=yes",
            "#fragment",
            "path/to/resource",
            "./relative",
            "../parent",
            "/absolute/path",
            "//example.org/path",
            "a/b:c",
            "/a:b",
        ] {
            assert!(
                is_valid_iri_reference(value),
                "expected valid IRI reference: {value:?}"
            );
        }
    }

    #[test]
    fn accepts_absolute_schemes_and_authority_forms() {
        for value in [
            "http:",
            "http:path",
            "http:/path",
            "http://",
            "http:///path",
            "http://user:password@example.org:443/path",
            "http://example.org:",
            "http://:8080/path",
            "urn:example:animal:ferret:nose",
            "mailto:person@example.org?subject=hello",
            "https://例え.テスト/こんにちは",
        ] {
            assert!(
                is_valid_iri_reference(value),
                "expected valid IRI reference: {value:?}"
            );
        }
    }

    #[test]
    fn accepts_ipv6_and_ipvfuture_literals() {
        for value in [
            "http://[::1]/",
            "http://[2001:db8::1]:443/",
            "http://[::ffff:192.0.2.1]/",
            "http://[v1.fe]/",
            "http://[V1.fe:tail]/",
        ] {
            assert!(
                is_valid_iri_reference(value),
                "expected valid IP literal: {value:?}"
            );
        }
    }

    #[test]
    fn accepts_ucschar_and_private_query_characters_only_in_their_grammar_slots() {
        for value in [
            "https://example.org/\u{00A0}\u{F900}",
            "https://example.org/?\u{E000}\u{F0000}\u{100000}",
            "?\u{10FFFD}",
        ] {
            assert!(
                is_valid_iri_reference(value),
                "expected valid Unicode IRI: {value:?}"
            );
        }

        for value in [
            "https://example.org/\u{E000}",
            "https://example.org/#\u{E000}",
            "https://\u{E000}/",
            "https://example.org/?\u{10FFFE}",
        ] {
            assert!(
                !is_valid_iri_reference(value),
                "expected invalid Unicode IRI: {value:?}"
            );
        }
    }

    #[test]
    fn rejects_malformed_ip_literals_and_ports() {
        for value in [
            "http://[::1",
            "http://[::1]]/",
            "http://[::1%25eth0]/",
            "http://[2001:db8:0:0:0:0:0:0:1]/",
            "http://[::ffff:192.00.2.1]/",
            "http://[::ffff:999.0.2.1]/",
            "http://[v]/",
            "http://[v1]/",
            "http://[v1.]/",
            "http://[v1.^]/",
            "http://[v1.fe%20]/",
            "http://example.org:abc/",
            "http://example.org:80:90/",
        ] {
            assert!(
                !is_valid_iri_reference(value),
                "expected invalid IP/port IRI: {value:?}"
            );
        }
    }

    #[test]
    fn rejects_bad_percent_escapes_and_disallowed_delimiters() {
        for value in [
            "http://example.org/%",
            "http://example.org/%G0",
            "http://example.org/%0",
            "http://example.org/a\\b",
            "http://example.org/a[b",
            "http://example.org/a b",
            "http://example.org/#frag#again",
            "http://example.org/?query#fragment#again",
        ] {
            assert!(
                !is_valid_iri_reference(value),
                "expected invalid syntax: {value:?}"
            );
        }
    }

    #[test]
    fn rejects_invalid_first_segment_colons() {
        for value in [":/path", "1scheme:value", "bad@name:value", "a%3Ab:c"] {
            assert!(
                !is_valid_iri_reference(value),
                "expected invalid relative prefix: {value:?}"
            );
        }
        assert!(is_valid_iri_reference("a:b"));
        assert!(is_valid_iri_reference("a/b:c"));
    }
}
