use super::super::super::model::{Comment, List, NamespaceDeclaration, Reply};
use super::super::super::{MAX_BYTES, P188};
use super::parser::parse_comment_list;
use super::validation::{limit, validate_model};
use crate::Result;

fn write_comment(out: &mut Vec<u8>, prefix: &str, comment: &Comment) {
    open_tag(out, prefix, "cm");
    write_attr(out, "id", &comment.id);
    write_attr(out, "authorId", &comment.author_id);
    if let Some(status) = comment.status {
        write_attr(out, "status", status.token());
    }
    write_attr(out, "created", &comment.created);
    if let Some(value) = &comment.start_date {
        write_attr(out, "startDate", value);
    }
    if let Some(value) = &comment.due_date {
        write_attr(out, "dueDate", value);
    }
    if let Some(values) = &comment.assigned_to {
        write_attr(out, "assignedTo", &values.join(" "));
    }
    if let Some(value) = &comment.complete {
        write_u32_attr(out, "complete", value.thousandths());
    }
    if let Some(value) = &comment.title {
        write_attr(out, "title", value);
    }
    write_namespaces(out, &comment.namespace_declarations);
    let has_children = !comment.anchors.is_empty()
        || comment.position.is_some()
        || comment.reply_list_present
        || comment.text_body_xml.is_some()
        || comment.extension_xml.is_some();
    if !has_children {
        out.extend_from_slice(b"/>");
        return;
    }
    out.push(b'>');
    for anchor in &comment.anchors {
        out.extend_from_slice(&anchor.xml);
    }
    if let Some(position) = comment.position {
        open_tag(out, prefix, "pos");
        write_attr(out, "x", &position.x.to_string());
        write_attr(out, "y", &position.y.to_string());
        out.extend_from_slice(b"/>");
    }
    if comment.reply_list_present {
        open_tag(out, prefix, "replyLst");
        write_namespaces(out, &comment.reply_list_namespace_declarations);
        if comment.replies.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for reply in &comment.replies {
                write_reply(out, prefix, reply);
            }
            close_tag(out, prefix, "replyLst");
        }
    }
    if let Some(xml) = &comment.text_body_xml {
        out.extend_from_slice(xml);
    }
    if let Some(xml) = &comment.extension_xml {
        out.extend_from_slice(xml);
    }
    close_tag(out, prefix, "cm");
}

fn write_reply(out: &mut Vec<u8>, prefix: &str, reply: &Reply) {
    open_tag(out, prefix, "reply");
    write_attr(out, "id", &reply.id);
    write_attr(out, "authorId", &reply.author_id);
    if let Some(status) = reply.status {
        write_attr(out, "status", status.token());
    }
    write_attr(out, "created", &reply.created);
    write_namespaces(out, &reply.namespace_declarations);
    if reply.text_body_xml.is_none() && reply.extension_xml.is_none() {
        out.extend_from_slice(b"/>");
    } else {
        out.push(b'>');
        if let Some(xml) = &reply.text_body_xml {
            out.extend_from_slice(xml);
        }
        if let Some(xml) = &reply.extension_xml {
            out.extend_from_slice(xml);
        }
        close_tag(out, prefix, "reply");
    }
}

fn open_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.push(b'<');
    qname(out, prefix, local);
}

fn close_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.extend_from_slice(b"</");
    qname(out, prefix, local);
    out.push(b'>');
}

fn qname(out: &mut Vec<u8>, prefix: &str, local: &str) {
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
}

fn write_namespace_binding(out: &mut Vec<u8>, prefix: &str, uri: &str) {
    out.extend_from_slice(b" xmlns");
    if !prefix.is_empty() {
        out.push(b':');
        out.extend_from_slice(prefix.as_bytes());
    }
    out.extend_from_slice(b"=\"");
    escape(out, uri);
    out.push(b'"');
}

fn write_namespaces(out: &mut Vec<u8>, declarations: &[NamespaceDeclaration]) {
    for declaration in declarations {
        write_namespace_binding(out, &declaration.prefix, &declaration.uri);
    }
}

fn write_attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'"');
}

fn write_u32_attr(out: &mut Vec<u8>, name: &str, value: u32) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    let mut digits = [0; 10];
    let mut cursor = digits.len();
    let mut remaining = value;
    loop {
        cursor -= 1;
        digits[cursor] = b'0' + (remaining % 10) as u8;
        remaining /= 10;
        if remaining == 0 {
            break;
        }
    }
    out.extend_from_slice(&digits[cursor..]);
    out.push(b'"');
}

fn escape(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

impl List {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        parse_comment_list(xml)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_model(self)?;
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        open_tag(&mut out, &self.root_prefix, "cmLst");
        write_namespace_binding(&mut out, &self.root_prefix, P188);
        write_namespaces(&mut out, &self.namespace_declarations);
        if self.comments.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for comment in &self.comments {
                write_comment(&mut out, &self.root_prefix, comment);
            }
            close_tag(&mut out, &self.root_prefix, "cmLst");
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized modern Comment bytes"));
        }
        parse_comment_list(&out)?;
        Ok(out)
    }
}
