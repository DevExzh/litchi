use super::super::semantic::{OpaqueXml, Reaction, ReactionInstance, Reactions};
use super::super::{P, P223};
use super::xml::{
    attr, attribute, close, no_attributes, only_attributes, open, scan, scan_with_context,
};
use crate::{Error, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn parse_reactions(xml: &[u8]) -> Result<Reactions> {
    let scan = scan(xml, "reactions")?;
    if scan.root.namespace != P223 || scan.root.local != "reactions" {
        return Err(invalid("reactions root must be p223:reactions"));
    }
    no_attributes(&scan.root.attributes, "reactions")?;
    let mut reactions = Vec::with_capacity(scan.children.len());
    for child in &scan.children {
        if child.namespace != P223 || child.local != "rxn" {
            return Err(invalid("reactions permits only p223:rxn children"));
        }
        reactions.push(parse_reaction(child, &scan.namespaces)?);
    }
    let value = Reactions {
        reactions,
        namespace_declarations: scan.namespaces,
    };
    value.validate()?;
    Ok(value)
}

fn parse_reaction(
    fragment: &super::xml::Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<Reaction> {
    only_attributes(&fragment.attributes, &["type"], "reaction")?;
    let reaction_type = attribute(&fragment.attributes, "type", true)?
        .unwrap()
        .to_owned();
    let scan = scan_with_context(&fragment.xml, "reaction", context)?;
    let mut instances = Vec::with_capacity(scan.children.len());
    for child in &scan.children {
        if child.namespace != P223 || child.local != "instance" {
            return Err(invalid("reaction permits only p223:instance children"));
        }
        instances.push(parse_instance(child, &scan.namespaces)?);
    }
    Ok(Reaction {
        reaction_type,
        instances,
        namespace_declarations: context.to_vec(),
    })
}

fn parse_instance(
    fragment: &super::xml::Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<ReactionInstance> {
    only_attributes(
        &fragment.attributes,
        &["time", "authorId"],
        "reaction instance",
    )?;
    let time = attribute(&fragment.attributes, "time", true)?
        .unwrap()
        .to_owned();
    let author_id = attribute(&fragment.attributes, "authorId", true)?
        .unwrap()
        .to_owned();
    let scan = scan_with_context(&fragment.xml, "reaction instance", context)?;
    let mut extension = None;
    for child in &scan.children {
        if child.namespace == P && child.local == "extLst" && extension.is_none() {
            extension = Some(OpaqueXml::new(child.xml.clone())?);
        } else {
            return Err(invalid("unexpected reaction instance child"));
        }
    }
    Ok(ReactionInstance {
        time,
        author_id,
        extension_xml: extension,
        namespace_declarations: context.to_vec(),
    })
}

pub(super) fn write_reactions(value: &Reactions) -> Result<Vec<u8>> {
    value.validate()?;
    let mut out = Vec::new();
    open(&mut out, "p223", "reactions");
    out.extend_from_slice(
        b" xmlns:p223=\"http://schemas.microsoft.com/office/powerpoint/2022/03/main\"",
    );
    out.extend_from_slice(
        b" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"",
    );
    for declaration in &value.namespace_declarations {
        if declaration.uri != P223
            && declaration.uri != P
            && declaration.prefix != "p223"
            && declaration.prefix != "p"
        {
            super::xml::namespaces(&mut out, std::slice::from_ref(declaration));
        }
    }
    if value.reactions.is_empty() {
        out.extend_from_slice(b"/>");
        return Ok(out);
    }
    out.push(b'>');
    for reaction in &value.reactions {
        open(&mut out, "p223", "rxn");
        attr(&mut out, "type", &reaction.reaction_type);
        if reaction.instances.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for instance in &reaction.instances {
                open(&mut out, "p223", "instance");
                attr(&mut out, "time", &instance.time);
                attr(&mut out, "authorId", &instance.author_id);
                if let Some(extension) = &instance.extension_xml {
                    out.push(b'>');
                    out.extend_from_slice(extension.as_bytes());
                    close(&mut out, "p223", "instance");
                } else {
                    out.extend_from_slice(b"/>");
                }
            }
            close(&mut out, "p223", "rxn");
        }
    }
    close(&mut out, "p223", "reactions");
    Ok(out)
}
