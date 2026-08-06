//! Bounded SpreadsheetML connection-part XML codec.

use super::invalid;
use super::model::*;
use litchi_core::sheet::Result;
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};

#[derive(Clone)]
struct Attr {
    q: String,
    ns: String,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
pub(super) struct Node {
    q: String,
    ns: String,
    l: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}

impl Connections {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid("connections part exceeds 16 MiB"));
        }
        let x = litchi_ooxml_common::mce::process_ooxml(xml)?;
        if x.len() > MAX_XML_BYTES {
            return Err(invalid("processed connections part exceeds 16 MiB"));
        }
        project(&parse_dom(x.as_ref())?)
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let ns = if strict {
            STRICT_NAMESPACE
        } else {
            CORE_NAMESPACE
        };
        let mut x = BoundedXml::new();
        x.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><connections xmlns=\"",
        )?;
        x.push_str(ns)?;
        x.push_str("\">")?;
        for c in &self.connections {
            write_connection(&mut x, c, strict)?;
        }
        x.push_str("</connections>")?;
        Ok(x.finish())
    }
}

pub(super) fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    loop {
        let d = rd.decoder();
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count += 1;
                if count > MAX_DOM_NODES || stack.len() >= MAX_DOM_DEPTH {
                    return Err(invalid("connections XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack)?);
            },
            Ok(Event::Empty(e)) => {
                count += 1;
                if count > MAX_DOM_NODES {
                    return Err(invalid("connections node limit exceeded"));
                }
                let n = make(&e, d, &stack)?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v))
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside connections"));
                }
            },
            Ok(Event::CData(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::CData(t.decode().map_err(xml_error)?.into_owned()))
                } else {
                    return Err(invalid("CDATA outside connections"));
                }
            },
            Ok(Event::Comment(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(
                        t.decode().map_err(xml_error)?.into_owned(),
                    ))
                }
            },
            Ok(Event::GeneralRef(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(
                        litchi_ooxml_common::xml::decode_xml_reference(&t)?,
                    ))
                } else {
                    return Err(invalid("entity outside connections"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated connections XML"));
    }
    root.ok_or_else(|| invalid("missing connections root"))
}
fn make(e: &BytesStart<'_>, d: Decoder, stack: &[Node]) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|x| x.bindings.clone()).unwrap_or_default();
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        raw.push((
            std::str::from_utf8(a.key.as_ref())
                .map_err(xml_error)?
                .to_string(),
            a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                .map_err(xml_error)?
                .into_owned(),
        ));
    }
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(x) = bindings.iter_mut().find(|x| x.0 == key) {
                x.1 = v.clone()
            } else {
                bindings.push((key, v.clone()))
            }
        }
    }
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    let ns = resolve(&bindings, pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            String::new()
        } else {
            resolve(&bindings, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content.push(Content::Node(n))
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(n: &Node) -> Result<Connections> {
    expect(n, "connections")?;
    noattrs(n)?;
    let mut out = Vec::new();
    for c in kids(n)? {
        // Future producers may place extension elements beside the typed
        // connection catalog. They stay in the source snapshot and are not
        // projected into the semantic model.
        if c.ns != CORE_NAMESPACE && c.ns != STRICT_NAMESPACE || c.l != "connection" {
            continue;
        }
        if out.len() >= MAX_CONNECTIONS {
            return Err(invalid("connection limit exceeded"));
        }
        out.push(parse_connection(c)?);
    }
    if out.is_empty() {
        return Err(invalid("connections requires at least one connection"));
    }
    let value = Connections { connections: out };
    validate(&value)?;
    Ok(value)
}
fn parse_connection(n: &Node) -> Result<Connection> {
    expect(n, "connection")?;
    let mut c = Connection {
        id: u32req(n, "id")?,
        source_file: aopt(n, "sourceFile")?,
        odc_file: aopt(n, "odcFile")?,
        keep_alive: bopt(n, "keepAlive")?,
        interval: u32opt(n, "interval")?,
        name: aopt(n, "name")?,
        description: aopt(n, "description")?,
        connection_type: u32opt(n, "type")?,
        reconnection_method: u32opt(n, "reconnectionMethod")?,
        refreshed_version: u8req(n, "refreshedVersion")?,
        min_refreshable_version: u8opt(n, "minRefreshableVersion")?,
        save_password: bopt(n, "savePassword")?,
        new_connection: bopt(n, "new")?,
        deleted: bopt(n, "deleted")?,
        only_use_connection_file: bopt(n, "onlyUseConnectionFile")?,
        background: bopt(n, "background")?,
        refresh_on_load: bopt(n, "refreshOnLoad")?,
        save_data: bopt(n, "saveData")?,
        credentials: aopt(n, "credentials")?.map(parse_credentials).transpose()?,
        single_sign_on_id: aopt(n, "singleSignOnId")?,
        database: None,
        olap: None,
        web: None,
        text: None,
        parameters: None,
        extension_xml: None,
    };
    only(
        n,
        &[
            "id",
            "sourceFile",
            "odcFile",
            "keepAlive",
            "interval",
            "name",
            "description",
            "type",
            "reconnectionMethod",
            "refreshedVersion",
            "minRefreshableVersion",
            "savePassword",
            "new",
            "deleted",
            "onlyUseConnectionFile",
            "background",
            "refreshOnLoad",
            "saveData",
            "credentials",
            "singleSignOnId",
        ],
    )?;
    let mut order = 0;
    for child in kids(n)? {
        if child.ns != CORE_NAMESPACE && child.ns != STRICT_NAMESPACE {
            continue;
        }
        let i = match child.l.as_str() {
            "dbPr" => 0,
            "olapPr" => 1,
            "webPr" => 2,
            "textPr" => 3,
            "parameters" => 4,
            "extLst" => 5,
            _ => continue,
        };
        if i < order {
            return Err(invalid("connection children out of order"));
        }
        order = i;
        match i {
            0 => set(&mut c.database, parse_db(child)?)?,
            1 => set(&mut c.olap, parse_olap(child)?)?,
            2 => set(&mut c.web, parse_web(child)?)?,
            3 => set(&mut c.text, parse_text(child)?)?,
            4 => set(&mut c.parameters, parse_parameters(child)?)?,
            5 => set(&mut c.extension_xml, node_xml(child, false)?)?,
            _ => return Err(invalid("unexpected connection child index")),
        }
    }
    Ok(c)
}
fn parse_db(n: &Node) -> Result<DatabaseProperties> {
    let v = DatabaseProperties {
        connection: req(n, "connection")?,
        command: aopt(n, "command")?,
        server_command: aopt(n, "serverCommand")?,
        command_type: u32opt(n, "commandType")?,
    };
    only(
        n,
        &["connection", "command", "serverCommand", "commandType"],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_olap(n: &Node) -> Result<OlapProperties> {
    let v = OlapProperties {
        local: bopt(n, "local")?,
        local_connection: aopt(n, "localConnection")?,
        local_refresh: bopt(n, "localRefresh")?,
        send_locale: bopt(n, "sendLocale")?,
        row_drill_count: u32opt(n, "rowDrillCount")?,
        server_fill: bopt(n, "serverFill")?,
        server_number_format: bopt(n, "serverNumberFormat")?,
        server_font: bopt(n, "serverFont")?,
        server_font_color: bopt(n, "serverFontColor")?,
    };
    only(
        n,
        &[
            "local",
            "localConnection",
            "localRefresh",
            "sendLocale",
            "rowDrillCount",
            "serverFill",
            "serverNumberFormat",
            "serverFont",
            "serverFontColor",
        ],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_web(n: &Node) -> Result<WebQueryProperties> {
    let mut v = WebQueryProperties {
        xml_source: bopt(n, "xml")?,
        source_data: bopt(n, "sourceData")?,
        parse_pre: bopt(n, "parsePre")?,
        consecutive: bopt(n, "consecutive")?,
        first_row: bopt(n, "firstRow")?,
        excel97: bopt(n, "xl97")?,
        text_dates: bopt(n, "textDates")?,
        excel2000: bopt(n, "xl2000")?,
        url: aopt(n, "url")?,
        post: aopt(n, "post")?,
        html_tables: bopt(n, "htmlTables")?,
        html_format: aopt(n, "htmlFormat")?.map(parse_html).transpose()?,
        edit_page: aopt(n, "editPage")?,
        tables: None,
    };
    only(
        n,
        &[
            "xml",
            "sourceData",
            "parsePre",
            "consecutive",
            "firstRow",
            "xl97",
            "textDates",
            "xl2000",
            "url",
            "post",
            "htmlTables",
            "htmlFormat",
            "editPage",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("webPr permits one tables child"));
    }
    if let Some(t) = c.first() {
        expect(t, "tables")?;
        v.tables = Some(parse_tables(t)?);
    }
    Ok(v)
}
fn parse_tables(n: &Node) -> Result<Vec<WebTableSelector>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for c in kids(n)? {
        if out.len() >= MAX_WEB_TABLES {
            return Err(invalid("web table selector limit exceeded"));
        }
        expect_any(c)?;
        out.push(match c.l.as_str() {
            "m" => {
                noattrs(c)?;
                leaf(c)?;
                WebTableSelector::Missing
            },
            "s" => {
                let v = req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::String(v)
            },
            "x" => {
                let v = u32req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::Index(v)
            },
            _ => return Err(invalid("invalid web table selector")),
        });
    }
    if out.is_empty() {
        return Err(invalid("tables requires a selector"));
    }
    check_count(count, out.len(), "tables")?;
    Ok(out)
}
fn parse_text(n: &Node) -> Result<TextImportProperties> {
    let mut v = TextImportProperties {
        prompt: bopt(n, "prompt")?,
        file_type: aopt(n, "fileType")?.map(parse_file).transpose()?,
        code_page: u32opt(n, "codePage")?,
        character_set: aopt(n, "characterSet")?,
        first_row: u32opt(n, "firstRow")?,
        source_file: aopt(n, "sourceFile")?,
        delimited: bopt(n, "delimited")?,
        decimal: aopt(n, "decimal")?,
        thousands: aopt(n, "thousands")?,
        tab: bopt(n, "tab")?,
        space: bopt(n, "space")?,
        comma: bopt(n, "comma")?,
        semicolon: bopt(n, "semicolon")?,
        consecutive: bopt(n, "consecutive")?,
        qualifier: aopt(n, "qualifier")?.map(parse_qualifier).transpose()?,
        delimiter: aopt(n, "delimiter")?,
        fields: None,
    };
    only(
        n,
        &[
            "prompt",
            "fileType",
            "codePage",
            "characterSet",
            "firstRow",
            "sourceFile",
            "delimited",
            "decimal",
            "thousands",
            "tab",
            "space",
            "comma",
            "semicolon",
            "consecutive",
            "qualifier",
            "delimiter",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("textPr permits one textFields child"));
    }
    if let Some(f) = c.first() {
        expect(f, "textFields")?;
        let count = u32opt(f, "count")?;
        only(f, &["count"])?;
        let mut fields = Vec::new();
        for e in kids(f)? {
            if fields.len() >= MAX_TEXT_FIELDS {
                return Err(invalid("text field limit exceeded"));
            }
            expect(e, "textField")?;
            fields.push(TextField {
                field_type: aopt(e, "type")?.map(parse_field).transpose()?,
                position: u32opt(e, "position")?,
            });
            only(e, &["type", "position"])?;
            leaf(e)?;
        }
        if fields.is_empty() {
            return Err(invalid("textFields requires a textField"));
        }
        check_count(count, fields.len(), "textFields")?;
        v.fields = Some(fields);
    }
    Ok(v)
}
fn parse_parameters(n: &Node) -> Result<Vec<ConnectionParameter>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for p in kids(n)? {
        if out.len() >= MAX_PARAMETERS {
            return Err(invalid("parameter limit exceeded"));
        }
        expect(p, "parameter")?;
        let double = match aopt(p, "double")? {
            Some(x) => {
                let v = x
                    .parse::<f64>()
                    .map_err(|_| invalid("invalid parameter double"))?;
                if !v.is_finite() {
                    return Err(invalid("non-finite parameter double"));
                }
                Some(v)
            },
            None => None,
        };
        out.push(ConnectionParameter {
            name: aopt(p, "name")?,
            sql_type: i32opt(p, "sqlType")?,
            parameter_type: aopt(p, "parameterType")?
                .map(parse_parameter_type)
                .transpose()?,
            refresh_on_change: bopt(p, "refreshOnChange")?,
            prompt: aopt(p, "prompt")?,
            boolean: bopt(p, "boolean")?,
            double,
            integer: i32opt(p, "integer")?,
            string: aopt(p, "string")?,
            cell: aopt(p, "cell")?,
        });
        only(
            p,
            &[
                "name",
                "sqlType",
                "parameterType",
                "refreshOnChange",
                "prompt",
                "boolean",
                "double",
                "integer",
                "string",
                "cell",
            ],
        )?;
        leaf(p)?;
    }
    if out.is_empty() {
        return Err(invalid("parameters requires a parameter"));
    }
    check_count(count, out.len(), "parameters")?;
    Ok(out)
}

pub(super) struct BoundedXml {
    pub(super) bytes: Vec<u8>,
}

impl BoundedXml {
    pub(super) fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.push_bytes(value.as_bytes())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        let mut encoded = [0; 4];
        let length = value.encode_utf8(&mut encoded).len();
        self.push_bytes(&encoded[..length])
    }

    pub(super) fn push_bytes(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("serialized connections length overflows"))?;
        if length > MAX_XML_BYTES {
            return Err(invalid("serialized connections part exceeds 16 MiB"));
        }
        self.bytes
            .try_reserve_exact(value.len())
            .map_err(|_| invalid("serialized connections output allocation failed"))?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn write_connection(x: &mut BoundedXml, c: &Connection, s: bool) -> Result<()> {
    x.push_str("<connection")?;
    num(x, "id", c.id)?;
    str_opt(x, "sourceFile", c.source_file.as_deref())?;
    str_opt(x, "odcFile", c.odc_file.as_deref())?;
    bool_opt(x, "keepAlive", c.keep_alive)?;
    num_opt(x, "interval", c.interval)?;
    str_opt(x, "name", c.name.as_deref())?;
    str_opt(x, "description", c.description.as_deref())?;
    num_opt(x, "type", c.connection_type)?;
    num_opt(x, "reconnectionMethod", c.reconnection_method)?;
    num(x, "refreshedVersion", c.refreshed_version)?;
    num_opt(x, "minRefreshableVersion", c.min_refreshable_version)?;
    bool_opt(x, "savePassword", c.save_password)?;
    bool_opt(x, "new", c.new_connection)?;
    bool_opt(x, "deleted", c.deleted)?;
    bool_opt(x, "onlyUseConnectionFile", c.only_use_connection_file)?;
    bool_opt(x, "background", c.background)?;
    bool_opt(x, "refreshOnLoad", c.refresh_on_load)?;
    bool_opt(x, "saveData", c.save_data)?;
    if let Some(v) = c.credentials {
        attr(x, "credentials", credentials_str(v))?;
    }
    str_opt(x, "singleSignOnId", c.single_sign_on_id.as_deref())?;
    if c.database.is_none()
        && c.olap.is_none()
        && c.web.is_none()
        && c.text.is_none()
        && c.parameters.is_none()
        && c.extension_xml.is_none()
    {
        x.push_str("/>")?;
        return Ok(());
    }
    x.push_char('>')?;
    if let Some(v) = &c.database {
        x.push_str("<dbPr")?;
        attr(x, "connection", &v.connection)?;
        str_opt(x, "command", v.command.as_deref())?;
        str_opt(x, "serverCommand", v.server_command.as_deref())?;
        num_opt(x, "commandType", v.command_type)?;
        x.push_str("/>")?;
    }
    if let Some(v) = &c.olap {
        write_olap(x, v)?;
    }
    if let Some(v) = &c.web {
        write_web(x, v)?;
    }
    if let Some(v) = &c.text {
        write_text(x, v)?;
    }
    if let Some(v) = &c.parameters {
        write_parameters(x, v)?;
    }
    if let Some(v) = &c.extension_xml {
        opaque(x, v, s)?;
    }
    x.push_str("</connection>")?;
    Ok(())
}
fn write_olap(x: &mut BoundedXml, v: &OlapProperties) -> Result<()> {
    x.push_str("<olapPr")?;
    for (n, b) in [
        ("local", v.local),
        ("localRefresh", v.local_refresh),
        ("sendLocale", v.send_locale),
        ("serverFill", v.server_fill),
        ("serverNumberFormat", v.server_number_format),
        ("serverFont", v.server_font),
        ("serverFontColor", v.server_font_color),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "localConnection", v.local_connection.as_deref())?;
    num_opt(x, "rowDrillCount", v.row_drill_count)?;
    x.push_str("/>")?;
    Ok(())
}
fn write_web(x: &mut BoundedXml, v: &WebQueryProperties) -> Result<()> {
    x.push_str("<webPr")?;
    for (n, b) in [
        ("xml", v.xml_source),
        ("sourceData", v.source_data),
        ("parsePre", v.parse_pre),
        ("consecutive", v.consecutive),
        ("firstRow", v.first_row),
        ("xl97", v.excel97),
        ("textDates", v.text_dates),
        ("xl2000", v.excel2000),
        ("htmlTables", v.html_tables),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "url", v.url.as_deref())?;
    str_opt(x, "post", v.post.as_deref())?;
    if let Some(h) = v.html_format {
        attr(x, "htmlFormat", html_str(h))?;
    }
    str_opt(x, "editPage", v.edit_page.as_deref())?;
    if let Some(t) = &v.tables {
        x.push_str("><tables")?;
        num(x, "count", t.len())?;
        x.push_char('>')?;
        for z in t {
            match z {
                WebTableSelector::Missing => x.push_str("<m/>")?,
                WebTableSelector::String(v) => {
                    x.push_str("<s")?;
                    attr(x, "v", v)?;
                    x.push_str("/>")?;
                },
                WebTableSelector::Index(v) => {
                    x.push_str("<x")?;
                    num(x, "v", *v)?;
                    x.push_str("/>")?;
                },
            }
        }
        x.push_str("</tables></webPr>")?;
    } else {
        x.push_str("/>")?;
    }
    Ok(())
}
fn write_text(x: &mut BoundedXml, v: &TextImportProperties) -> Result<()> {
    x.push_str("<textPr")?;
    bool_opt(x, "prompt", v.prompt)?;
    if let Some(z) = v.file_type {
        attr(x, "fileType", file_str(z))?;
    }
    num_opt(x, "codePage", v.code_page)?;
    str_opt(x, "characterSet", v.character_set.as_deref())?;
    num_opt(x, "firstRow", v.first_row)?;
    str_opt(x, "sourceFile", v.source_file.as_deref())?;
    for (n, b) in [
        ("delimited", v.delimited),
        ("tab", v.tab),
        ("space", v.space),
        ("comma", v.comma),
        ("semicolon", v.semicolon),
        ("consecutive", v.consecutive),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "decimal", v.decimal.as_deref())?;
    str_opt(x, "thousands", v.thousands.as_deref())?;
    if let Some(z) = v.qualifier {
        attr(x, "qualifier", qualifier_str(z))?;
    }
    str_opt(x, "delimiter", v.delimiter.as_deref())?;
    if let Some(f) = &v.fields {
        x.push_str("><textFields")?;
        num(x, "count", f.len())?;
        x.push_char('>')?;
        for z in f {
            x.push_str("<textField")?;
            if let Some(t) = z.field_type {
                attr(x, "type", field_str(t))?;
            }
            num_opt(x, "position", z.position)?;
            x.push_str("/>")?;
        }
        x.push_str("</textFields></textPr>")?;
    } else {
        x.push_str("/>")?;
    }
    Ok(())
}
fn write_parameters(x: &mut BoundedXml, v: &[ConnectionParameter]) -> Result<()> {
    x.push_str("<parameters")?;
    num(x, "count", v.len())?;
    x.push_char('>')?;
    for p in v {
        x.push_str("<parameter")?;
        str_opt(x, "name", p.name.as_deref())?;
        num_opt(x, "sqlType", p.sql_type)?;
        if let Some(z) = p.parameter_type {
            attr(x, "parameterType", parameter_str(z))?;
        }
        bool_opt(x, "refreshOnChange", p.refresh_on_change)?;
        str_opt(x, "prompt", p.prompt.as_deref())?;
        bool_opt(x, "boolean", p.boolean)?;
        if let Some(z) = p.double {
            let value = z.to_string();
            attr(x, "double", &value)?;
        }
        num_opt(x, "integer", p.integer)?;
        str_opt(x, "string", p.string.as_deref())?;
        str_opt(x, "cell", p.cell.as_deref())?;
        x.push_str("/>")?;
    }
    x.push_str("</parameters>")?;
    Ok(())
}

fn opaque(x: &mut BoundedXml, b: &[u8], strict: bool) -> Result<()> {
    parse_dom(b)?;
    let mut s = std::str::from_utf8(b).map_err(xml_error)?.to_string();
    if strict {
        s = s.replace(CORE_NAMESPACE, STRICT_NAMESPACE)
    } else {
        s = s.replace(STRICT_NAMESPACE, CORE_NAMESPACE)
    }
    x.push_str(&s)?;
    Ok(())
}
fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    node_write(&mut x, n, s)?;
    Ok(x.into_bytes())
}
fn node_write(x: &mut String, n: &Node, s: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in &n.bindings {
        if p.is_empty() {
            x.push_str(" xmlns=\"")
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"")
        }
        esc(
            x,
            if s && u == CORE_NAMESPACE {
                STRICT_NAMESPACE
            } else if !s && u == STRICT_NAMESPACE {
                CORE_NAMESPACE
            } else {
                u
            },
        );
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n, s)?,
            Content::Text(v) => text_escape(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}

pub(super) fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed connections")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("connection leaf has children"))
    }
}
pub(super) fn expect(n: &Node, l: &str) -> Result<()> {
    if (n.ns == CORE_NAMESPACE || n.ns == STRICT_NAMESPACE) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected SpreadsheetML {l}")))
    }
}
fn expect_any(n: &Node) -> Result<()> {
    if n.ns == CORE_NAMESPACE || n.ns == STRICT_NAMESPACE {
        Ok(())
    } else {
        Err(invalid("expected SpreadsheetML child"))
    }
}
fn aopt(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns.is_empty() && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            bounded(&a.v)?;
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
pub(super) fn req(n: &Node, l: &str) -> Result<String> {
    aopt(n, l)?.ok_or_else(|| invalid(format!("missing required attribute '{l}'")))
}
fn bopt(n: &Node, l: &str) -> Result<Option<bool>> {
    match aopt(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn u32opt(n: &Node, l: &str) -> Result<Option<u32>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid u32 '{l}'"))))
        .transpose()
}
pub(super) fn u32req(n: &Node, l: &str) -> Result<u32> {
    u32opt(n, l)?.ok_or_else(|| invalid(format!("missing u32 '{l}'")))
}
fn u8opt(n: &Node, l: &str) -> Result<Option<u8>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid u8 '{l}'"))))
        .transpose()
}
fn u8req(n: &Node, l: &str) -> Result<u8> {
    u8opt(n, l)?.ok_or_else(|| invalid(format!("missing u8 '{l}'")))
}
fn i32opt(n: &Node, l: &str) -> Result<Option<i32>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid i32 '{l}'"))))
        .transpose()
}
fn only(n: &Node, a: &[&str]) -> Result<()> {
    for x in &n.attrs {
        if !x.ns.is_empty() || !a.contains(&x.l.as_str()) {
            // Attribute extensions are deliberately opaque. The package
            // transaction retains their original source span while typed
            // access continues to validate every known attribute.
            continue;
        }
    }
    Ok(())
}

pub(super) fn only_unqualified(n: &Node, allowed: &[&str]) -> Result<()> {
    for attribute in &n.attrs {
        if attribute.ns.is_empty() && !allowed.contains(&attribute.l.as_str()) {
            return Err(invalid(format!("unexpected attribute '{}'", attribute.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    only(n, &[])
}
fn set<T>(s: &mut Option<T>, v: T) -> Result<()> {
    if s.replace(v).is_some() {
        Err(invalid("duplicate connection property"))
    } else {
        Ok(())
    }
}
fn check_count(c: Option<u32>, actual: usize, n: &str) -> Result<()> {
    if c.is_some_and(|x| x as usize != actual) {
        Err(invalid(format!("{n} count mismatch")))
    } else {
        Ok(())
    }
}
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(b: &[(String, String)], p: &str) -> Result<String> {
    if p == "xml" {
        return Ok("http://www.w3.org/XML/1998/namespace".into());
    }
    b.iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
pub(super) fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING_BYTES {
        Err(invalid("connection string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn attr(x: &mut BoundedXml, n: &str, v: &str) -> Result<()> {
    x.push_char(' ')?;
    x.push_str(n)?;
    x.push_str("=\"")?;
    esc_bounded(x, v)?;
    x.push_char('"')
}
fn str_opt(x: &mut BoundedXml, n: &str, v: Option<&str>) -> Result<()> {
    if let Some(v) = v {
        attr(x, n, v)?;
    }
    Ok(())
}
fn bool_opt(x: &mut BoundedXml, n: &str, v: Option<bool>) -> Result<()> {
    if let Some(v) = v {
        attr(x, n, if v { "1" } else { "0" })?;
    }
    Ok(())
}
fn num<T: std::fmt::Display>(x: &mut BoundedXml, n: &str, v: T) -> Result<()> {
    attr(x, n, &v.to_string())
}
fn num_opt<T: std::fmt::Display>(x: &mut BoundedXml, n: &str, v: Option<T>) -> Result<()> {
    if let Some(v) = v {
        num(x, n, v)?;
    }
    Ok(())
}
fn esc_bounded(x: &mut BoundedXml, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '"' => x.push_str("&quot;")?,
            '\r' => x.push_str("&#xD;")?,
            '\n' => x.push_str("&#xA;")?,
            '\t' => x.push_str("&#x9;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
fn esc(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn text_escape(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! en{($p:ident,$w:ident,$t:ty,$($s:literal=>$v:path),+)=>{fn $p(s:String)->Result<$t>{match s.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{s}'")))}}fn $w(v:$t)->&'static str{match v{$($v=>$s,)+}}}}
en!(parse_credentials,credentials_str,CredentialsMethod,"integrated"=>CredentialsMethod::Integrated,"none"=>CredentialsMethod::None,"stored"=>CredentialsMethod::Stored);
en!(parse_html,html_str,HtmlFormatting,"none"=>HtmlFormatting::None,"rtf"=>HtmlFormatting::RichText,"all"=>HtmlFormatting::All);
en!(parse_file,file_str,TextFileType,"mac"=>TextFileType::Mac,"win"=>TextFileType::Windows,"dos"=>TextFileType::Dos);
en!(parse_qualifier,qualifier_str,TextQualifier,"doubleQuote"=>TextQualifier::DoubleQuote,"singleQuote"=>TextQualifier::SingleQuote,"none"=>TextQualifier::None);
en!(parse_parameter_type,parameter_str,ParameterType,"prompt"=>ParameterType::Prompt,"value"=>ParameterType::Value,"cell"=>ParameterType::Cell);
en!(parse_field,field_str,TextFieldType,"general"=>TextFieldType::General,"text"=>TextFieldType::Text,"MDY"=>TextFieldType::MonthDayYear,"DMY"=>TextFieldType::DayMonthYear,"YMD"=>TextFieldType::YearMonthDay,"MYD"=>TextFieldType::MonthYearDay,"DYM"=>TextFieldType::DayYearMonth,"YDM"=>TextFieldType::YearDayMonth,"skip"=>TextFieldType::Skip,"EMD"=>TextFieldType::EastAsianYearMonthDay);

fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

/// Patch typed connection fields inside their original XML spans.
///
/// Connection and query-table parts contain many producer extensions. The
/// transaction therefore changes scalar attributes in place and replaces only
/// the known property child whose typed value changed. Unrelated source bytes
/// remain untouched; structural collection edits retain existing connection
/// blocks and append new canonical blocks.
pub(super) fn patch_connections_source(
    source: &[u8],
    before: &Connections,
    after: &Connections,
    strict: bool,
) -> Result<Vec<u8>> {
    if before == after {
        return Ok(source.to_vec());
    }
    let tree = SourceTree::parse(source)?;
    if tree.nodes[tree.root].local != "connections" || tree.nodes[tree.root].self_closing {
        return after.to_xml(strict);
    }
    let nodes = tree.nodes[tree.root]
        .children
        .iter()
        .copied()
        .filter(|node| tree.nodes[*node].local == "connection")
        .collect::<Vec<_>>();
    if nodes.len() != before.connections.len() {
        return after.to_xml(strict);
    }
    let mut source_ids = Vec::with_capacity(nodes.len());
    for node in &nodes {
        let id = tree
            .attribute(source, *node, "id")?
            .ok_or_else(|| invalid("connection source is missing its id"))?
            .parse::<u32>()
            .map_err(|_| invalid("connection source has an invalid id"))?;
        source_ids.push(id);
    }
    let mut edits = Vec::new();
    for (position, connection) in after.connections.iter().enumerate() {
        if let Some(before_index) = before
            .connections
            .iter()
            .position(|candidate| candidate.id == connection.id)
        {
            let source_node = nodes
                .iter()
                .enumerate()
                .find(|(index, _)| source_ids[*index] == connection.id)
                .map(|(index, _)| index)
                .ok_or_else(|| invalid("connection source identity changed"))?;
            let node = nodes[source_node];
            let replacement = patch_connection_source(
                &source[tree.nodes[node].start..tree.nodes[node].end],
                &before.connections[before_index],
                connection,
                strict,
            )?;
            edits.push(SourceEdit {
                range: tree.nodes[node].start..tree.nodes[node].end,
                replacement,
            });
        } else {
            let _ = position;
            let replacement = canonical_connection(connection, strict)?;
            edits.push(SourceEdit {
                range: tree.nodes[tree.root].end_start..tree.nodes[tree.root].end_start,
                replacement,
            });
        }
    }
    for (index, node) in nodes.iter().enumerate() {
        if !after
            .connections
            .iter()
            .any(|connection| connection.id == source_ids[index])
        {
            edits.push(SourceEdit {
                range: tree.nodes[*node].start..tree.nodes[*node].end,
                replacement: Vec::new(),
            });
        }
    }
    apply_source_edits(source, edits)
}

fn patch_connection_source(
    source: &[u8],
    before: &Connection,
    after: &Connection,
    strict: bool,
) -> Result<Vec<u8>> {
    let tree = SourceTree::parse(source)?;
    let node = tree.root;
    if tree.nodes[node].local != "connection" {
        return Err(invalid("connection source has an invalid root"));
    }
    let mut edits = Vec::new();
    patch_optional_attr(
        &tree,
        node,
        "id",
        Some(&before.id.to_string()),
        Some(&after.id.to_string()),
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "sourceFile",
        before.source_file.as_deref(),
        after.source_file.as_deref(),
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "odcFile",
        before.odc_file.as_deref(),
        after.odc_file.as_deref(),
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "keepAlive",
        before.keep_alive,
        after.keep_alive,
        &mut edits,
    )?;
    patch_number(
        &tree,
        node,
        "interval",
        before.interval,
        after.interval,
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "name",
        before.name.as_deref(),
        after.name.as_deref(),
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "description",
        before.description.as_deref(),
        after.description.as_deref(),
        &mut edits,
    )?;
    patch_number(
        &tree,
        node,
        "type",
        before.connection_type,
        after.connection_type,
        &mut edits,
    )?;
    patch_number(
        &tree,
        node,
        "reconnectionMethod",
        before.reconnection_method,
        after.reconnection_method,
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "refreshedVersion",
        Some(&before.refreshed_version.to_string()),
        Some(&after.refreshed_version.to_string()),
        &mut edits,
    )?;
    patch_number(
        &tree,
        node,
        "minRefreshableVersion",
        before.min_refreshable_version,
        after.min_refreshable_version,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "savePassword",
        before.save_password,
        after.save_password,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "new",
        before.new_connection,
        after.new_connection,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "deleted",
        before.deleted,
        after.deleted,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "onlyUseConnectionFile",
        before.only_use_connection_file,
        after.only_use_connection_file,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "background",
        before.background,
        after.background,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "refreshOnLoad",
        before.refresh_on_load,
        after.refresh_on_load,
        &mut edits,
    )?;
    patch_bool(
        &tree,
        node,
        "saveData",
        before.save_data,
        after.save_data,
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "credentials",
        before.credentials.map(credentials_str),
        after.credentials.map(credentials_str),
        &mut edits,
    )?;
    patch_optional_attr(
        &tree,
        node,
        "singleSignOnId",
        before.single_sign_on_id.as_deref(),
        after.single_sign_on_id.as_deref(),
        &mut edits,
    )?;

    patch_child(
        &tree,
        source,
        node,
        "dbPr",
        before.database != after.database,
        after
            .database
            .is_some()
            .then(|| canonical_child(after, "dbPr", strict))
            .transpose()?,
        &mut edits,
    )?;
    patch_child(
        &tree,
        source,
        node,
        "olapPr",
        before.olap != after.olap,
        after
            .olap
            .is_some()
            .then(|| canonical_child(after, "olapPr", strict))
            .transpose()?,
        &mut edits,
    )?;
    patch_child(
        &tree,
        source,
        node,
        "webPr",
        before.web != after.web,
        after
            .web
            .is_some()
            .then(|| canonical_child(after, "webPr", strict))
            .transpose()?,
        &mut edits,
    )?;
    patch_child(
        &tree,
        source,
        node,
        "textPr",
        before.text != after.text,
        after
            .text
            .is_some()
            .then(|| canonical_child(after, "textPr", strict))
            .transpose()?,
        &mut edits,
    )?;
    patch_child(
        &tree,
        source,
        node,
        "parameters",
        before.parameters != after.parameters,
        after
            .parameters
            .is_some()
            .then(|| canonical_child(after, "parameters", strict))
            .transpose()?,
        &mut edits,
    )?;
    patch_child(
        &tree,
        source,
        node,
        "extLst",
        before.extension_xml != after.extension_xml,
        after
            .extension_xml
            .as_ref()
            .map(|_| canonical_child(after, "extLst", strict))
            .transpose()?,
        &mut edits,
    )?;
    apply_source_edits(source, edits)
}

fn canonical_connection(value: &Connection, strict: bool) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    write_connection(&mut output, value, strict)?;
    Ok(output.finish())
}

fn canonical_child(value: &Connection, name: &str, strict: bool) -> Result<Vec<u8>> {
    let source = canonical_connection(value, strict)?;
    let tree = SourceTree::parse(&source)?;
    let node = tree.nodes[tree.root]
        .children
        .iter()
        .copied()
        .find(|node| tree.nodes[*node].local == name)
        .ok_or_else(|| invalid(format!("canonical connection is missing '{name}'")))?;
    Ok(source[tree.nodes[node].start..tree.nodes[node].end].to_vec())
}

fn patch_child(
    tree: &SourceTree,
    _source: &[u8],
    parent: usize,
    name: &str,
    changed: bool,
    replacement: Option<Vec<u8>>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if !changed {
        return Ok(());
    }
    let existing = tree.child(parent, name)?;
    match (existing, replacement) {
        (Some(node), Some(replacement)) => edits.push(SourceEdit {
            range: tree.nodes[node].start..tree.nodes[node].end,
            replacement,
        }),
        (Some(node), None) => edits.push(SourceEdit {
            range: tree.nodes[node].start..tree.nodes[node].end,
            replacement: Vec::new(),
        }),
        (None, Some(replacement)) => {
            if tree.nodes[parent].self_closing {
                return Err(invalid(
                    "connection source needs a structural child insertion",
                ));
            }
            edits.push(SourceEdit {
                range: tree.nodes[parent].end_start..tree.nodes[parent].end_start,
                replacement,
            });
        },
        (None, None) => {},
    }
    Ok(())
}

fn patch_bool(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<bool>,
    after: Option<bool>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    patch_optional_attr(
        tree,
        node,
        name,
        before.map(|value| if value { "1" } else { "0" }),
        after.map(|value| if value { "1" } else { "0" }),
        edits,
    )
}

fn patch_number<T: std::fmt::Display + Copy>(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<T>,
    after: Option<T>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    let before = before.map(|value| value.to_string());
    let after = after.map(|value| value.to_string());
    patch_optional_attr(tree, node, name, before.as_deref(), after.as_deref(), edits)
}

fn patch_optional_attr(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let attribute = tree.nodes[node]
        .attrs
        .iter()
        .find(|attribute| attribute.local == name);
    match (attribute, after) {
        (Some(attribute), Some(value)) => edits.push(SourceEdit {
            range: attribute.value_start..attribute.value_end,
            replacement: escape_attribute(value),
        }),
        (Some(attribute), None) => edits.push(SourceEdit {
            range: attribute.start..attribute.value_end + 1,
            replacement: Vec::new(),
        }),
        (None, Some(value)) if before.is_none() => edits.push(SourceEdit {
            range: tree.nodes[node].close_pos..tree.nodes[node].close_pos,
            replacement: format!(
                " {name}=\"{}\"",
                String::from_utf8_lossy(&escape_attribute(value))
            )
            .into_bytes(),
        }),
        (None, Some(_)) => {
            return Err(invalid(format!(
                "connection source is missing attribute '{name}'"
            )));
        },
        (None, None) => {
            if before.is_some() {
                return Err(invalid(format!(
                    "connection source is missing attribute '{name}'"
                )));
            }
        },
    }
    Ok(())
}

fn escape_attribute(value: &str) -> Vec<u8> {
    let mut result = Vec::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => result.extend_from_slice(b"&amp;"),
            '<' => result.extend_from_slice(b"&lt;"),
            '"' => result.extend_from_slice(b"&quot;"),
            '\r' => result.extend_from_slice(b"&#xD;"),
            '\n' => result.extend_from_slice(b"&#xA;"),
            '\t' => result.extend_from_slice(b"&#x9;"),
            _ => {
                let mut encoded = [0; 4];
                result.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
    result
}

#[derive(Debug)]
struct SourceEdit {
    range: std::ops::Range<usize>,
    replacement: Vec<u8>,
}

fn apply_source_edits(source: &[u8], mut edits: Vec<SourceEdit>) -> Result<Vec<u8>> {
    edits.sort_by(|left, right| right.range.start.cmp(&left.range.start));
    for pair in edits.windows(2) {
        if pair[0].range.start < pair[1].range.end {
            return Err(invalid("connection source edits overlap"));
        }
    }
    let mut result = source.to_vec();
    for edit in edits {
        if edit.range.end > result.len() {
            return Err(invalid("connection source edit is out of bounds"));
        }
        result.splice(edit.range, edit.replacement);
    }
    Ok(result)
}

#[derive(Debug)]
struct SourceTree {
    nodes: Vec<SourceNode>,
    root: usize,
}

#[derive(Debug)]
struct SourceNode {
    local: String,
    start: usize,
    end_start: usize,
    end: usize,
    close_pos: usize,
    self_closing: bool,
    attrs: Vec<SourceAttribute>,
    children: Vec<usize>,
}

#[derive(Debug)]
struct SourceAttribute {
    local: String,
    start: usize,
    value_start: usize,
    value_end: usize,
}

impl SourceTree {
    fn parse(source: &[u8]) -> Result<Self> {
        let mut nodes: Vec<SourceNode> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut root: Option<usize> = None;
        let mut position = 0;
        while position < source.len() {
            if source[position] != b'<' {
                position += 1;
                continue;
            }
            if source[position..].starts_with(b"<?") {
                position = find_source_bytes(source, position + 2, b"?>")? + 2;
                continue;
            }
            if source[position..].starts_with(b"<!--") {
                position = find_source_bytes(source, position + 4, b"-->")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<![CDATA[") {
                position = find_source_bytes(source, position + 9, b"]]>")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<!") {
                position = source_tag_end(source, position)? + 1;
                continue;
            }
            if source[position..].starts_with(b"</") {
                let end = source_tag_end(source, position)?;
                let name_start = position + 2;
                let name_end = source_name_end(source, name_start);
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("connection source has an unmatched closing tag"))?;
                if nodes[node].local != source_local(&source[name_start..name_end])? {
                    return Err(invalid("connection source has mismatched tags"));
                }
                nodes[node].end_start = position;
                nodes[node].end = end + 1;
                position = end + 1;
                continue;
            }
            let end = source_tag_end(source, position)?;
            let (local, attrs, close_pos, self_closing) = source_start_tag(source, position, end)?;
            let node = nodes.len();
            nodes.push(SourceNode {
                local,
                start: position,
                end_start: end + 1,
                end: end + 1,
                close_pos,
                self_closing,
                attrs,
                children: Vec::new(),
            });
            if let Some(parent) = stack.last().copied() {
                nodes[parent].children.push(node);
            } else if root.replace(node).is_some() {
                return Err(invalid("connection source has multiple roots"));
            }
            if !self_closing {
                stack.push(node);
            }
            position = end + 1;
        }
        if !stack.is_empty() {
            return Err(invalid("connection source has unterminated markup"));
        }
        Ok(Self {
            nodes,
            root: root.ok_or_else(|| invalid("connection source has no root"))?,
        })
    }

    fn attribute(&self, source: &[u8], node: usize, name: &str) -> Result<Option<String>> {
        Ok(self.nodes[node]
            .attrs
            .iter()
            .find(|attribute| attribute.local == name)
            .map(|attribute| {
                std::str::from_utf8(&source[attribute.value_start..attribute.value_end])
                    .map(str::to_owned)
                    .map_err(xml_error)
            })
            .transpose()?)
    }

    fn child(&self, node: usize, name: &str) -> Result<Option<usize>> {
        let mut result = None;
        for child in &self.nodes[node].children {
            if self.nodes[*child].local == name {
                if result.replace(*child).is_some() {
                    return Err(invalid(format!("connection source has duplicate '{name}'")));
                }
            }
        }
        Ok(result)
    }
}

fn find_source_bytes(source: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    source[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| start + position)
        .ok_or_else(|| invalid("connection source has an unterminated declaration"))
}

fn source_tag_end(source: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (offset, byte) in source[start + 1..].iter().enumerate() {
        match (quote, byte) {
            (Some(value), byte) if *byte == value => quote = None,
            (None, b'\'' | b'\"') => quote = Some(*byte),
            (None, b'>') => return Ok(start + 1 + offset),
            _ => {},
        }
    }
    Err(invalid("connection source has an unterminated tag"))
}

fn source_start_tag(
    source: &[u8],
    start: usize,
    end: usize,
) -> Result<(String, Vec<SourceAttribute>, usize, bool)> {
    let name_start = start + 1;
    let name_end = source_name_end(source, name_start);
    if name_start == name_end {
        return Err(invalid("connection source has an empty element name"));
    }
    let mut attributes = Vec::new();
    let mut position = name_end;
    while position < end {
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] == b'/' {
            break;
        }
        let attr_start = position;
        let attr_end = source_name_end(source, position);
        if attr_end == attr_start {
            return Err(invalid("connection source has an invalid attribute"));
        }
        position = attr_end;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] != b'=' {
            return Err(invalid("connection source attribute is missing '='"));
        }
        position += 1;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        let quote = *source
            .get(position)
            .ok_or_else(|| invalid("connection source attribute is missing quotes"))?;
        if !matches!(quote, b'\'' | b'\"') {
            return Err(invalid("connection source attribute is missing quotes"));
        }
        position += 1;
        let value_start = position;
        while position < end && source[position] != quote {
            position += 1;
        }
        if position >= end {
            return Err(invalid("connection source attribute is unterminated"));
        }
        let value_end = position;
        position += 1;
        attributes.push(SourceAttribute {
            local: source_local(&source[attr_start..attr_end])?,
            start: attr_start,
            value_start,
            value_end,
        });
    }
    let self_closing = source[..end]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .is_some_and(|position| source[position] == b'/');
    let close_pos = if self_closing {
        source[..end]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .unwrap_or(end)
    } else {
        end
    };
    Ok((
        source_local(&source[name_start..name_end])?,
        attributes,
        close_pos,
        self_closing,
    ))
}

fn source_name_end(source: &[u8], mut position: usize) -> usize {
    while position < source.len()
        && !source[position].is_ascii_whitespace()
        && !matches!(source[position], b'/' | b'>' | b'=')
    {
        position += 1;
    }
    position
}

fn source_local(value: &[u8]) -> Result<String> {
    let value = std::str::from_utf8(value).map_err(xml_error)?;
    Ok(value.rsplit(':').next().unwrap_or(value).to_owned())
}
