use crate::sheet::eval::engine::EvalCtx;
#[cfg(feature = "eval_engine_web_functions")]
use crate::sheet::eval::engine::{evaluate_expression, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_encodeurl(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    #[cfg(feature = "eval_engine_web_functions")]
    {
        if args.len() != 1 {
            return Ok(CellValue::Error("ENCODEURL expects 1 argument".to_string()));
        }
        let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?);

        let encoded = urlencoding::encode(&text).to_string();
        Ok(CellValue::String(encoded))
    }
    #[cfg(not(feature = "eval_engine_web_functions"))]
    {
        let _ = (ctx, current_sheet, args);
        Ok(CellValue::Error("#NAME?".to_string()))
    }
}

pub(crate) async fn eval_webservice(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    #[cfg(feature = "eval_engine_web_functions")]
    {
        if args.len() != 1 {
            return Ok(CellValue::Error(
                "WEBSERVICE expects 1 argument".to_string(),
            ));
        }
        let url = evaluate_expression(ctx, current_sheet, &args[0]).await?;
        let url_str = to_text(&url);

        if url_str.len() > 2048 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }

        if !url_str.starts_with("http://") && !url_str.starts_with("https://") {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }

        let client = ctx.http_client();
        match client.get(&url_str).send().await {
            Ok(response) => match response.text().await {
                Ok(body) => {
                    if body.len() > 32767 {
                        Ok(CellValue::Error("#VALUE!".to_string()))
                    } else {
                        Ok(CellValue::String(body))
                    }
                },
                Err(_) => Ok(CellValue::Error("#VALUE!".to_string())),
            },
            Err(_) => Ok(CellValue::Error("#VALUE!".to_string())),
        }
    }
    #[cfg(not(feature = "eval_engine_web_functions"))]
    {
        let _ = (ctx, current_sheet, args);
        Ok(CellValue::Error("#NAME?".to_string()))
    }
}

pub(crate) async fn eval_filterxml(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    #[cfg(feature = "eval_engine_web_functions")]
    {
        if args.len() != 2 {
            return Ok(CellValue::Error(
                "FILTERXML expects 2 arguments (xml, xpath)".to_string(),
            ));
        }
        let xml_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
        let xml = to_text(&xml_val);
        let xpath_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        let xpath_str = to_text(&xpath_val);

        let package = match sxd_document::parser::parse(&xml) {
            Ok(p) => p,
            Err(_) => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        let document = package.as_document();

        let factory = sxd_xpath::Factory::new();
        let xpath = match factory.build(&xpath_str) {
            Ok(Some(xpath)) => xpath,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        };

        let context = sxd_xpath::Context::new();
        match xpath.evaluate(&context, document.root()) {
            Ok(value) => match value {
                sxd_xpath::Value::Nodeset(ns) => {
                    if let Some(node) = ns.document_order_first() {
                        Ok(CellValue::String(node.string_value()))
                    } else {
                        Ok(CellValue::Error("#VALUE!".to_string()))
                    }
                },
                sxd_xpath::Value::Boolean(b) => Ok(CellValue::Bool(b)),
                sxd_xpath::Value::Number(n) => Ok(CellValue::Float(n)),
                sxd_xpath::Value::String(s) => Ok(CellValue::String(s)),
            },
            Err(_) => Ok(CellValue::Error("#VALUE!".to_string())),
        }
    }
    #[cfg(not(feature = "eval_engine_web_functions"))]
    {
        let _ = (ctx, current_sheet, args);
        Ok(CellValue::Error("#NAME?".to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    // Tests for when eval_engine_web_functions feature is NOT enabled
    #[cfg(not(feature = "eval_engine_web_functions"))]
    mod no_feature_tests {
        use super::*;

        #[tokio::test]
        async fn test_encodeurl_without_feature() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("hello world")];
            let result = eval_encodeurl(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#NAME?".to_string()));
        }

        #[tokio::test]
        async fn test_webservice_without_feature() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("https://example.com")];
            let result = eval_webservice(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#NAME?".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_without_feature() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("<root></root>"), str_expr("/root")];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#NAME?".to_string()));
        }
    }

    // Tests for when eval_engine_web_functions feature IS enabled
    #[cfg(feature = "eval_engine_web_functions")]
    mod feature_tests {
        use super::*;

        #[tokio::test]
        async fn test_encodeurl_simple() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("hello world")];
            let result = eval_encodeurl(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::String("hello%20world".to_string()));
        }

        #[tokio::test]
        async fn test_encodeurl_special_chars() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("foo/bar+baz?key=value")];
            let result = eval_encodeurl(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(
                result,
                CellValue::String("foo%2Fbar%2Bbaz%3Fkey%3Dvalue".to_string())
            );
        }

        #[tokio::test]
        async fn test_encodeurl_wrong_args() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args: Vec<Expr> = vec![];
            let result = eval_encodeurl(ctx, "Sheet1", &args).await.unwrap();
            match result {
                CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
                _ => panic!("Expected Error result, got {:?}", result),
            }
        }

        #[tokio::test]
        async fn test_filterxml_basic() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<root><item>Hello</item></root>";
            let xpath = "/root/item";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::String("Hello".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_number_result() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<data><value>42.5</value></data>";
            let xpath = "/data/value";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            // XPath on element returns the text content as a string
            assert_eq!(result, CellValue::String("42.5".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_attribute() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<user id='123' name='John'/>";
            let xpath = "/user/@name";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::String("John".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_invalid_xml() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<invalid>unclosed";
            let xpath = "/root";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_invalid_xpath() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<root></root>";
            let xpath = "[[[invalid";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_no_match() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let xml = "<root><item>value</item></root>";
            let xpath = "/root/nonexistent";
            let args = vec![str_expr(xml), str_expr(xpath)];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
        }

        #[tokio::test]
        async fn test_filterxml_wrong_args() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args: Vec<Expr> = vec![str_expr("<root></root>")];
            let result = eval_filterxml(ctx, "Sheet1", &args).await.unwrap();
            match result {
                CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
                _ => panic!("Expected Error result, got {:?}", result),
            }
        }

        #[tokio::test]
        async fn test_webservice_wrong_args() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args: Vec<Expr> = vec![];
            let result = eval_webservice(ctx, "Sheet1", &args).await.unwrap();
            match result {
                CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
                _ => panic!("Expected Error result, got {:?}", result),
            }
        }

        #[tokio::test]
        async fn test_webservice_invalid_url() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let args = vec![str_expr("ftp://invalid.protocol.com")];
            let result = eval_webservice(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
        }

        #[tokio::test]
        async fn test_webservice_url_too_long() {
            let engine = TestEngine::new();
            let ctx = engine.ctx();
            let long_url = format!("https://example.com/{}", "a".repeat(2048));
            let args = vec![str_expr(&long_url)];
            let result = eval_webservice(ctx, "Sheet1", &args).await.unwrap();
            assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
        }
    }
}
