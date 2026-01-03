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
