//! Cell values and formula-expression decoding.

use super::TableDataExtractor;
use crate::Result;
use crate::numbers::cell::CellValue;
use crate::protobuf::{tsce, tst};

impl TableDataExtractor<'_> {
    /// Parse a single Cell protobuf message into a CellValue
    pub(super) fn parse_cell(&self, cell: &tst::Cell) -> Result<CellValue> {
        use tst::CellValueType;

        match cell.value_type() {
            CellValueType::EmptyCellValueType => Ok(CellValue::Empty),

            CellValueType::NumberCellValueType => {
                let number = cell.number_value.unwrap_or(0.0);
                Ok(CellValue::Number(number))
            },

            CellValueType::StringCellValueType => {
                // String cells can store text directly or reference the string table
                if let Some(ref string_val) = cell.string_value
                    && !string_val.is_empty()
                {
                    Ok(CellValue::Text(string_val.clone()))
                } else {
                    // In some cases, strings are stored via references
                    // The actual string data would be in the string_table
                    // For a production implementation, we would:
                    // 1. Check cell.text_style or cell.cell_style for a reference
                    // 2. Resolve that reference to find the actual string value
                    // 3. Look up the string in the string_table using a key
                    //
                    // For now, return Empty if no direct string value is present
                    Ok(CellValue::Empty)
                }
            },

            CellValueType::BoolCellValueType => {
                let bool_val = cell.bool_value.unwrap_or(false);
                Ok(CellValue::Boolean(bool_val))
            },

            CellValueType::DateCellValueType => {
                // Date values are stored as numbers (timestamp)
                let date_num = cell.number_value.unwrap_or(0.0);
                // Convert to string representation
                // Note: Apple's date epoch is different from Unix epoch
                Ok(CellValue::Date(format!("{}", date_num)))
            },

            CellValueType::DurationCellValueType => {
                let duration = cell.number_value.unwrap_or(0.0);
                Ok(CellValue::Duration(duration))
            },

            CellValueType::ErrorCellValueType => Ok(CellValue::Error("ERROR".to_string())),

            CellValueType::ProvidedCellValueType => {
                // Provided values may come from formulas or other sources
                if let Some(ref formula) = cell.formula {
                    // Extract formula string representation
                    let formula_str = self.extract_formula_string(formula)?;
                    Ok(CellValue::Formula(formula_str))
                } else {
                    Ok(CellValue::Empty)
                }
            },

            CellValueType::RichTextCellType => {
                // Rich text requires resolving the richTextPayload reference
                if let Some(ref payload_ref) = cell.rich_text_payload {
                    if let Some(text) = self.extract_rich_text(payload_ref.identifier)? {
                        Ok(CellValue::Text(text))
                    } else {
                        Ok(CellValue::Empty)
                    }
                } else {
                    Ok(CellValue::Empty)
                }
            },
        }
    }

    /// Extract formula string from FormulaArchive
    ///
    ///   - Reconstructs formula text from Abstract Syntax Tree
    ///   - Handles operators, functions, cell references, and constants
    ///   - Based on TSCE.ASTNodeArrayArchive protobuf structure
    ///   - Implements reverse-polish notation to infix conversion
    ///
    /// iWork stores formulas as Abstract Syntax Trees (AST) in reverse-polish
    /// notation (postfix). This function reconstructs the formula text by
    /// traversing the AST and converting it to standard infix notation.
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of AST nodes. Uses a stack-based algorithm
    /// for efficient conversion.
    fn extract_formula_string(&self, formula: &tsce::FormulaArchive) -> Result<String> {
        use crate::protobuf::tsce::ast_node_array_archive::AstNodeType;

        let ast_array = &formula.ast_node_array;

        // Formulas are stored in reverse-polish notation (postfix)
        // We need to convert to infix notation using a stack
        if ast_array.ast_node.is_empty() {
            return Ok("=".to_string());
        }

        // Stack to hold expression parts during reconstruction
        let mut expr_stack: Vec<String> = Vec::new();

        // Process each AST node
        for node in &ast_array.ast_node {
            let ast_node_type = node.ast_node_type();

            match ast_node_type {
                // Arithmetic operators (binary)
                AstNodeType::AdditionNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}+{})", left, right));
                },
                AstNodeType::SubtractionNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}-{})", left, right));
                },
                AstNodeType::MultiplicationNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}*{})", left, right));
                },
                AstNodeType::DivisionNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}/{})", left, right));
                },
                AstNodeType::PowerNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}^{})", left, right));
                },

                // Note: Comparison operators are handled differently in Numbers AST
                // They're not separate node types but may be represented through function nodes
                // For simplicity, we skip explicit handling here and rely on function dispatch

                // Constants
                AstNodeType::NumberNode => {
                    if let Some(number) = node.ast_number_node_number {
                        expr_stack.push(number.to_string());
                    }
                },
                AstNodeType::StringNode => {
                    if let Some(ref string) = node.ast_string_node_string {
                        expr_stack.push(format!("\"{}\"", string));
                    }
                },
                AstNodeType::BooleanNode => {
                    if let Some(boolean) = node.ast_boolean_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_string());
                    }
                },

                // Cell references
                AstNodeType::CellReferenceNode => {
                    if let Some(ref cell_ref) = node.ast_local_cell_reference_node_reference {
                        // Convert row/column handles to A1 notation
                        let col_letter = self.column_index_to_letter(cell_ref.column_handle);
                        let row_num = cell_ref.row_handle + 1; // 0-based to 1-based
                        let col_sticky = if cell_ref.column_is_sticky != 0 {
                            "$"
                        } else {
                            ""
                        };
                        let row_sticky = if cell_ref.row_is_sticky != 0 { "$" } else { "" };
                        expr_stack.push(format!(
                            "{}{}{}{}",
                            col_sticky, col_letter, row_sticky, row_num
                        ));
                    } else if let Some(ref cross_ref) =
                        node.ast_cross_table_cell_reference_node_reference
                    {
                        // Cross-table reference
                        let col_letter = self.column_index_to_letter(cross_ref.column_handle);
                        let row_num = cross_ref.row_handle + 1;
                        expr_stack.push(format!("{}::{}{}", "Table", col_letter, row_num));
                    }
                },

                // Functions
                AstNodeType::FunctionNode => {
                    if let Some(function_index) = node.ast_function_node_index {
                        let num_args = node.ast_function_node_num_args.unwrap_or(0);
                        let function_name = self.get_function_name(function_index);

                        // Pop arguments from stack (in reverse order)
                        let mut args = Vec::new();
                        for _ in 0..num_args {
                            if let Some(arg) = expr_stack.pop() {
                                args.push(arg);
                            }
                        }
                        args.reverse();

                        let args_str = args.join(",");
                        expr_stack.push(format!("{}({})", function_name, args_str));
                    }
                },

                // List (for function arguments)
                AstNodeType::ListNode => {
                    if let Some(num_args) = node.ast_list_node_num_args {
                        // Collect arguments
                        let mut args = Vec::new();
                        for _ in 0..num_args {
                            if let Some(arg) = expr_stack.pop() {
                                args.push(arg);
                            }
                        }
                        args.reverse();
                        expr_stack.push(args.join(","));
                    }
                },

                // Unary operators - represented differently in the AST
                // Numbers uses NegationNode instead of UnaryMinusNode
                AstNodeType::NegationNode => {
                    if let Some(operand) = expr_stack.pop() {
                        expr_stack.push(format!("-({})", operand));
                    }
                },

                // Concatenation
                AstNodeType::ConcatenationNode if expr_stack.len() >= 2 => {
                    let right = expr_stack.pop().unwrap();
                    let left = expr_stack.pop().unwrap();
                    expr_stack.push(format!("({}&{})", left, right));
                },

                // Other node types - handle gracefully
                _ => {
                    // Unknown or special node types - keep processing
                    // (e.g., whitespace nodes, thunk nodes, etc.)
                },
            }
        }

        // The final result should be on top of the stack
        let result = if expr_stack.is_empty() {
            "=FORMULA()".to_string()
        } else {
            format!("={}", expr_stack.pop().unwrap())
        };

        Ok(result)
    }

    /// Convert column index to Excel-style letter (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
    fn column_index_to_letter(&self, index: u32) -> String {
        let mut result = String::new();
        let mut idx = index;

        loop {
            let remainder = idx % 26;
            result.insert(0, (b'A' + remainder as u8) as char);
            if idx < 26 {
                break;
            }
            idx = idx / 26 - 1;
        }

        result
    }

    /// Get function name from function index
    /// Based on Numbers built-in function list
    fn get_function_name(&self, index: u32) -> String {
        // Common function indices (based on analysis of Numbers documents)
        // This mapping comes from observing Numbers files and documentation
        match index {
            0 => "SUM".to_string(),
            1 => "AVERAGE".to_string(),
            2 => "COUNT".to_string(),
            3 => "MAX".to_string(),
            4 => "MIN".to_string(),
            5 => "PRODUCT".to_string(),
            6 => "IF".to_string(),
            7 => "AND".to_string(),
            8 => "OR".to_string(),
            9 => "NOT".to_string(),
            10 => "ROUND".to_string(),
            11 => "SQRT".to_string(),
            12 => "ABS".to_string(),
            13 => "CONCATENATE".to_string(),
            14 => "LEFT".to_string(),
            15 => "RIGHT".to_string(),
            16 => "MID".to_string(),
            17 => "LEN".to_string(),
            18 => "UPPER".to_string(),
            19 => "LOWER".to_string(),
            20 => "PROPER".to_string(),
            21 => "TRIM".to_string(),
            22 => "SUBSTITUTE".to_string(),
            23 => "FIND".to_string(),
            24 => "SEARCH".to_string(),
            25 => "NOW".to_string(),
            26 => "TODAY".to_string(),
            27 => "DATE".to_string(),
            28 => "TIME".to_string(),
            29 => "YEAR".to_string(),
            30 => "MONTH".to_string(),
            31 => "DAY".to_string(),
            32 => "HOUR".to_string(),
            33 => "MINUTE".to_string(),
            34 => "SECOND".to_string(),
            35 => "WEEKDAY".to_string(),
            36 => "VLOOKUP".to_string(),
            37 => "HLOOKUP".to_string(),
            38 => "INDEX".to_string(),
            39 => "MATCH".to_string(),
            40 => "CHOOSE".to_string(),
            // More functions exist, but these are the most common
            _ => format!("FUNC{}", index),
        }
    }
}
