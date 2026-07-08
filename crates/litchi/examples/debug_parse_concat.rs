use litchi::sheet::eval::parser;

fn main() {
    let expr = "CONCAT(\"Hello\",\" \",\"World\")";
    println!("input: {}", expr);
    match parser::parse_expression("Funcs", expr) {
        Some(ast) => println!("parsed: {:?}", ast),
        None => println!("parse_expression returned None"),
    }
}
