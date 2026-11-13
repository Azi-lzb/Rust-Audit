//! Restricted Decimal expression evaluator used by the S2 rule engine.
//!
//! The Python implementation evaluates a restricted expression AST with no
//! builtins and a 50 digit Decimal context. This module implements the numeric,
//! comparison, boolean, and AND/OR/NOT subset used by S2 rules without Python
//! or a general-purpose `eval` facility.

use crate::unit_conversion::{as_decimal, DecimalTuple};
use anyhow::{anyhow, bail, Result};
use std::cmp::Ordering;

const PRECISION: usize = 50;
const MAX_SCALE_SHIFT: usize = 100_000;

#[derive(Clone, Debug, PartialEq)]
enum Value {
    Decimal(DecimalTuple),
    Boolean(bool),
}

impl Value {
    fn truthy(&self) -> bool {
        match self {
            Self::Decimal(value) => value.digits.iter().any(|digit| *digit != 0),
            Self::Boolean(value) => *value,
        }
    }

    fn decimal(self) -> Result<DecimalTuple> {
        match self {
            Self::Decimal(value) => Ok(value),
            Self::Boolean(value) => Ok(DecimalTuple {
                sign: 0,
                digits: vec![u8::from(value)],
                exponent: 0,
            }),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(String),
    Identifier(String),
    Plus,
    Minus,
    Star,
    Power,
    Slash,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    Equal,
    NotEqual,
    LeftParen,
    RightParen,
    Comma,
    End,
}

/// Evaluate one already-substituted S2 expression, then apply Python truthiness.
pub(crate) fn evaluate(expression: &str) -> Result<bool> {
    let normalized = normalize_operators(expression);
    let tokens = tokenize(&normalized)?;
    let mut parser = Parser {
        tokens,
        position: 0,
        evaluating: true,
    };
    let value = parser.parse_or()?;
    if parser.peek() != &Token::End {
        bail!("表达式包含无法识别的内容：{}", parser.token_description());
    }
    Ok(value.truthy())
}

fn normalize_operators(expression: &str) -> String {
    let chars = expression.chars().collect::<Vec<_>>();
    let mut output = String::with_capacity(expression.len());
    let mut index = 0;
    while index < chars.len() {
        if index + 2 < chars.len() && chars[index..index + 3] == ['=', '=', '='] {
            output.push_str("==");
            index += 3;
        } else if index + 1 < chars.len() && chars[index] == '<' && chars[index + 1] == '>' {
            output.push_str("!=");
            index += 2;
        } else if chars[index] == '=' {
            let previous = index.checked_sub(1).and_then(|i| chars.get(i)).copied();
            let next = chars.get(index + 1).copied();
            if !matches!(previous, Some('<' | '>' | '=' | '!')) && next != Some('=') {
                output.push_str("==");
            } else {
                output.push('=');
            }
            index += 1;
        } else {
            output.push(chars[index]);
            index += 1;
        }
    }
    output
}

fn tokenize(input: &str) -> Result<Vec<Token>> {
    let chars = input.chars().collect::<Vec<_>>();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < chars.len() {
        let ch = chars[index];
        if ch.is_whitespace() {
            index += 1;
            continue;
        }
        if ch.is_ascii_digit()
            || (ch == '.'
                && chars
                    .get(index + 1)
                    .is_some_and(|next| next.is_ascii_digit()))
        {
            let start = index;
            let mut saw_dot = false;
            if ch == '.' {
                saw_dot = true;
                index += 1;
            }
            while index < chars.len() && (chars[index].is_ascii_digit() || chars[index] == '_') {
                index += 1;
            }
            if !saw_dot && chars.get(index) == Some(&'.') {
                index += 1;
                while index < chars.len() && (chars[index].is_ascii_digit() || chars[index] == '_')
                {
                    index += 1;
                }
            }
            if matches!(chars.get(index), Some('e' | 'E')) {
                index += 1;
                if matches!(chars.get(index), Some('+' | '-')) {
                    index += 1;
                }
                let exponent_start = index;
                while index < chars.len() && (chars[index].is_ascii_digit() || chars[index] == '_')
                {
                    index += 1;
                }
                if index == exponent_start {
                    bail!("数值指数缺少数字");
                }
            }
            tokens.push(Token::Number(chars[start..index].iter().collect()));
            continue;
        }
        if ch == '_' || ch.is_alphabetic() {
            let start = index;
            index += 1;
            while index < chars.len() && (chars[index] == '_' || chars[index].is_alphanumeric()) {
                index += 1;
            }
            tokens.push(Token::Identifier(chars[start..index].iter().collect()));
            continue;
        }
        let next = chars.get(index + 1).copied();
        match (ch, next) {
            ('*', Some('*')) => {
                tokens.push(Token::Power);
                index += 2;
            }
            ('<', Some('=')) => {
                tokens.push(Token::LessEqual);
                index += 2;
            }
            ('>', Some('=')) => {
                tokens.push(Token::GreaterEqual);
                index += 2;
            }
            ('=', Some('=')) => {
                tokens.push(Token::Equal);
                index += 2;
            }
            ('!', Some('=')) => {
                tokens.push(Token::NotEqual);
                index += 2;
            }
            ('+', _) => {
                tokens.push(Token::Plus);
                index += 1;
            }
            ('-', _) => {
                tokens.push(Token::Minus);
                index += 1;
            }
            ('*', _) => {
                tokens.push(Token::Star);
                index += 1;
            }
            ('/', _) => {
                tokens.push(Token::Slash);
                index += 1;
            }
            ('<', _) => {
                tokens.push(Token::Less);
                index += 1;
            }
            ('>', _) => {
                tokens.push(Token::Greater);
                index += 1;
            }
            ('(', _) => {
                tokens.push(Token::LeftParen);
                index += 1;
            }
            (')', _) => {
                tokens.push(Token::RightParen);
                index += 1;
            }
            (',', _) => {
                tokens.push(Token::Comma);
                index += 1;
            }
            _ => bail!("表达式包含不支持的字符：{ch}"),
        }
    }
    tokens.push(Token::End);
    Ok(tokens)
}

struct Parser {
    tokens: Vec<Token>,
    position: usize,
    evaluating: bool,
}

impl Parser {
    fn peek(&self) -> &Token {
        self.tokens.get(self.position).unwrap_or(&Token::End)
    }

    fn take(&mut self) -> Token {
        let token = self.peek().clone();
        self.position = (self.position + 1).min(self.tokens.len());
        token
    }

    fn token_description(&self) -> String {
        format!("{:?}", self.peek())
    }

    // Skipped branches are still parsed, but perform no runtime operations.
    fn parse_with_evaluation(
        &mut self,
        enabled: bool,
        parse: fn(&mut Self) -> Result<Value>,
    ) -> Result<Value> {
        let previous = self.evaluating;
        self.evaluating = previous && enabled;
        let result = parse(self);
        self.evaluating = previous;
        result
    }

    fn parse_or(&mut self) -> Result<Value> {
        let mut left = self.parse_and()?;
        while matches!(self.peek(), Token::Identifier(word) if word.eq_ignore_ascii_case("or")) {
            self.take();
            let right = self.parse_with_evaluation(!left.truthy(), Self::parse_and)?;
            if !left.truthy() {
                left = right;
            }
        }
        Ok(left)
    }

    fn parse_and(&mut self) -> Result<Value> {
        let mut left = self.parse_not()?;
        while matches!(self.peek(), Token::Identifier(word) if word.eq_ignore_ascii_case("and")) {
            self.take();
            let right = self.parse_with_evaluation(left.truthy(), Self::parse_not)?;
            if left.truthy() {
                left = right;
            }
        }
        Ok(left)
    }

    fn parse_not(&mut self) -> Result<Value> {
        if matches!(self.peek(), Token::Identifier(word) if word.eq_ignore_ascii_case("not")) {
            self.take();
            return Ok(Value::Boolean(!self.parse_not()?.truthy()));
        }
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Value> {
        let mut left = self.parse_sum()?;
        let mut compared = false;
        let mut matches = true;
        loop {
            let operation = match self.peek() {
                Token::Less => Some("<"),
                Token::LessEqual => Some("<="),
                Token::Greater => Some(">"),
                Token::GreaterEqual => Some(">="),
                Token::Equal => Some("=="),
                Token::NotEqual => Some("!="),
                _ => None,
            };
            let Some(operation) = operation else { break };
            self.take();
            let right = self.parse_with_evaluation(matches, Self::parse_sum)?;
            if self.evaluating && matches {
                matches = compare_values(&left, &right, operation)?;
            }
            left = right;
            compared = true;
        }
        Ok(if compared {
            Value::Boolean(matches)
        } else {
            left
        })
    }

    fn parse_sum(&mut self) -> Result<Value> {
        let mut left = self.parse_product()?;
        loop {
            let operation = match self.peek() {
                Token::Plus => '+',
                Token::Minus => '-',
                _ => break,
            };
            self.take();
            let right = self.parse_product()?;
            if self.evaluating {
                left = arithmetic(left, right, operation)?;
            }
        }
        Ok(left)
    }

    fn parse_product(&mut self) -> Result<Value> {
        let mut left = self.parse_unary()?;
        loop {
            let operation = match self.peek() {
                Token::Star => '*',
                Token::Slash => '/',
                _ => break,
            };
            self.take();
            let right = self.parse_unary()?;
            if self.evaluating {
                left = arithmetic(left, right, operation)?;
            }
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<Value> {
        match self.peek() {
            Token::Plus => {
                self.take();
                self.parse_unary()
            }
            Token::Minus => {
                self.take();
                let value = self.parse_unary()?;
                if !self.evaluating {
                    return Ok(Value::Boolean(false));
                }
                let mut number = value.decimal()?;
                if number.digits.iter().any(|digit| *digit != 0) {
                    number.sign ^= 1;
                }
                Ok(Value::Decimal(number))
            }
            _ => self.parse_power(),
        }
    }

    fn parse_power(&mut self) -> Result<Value> {
        let left = self.parse_primary()?;
        if self.peek() == &Token::Power {
            self.take();
            let exponent = self.parse_unary()?;
            if !self.evaluating {
                return Ok(Value::Boolean(false));
            }
            return Ok(Value::Decimal(decimal_power(
                left.decimal()?,
                &exponent.decimal()?,
            )?));
        }
        Ok(left)
    }

    fn parse_primary(&mut self) -> Result<Value> {
        match self.take() {
            Token::Number(text) => {
                let text = text.replace('_', "");
                as_decimal(&text)
                    .map(Value::Decimal)
                    .map_err(|error| anyhow!("{error}"))
            }
            Token::Identifier(name) if name == "True" => Ok(Value::Boolean(true)),
            Token::Identifier(name) if name == "False" => Ok(Value::Boolean(false)),
            Token::Identifier(name) => {
                if self.peek() == &Token::LeftParen {
                    self.take();
                    let mut arguments = Vec::new();
                    if self.peek() != &Token::RightParen {
                        loop {
                            arguments.push(self.parse_or()?);
                            if self.peek() != &Token::Comma {
                                break;
                            }
                            self.take();
                        }
                    }
                    if self.take() != Token::RightParen {
                        bail!("函数调用缺少右括号");
                    }
                    if self.evaluating {
                        call_function(&name, &arguments)
                    } else {
                        Ok(Value::Boolean(false))
                    }
                } else if !self.evaluating {
                    Ok(Value::Boolean(false))
                } else {
                    bail!("name '{name}' is not defined")
                }
            }
            Token::LeftParen => {
                let value = self.parse_or()?;
                if self.take() != Token::RightParen {
                    bail!("表达式缺少右括号");
                }
                Ok(value)
            }
            token => bail!("表达式缺少操作数：{token:?}"),
        }
    }
}

fn call_function(name: &str, arguments: &[Value]) -> Result<Value> {
    let normalized = name.trim_start_matches('_').to_ascii_uppercase();
    match normalized.as_str() {
        "AND" => Ok(Value::Boolean(arguments.iter().all(Value::truthy))),
        "OR" => Ok(Value::Boolean(arguments.iter().any(Value::truthy))),
        "NOT" if arguments.len() == 1 => Ok(Value::Boolean(!arguments[0].truthy())),
        "NOT" => bail!("_NOT() takes exactly one argument"),
        _ => bail!("name '{name}' is not defined"),
    }
}

fn arithmetic(left: Value, right: Value, operation: char) -> Result<Value> {
    let left = left.decimal()?;
    let right = right.decimal()?;
    let result = match operation {
        '+' => decimal_add(&left, &right)?,
        '-' => decimal_add(&left, &negate(right))?,
        '*' => decimal_multiply(&left, &right)?,
        '/' => decimal_divide(&left, &right)?,
        _ => unreachable!("only supported operators reach arithmetic"),
    };
    Ok(Value::Decimal(result))
}

fn compare_values(left: &Value, right: &Value, operation: &str) -> Result<bool> {
    let left = match left {
        Value::Decimal(value) => value.clone(),
        Value::Boolean(value) => bool_decimal(*value),
    };
    let right = match right {
        Value::Decimal(value) => value.clone(),
        Value::Boolean(value) => bool_decimal(*value),
    };
    let ordering = compare_decimal(&left, &right);
    Ok(match operation {
        "<" => ordering == Ordering::Less,
        "<=" => ordering != Ordering::Greater,
        ">" => ordering == Ordering::Greater,
        ">=" => ordering != Ordering::Less,
        "==" => ordering == Ordering::Equal,
        "!=" => ordering != Ordering::Equal,
        _ => unreachable!("only comparisons reach compare_values"),
    })
}

fn bool_decimal(value: bool) -> DecimalTuple {
    DecimalTuple {
        sign: 0,
        digits: vec![u8::from(value)],
        exponent: 0,
    }
}

fn negate(mut value: DecimalTuple) -> DecimalTuple {
    if value.digits.iter().any(|digit| *digit != 0) {
        value.sign ^= 1;
    }
    value
}

fn decimal_add(left: &DecimalTuple, right: &DecimalTuple) -> Result<DecimalTuple> {
    let exponent = left.exponent.min(right.exponent);
    let left_shift = usize::try_from(i128::from(left.exponent) - i128::from(exponent))
        .map_err(|_| anyhow!("Decimal 指数差超出可处理范围"))?;
    let right_shift = usize::try_from(i128::from(right.exponent) - i128::from(exponent))
        .map_err(|_| anyhow!("Decimal 指数差超出可处理范围"))?;
    if left_shift > MAX_SCALE_SHIFT || right_shift > MAX_SCALE_SHIFT {
        bail!("Decimal 指数差超出可处理范围");
    }
    let mut left_digits = normalized_digits(&left.digits);
    let mut right_digits = normalized_digits(&right.digits);
    left_digits.resize(left_digits.len() + left_shift, 0);
    right_digits.resize(right_digits.len() + right_shift, 0);
    let (sign, digits) = if left.sign == right.sign {
        (left.sign, add_abs(&left_digits, &right_digits))
    } else {
        match compare_abs(&left_digits, &right_digits) {
            Ordering::Equal => (0, vec![0]),
            Ordering::Greater => (left.sign, subtract_abs(&left_digits, &right_digits)),
            Ordering::Less => (right.sign, subtract_abs(&right_digits, &left_digits)),
        }
    };
    round_context(DecimalTuple {
        sign,
        digits,
        exponent,
    })
}

fn decimal_multiply(left: &DecimalTuple, right: &DecimalTuple) -> Result<DecimalTuple> {
    let exponent = left
        .exponent
        .checked_add(right.exponent)
        .ok_or_else(|| anyhow!("Decimal 指数超出可处理范围"))?;
    if is_zero(left) || is_zero(right) {
        return Ok(DecimalTuple {
            sign: left.sign ^ right.sign,
            digits: vec![0],
            exponent,
        });
    }
    round_context(DecimalTuple {
        sign: left.sign ^ right.sign,
        digits: multiply_abs(&left.digits, &right.digits),
        exponent,
    })
}

fn decimal_divide(left: &DecimalTuple, right: &DecimalTuple) -> Result<DecimalTuple> {
    if is_zero(right) {
        bail!("DivisionByZero");
    }
    if is_zero(left) {
        return Ok(DecimalTuple {
            sign: left.sign ^ right.sign,
            digits: vec![0],
            exponent: 0,
        });
    }
    let left_digits = normalized_digits(&left.digits);
    let right_digits = normalized_digits(&right.digits);
    let left_significand = padded_digits(&left_digits, right_digits.len());
    let right_significand = padded_digits(&right_digits, left_digits.len());
    let coefficient_exponent = left_digits.len() as i64
        - right_digits.len() as i64
        - i64::from(left_significand < right_significand);
    // Keep intermediate order wider than the stored exponent: normalization may
    // bring it back into i64 (e.g. 1e<i64::MIN> / 1).
    let exponent_difference = i128::from(left.exponent) - i128::from(right.exponent);
    let result_order = i128::from(coefficient_exponent) + exponent_difference;
    let numerator_shift = PRECISION as i64 - 1 - coefficient_exponent;
    let mut numerator = left_digits;
    let mut denominator = right_digits;
    if numerator_shift >= 0 {
        if numerator_shift as usize > MAX_SCALE_SHIFT {
            bail!("Decimal division precision shift exceeds supported range");
        }
        numerator.resize(numerator.len() + numerator_shift as usize, 0);
    } else {
        if (-numerator_shift) as usize > MAX_SCALE_SHIFT {
            bail!("Decimal division precision shift exceeds supported range");
        }
        denominator.resize(denominator.len() + (-numerator_shift) as usize, 0);
    }
    let (mut quotient, remainder) = divide_abs(&numerator, &denominator);
    let rounding = compare_abs(&multiply_small(&remainder, 2), &denominator);
    if rounding == Ordering::Greater
        || (rounding == Ordering::Equal && quotient.last().copied().unwrap_or(0) % 2 == 1)
    {
        increment_digits(&mut quotient);
    }
    let mut exponent = result_order - (PRECISION as i128 - 1);
    if quotient.len() > PRECISION {
        quotient.pop();
        exponent += 1;
    }
    while quotient.len() > 1 && quotient.last() == Some(&0) {
        quotient.pop();
        exponent += 1;
    }
    Ok(DecimalTuple {
        sign: left.sign ^ right.sign,
        digits: quotient,
        exponent: i64::try_from(exponent).map_err(|_| anyhow!("Decimal 指数超出可处理范围"))?,
    })
}

fn decimal_power(base: DecimalTuple, exponent: &DecimalTuple) -> Result<DecimalTuple> {
    let power = bounded_integer_power(exponent)?;
    if power < 0 && is_zero(&base) {
        bail!("DivisionByZero");
    }
    let mut remaining = power.unsigned_abs();
    let mut factor = base;
    let mut result = bool_decimal(true);
    while remaining > 0 {
        if remaining % 2 == 1 {
            result = decimal_multiply(&result, &factor)?;
        }
        remaining /= 2;
        if remaining > 0 {
            factor = decimal_multiply(&factor, &factor)?;
        }
    }
    if power < 0 {
        decimal_divide(&bool_decimal(true), &result)
    } else {
        Ok(result)
    }
}

/// Inspect the compact Decimal tuple directly; do not expand exponent padding.
fn bounded_integer_power(exponent: &DecimalTuple) -> Result<i64> {
    if is_zero(exponent) {
        return Ok(0);
    }
    let digits = if exponent.exponent < 0 {
        let fractional = usize::try_from(exponent.exponent.unsigned_abs())
            .map_err(|_| anyhow!("Decimal power currently requires an integer exponent"))?;
        if fractional > exponent.digits.len()
            || exponent.digits[exponent.digits.len() - fractional..]
                .iter()
                .any(|digit| *digit != 0)
        {
            bail!("Decimal power currently requires an integer exponent");
        }
        &exponent.digits[..exponent.digits.len() - fractional]
    } else {
        &exponent.digits[..]
    };
    let mut power = 0_i64;
    for digit in digits {
        power = power * 10 + i64::from(*digit);
        if power > 10_000 {
            bail!("Decimal power exceeds supported exponent range (-10000..10000)");
        }
    }
    if exponent.exponent > 0 {
        if exponent.exponent > 4 {
            bail!("Decimal power exceeds supported exponent range (-10000..10000)");
        }
        for _ in 0..exponent.exponent {
            power *= 10;
            if power > 10_000 {
                bail!("Decimal power exceeds supported exponent range (-10000..10000)");
            }
        }
    }
    Ok(if exponent.sign == 1 { -power } else { power })
}

fn round_context(mut value: DecimalTuple) -> Result<DecimalTuple> {
    let mut digits = normalized_digits(&value.digits);
    if digits.len() <= PRECISION {
        value.digits = digits;
        return Ok(value);
    }
    let dropped = digits.len() - PRECISION;
    let first = digits[PRECISION];
    let rest_nonzero = digits[PRECISION + 1..].iter().any(|digit| *digit != 0);
    let last = digits[PRECISION - 1];
    let round_up = first > 5 || (first == 5 && (rest_nonzero || last % 2 == 1));
    digits.truncate(PRECISION);
    if round_up {
        increment_digits(&mut digits);
    }
    if digits.len() > PRECISION {
        digits.pop();
        digits.fill(0);
        digits[0] = 1;
        value.exponent = value
            .exponent
            .checked_add(i64::try_from(dropped)?)
            .and_then(|e| e.checked_add(1))
            .ok_or_else(|| anyhow!("Decimal 指数超出可处理范围"))?;
    } else {
        value.exponent = value
            .exponent
            .checked_add(i64::try_from(dropped)?)
            .ok_or_else(|| anyhow!("Decimal 指数超出可处理范围"))?;
    }
    value.digits = digits;
    Ok(value)
}

fn compare_decimal(left: &DecimalTuple, right: &DecimalTuple) -> Ordering {
    let left_zero = is_zero(left);
    let right_zero = is_zero(right);
    match (left_zero, right_zero) {
        (true, true) => return Ordering::Equal,
        (true, false) => {
            return if right.sign == 1 {
                Ordering::Greater
            } else {
                Ordering::Less
            }
        }
        (false, true) => {
            return if left.sign == 1 {
                Ordering::Less
            } else {
                Ordering::Greater
            }
        }
        _ => {}
    }
    match (left.sign, right.sign) {
        (0, 1) => Ordering::Greater,
        (1, 0) => Ordering::Less,
        _ => {
            let ordering = compare_abs_decimal(left, right);
            if left.sign == 1 {
                ordering.reverse()
            } else {
                ordering
            }
        }
    }
}

fn compare_abs_decimal(left: &DecimalTuple, right: &DecimalTuple) -> Ordering {
    let left_digits = normalized_digits(&left.digits);
    let right_digits = normalized_digits(&right.digits);
    let left_order = i128::from(left.exponent) + left_digits.len() as i128;
    let right_order = i128::from(right.exponent) + right_digits.len() as i128;
    match left_order.cmp(&right_order) {
        Ordering::Equal => {
            let count = left_digits.len().max(right_digits.len());
            for index in 0..count {
                let left = left_digits.get(index).copied().unwrap_or(0);
                let right = right_digits.get(index).copied().unwrap_or(0);
                match left.cmp(&right) {
                    Ordering::Equal => {}
                    ordering => return ordering,
                }
            }
            Ordering::Equal
        }
        ordering => ordering,
    }
}

fn is_zero(value: &DecimalTuple) -> bool {
    value.digits.iter().all(|digit| *digit == 0)
}

pub(crate) fn decimal_plain(value: &DecimalTuple) -> Result<String> {
    crate::unit_conversion::decimal_plain_limited(value, MAX_SCALE_SHIFT)
        .map_err(|error| anyhow!("{error}"))
}

pub(crate) fn python_float_repr(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value == f64::INFINITY {
        return "inf".to_owned();
    }
    if value == f64::NEG_INFINITY {
        return "-inf".to_owned();
    }
    if value == 0.0 {
        return if value.is_sign_negative() {
            "-0.0".to_owned()
        } else {
            "0.0".to_owned()
        };
    }

    let raw = value.abs().to_string();
    let (mantissa, explicit_exponent) = raw
        .split_once(['e', 'E'])
        .map(|(mantissa, exponent)| (mantissa, exponent.parse::<i32>().unwrap_or(0)))
        .unwrap_or((&raw, 0));
    let decimal_places = mantissa.find('.').unwrap_or(mantissa.len()) as i32;
    let mut digits = mantissa
        .bytes()
        .filter(u8::is_ascii_digit)
        .map(char::from)
        .collect::<String>();
    let leading_zeroes = digits.bytes().take_while(|digit| *digit == b'0').count();
    if leading_zeroes == digits.len() {
        return if value.is_sign_negative() {
            "-0.0".to_owned()
        } else {
            "0.0".to_owned()
        };
    }
    digits.drain(..leading_zeroes);
    while digits.len() > 1 && digits.ends_with('0') {
        digits.pop();
    }
    let decimal_exponent = decimal_places + explicit_exponent - leading_zeroes as i32 - 1;
    let sign = if value.is_sign_negative() { "-" } else { "" };

    if decimal_exponent < -4 || decimal_exponent >= 16 {
        let mut scientific = digits[..1].to_owned();
        if digits.len() > 1 {
            scientific.push('.');
            scientific.push_str(&digits[1..]);
        }
        let exponent = if decimal_exponent >= 0 {
            format!("+{decimal_exponent:02}")
        } else {
            format!("-{abs:02}", abs = decimal_exponent.unsigned_abs())
        };
        format!("{sign}{scientific}e{exponent}")
    } else {
        let point = decimal_exponent + 1;
        let plain = if point <= 0 {
            format!("0.{}{}", "0".repeat((-point) as usize), digits)
        } else if point as usize >= digits.len() {
            format!("{}{}", digits, "0".repeat(point as usize - digits.len()))
        } else {
            format!(
                "{}.{}",
                &digits[..point as usize],
                &digits[point as usize..]
            )
        };
        let suffix = if plain.contains('.') { "" } else { ".0" };
        format!("{sign}{plain}{suffix}")
    }
}

fn digits_string(digits: &[u8]) -> String {
    digits
        .iter()
        .map(|digit| char::from(b'0' + *digit))
        .collect()
}

fn normalized_digits(digits: &[u8]) -> Vec<u8> {
    let first = digits
        .iter()
        .position(|digit| *digit != 0)
        .unwrap_or(digits.len().saturating_sub(1));
    digits[first..].to_vec()
}

fn padded_digits(digits: &[u8], length: usize) -> Vec<u8> {
    let mut result = digits.to_vec();
    result.resize(length, 0);
    result
}

fn compare_abs(left: &[u8], right: &[u8]) -> Ordering {
    let left = normalized_digits(left);
    let right = normalized_digits(right);
    left.len().cmp(&right.len()).then_with(|| left.cmp(&right))
}

fn add_abs(left: &[u8], right: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(left.len().max(right.len()) + 1);
    let (mut i, mut j, mut carry) = (left.len(), right.len(), 0_u8);
    while i > 0 || j > 0 || carry > 0 {
        let a = if i > 0 {
            i -= 1;
            left[i]
        } else {
            0
        };
        let b = if j > 0 {
            j -= 1;
            right[j]
        } else {
            0
        };
        let sum = a + b + carry;
        output.push(sum % 10);
        carry = sum / 10;
    }
    output.reverse();
    output
}

fn subtract_abs(left: &[u8], right: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(left.len());
    let (mut i, mut j, mut borrow) = (left.len(), right.len(), 0_i16);
    while i > 0 {
        i -= 1;
        let a = i16::from(left[i]) - borrow;
        let b = if j > 0 {
            j -= 1;
            i16::from(right[j])
        } else {
            0
        };
        if a < b {
            output.push((a + 10 - b) as u8);
            borrow = 1;
        } else {
            output.push((a - b) as u8);
            borrow = 0;
        }
    }
    output.reverse();
    let first = output
        .iter()
        .position(|digit| *digit != 0)
        .unwrap_or(output.len().saturating_sub(1));
    output.drain(..first);
    if output.is_empty() {
        output.push(0);
    }
    output
}

fn multiply_abs(left: &[u8], right: &[u8]) -> Vec<u8> {
    let mut result = vec![0_u32; left.len() + right.len()];
    for (left_index, left_digit) in left.iter().enumerate() {
        for (right_index, right_digit) in right.iter().enumerate() {
            result[left_index + right_index + 1] +=
                u32::from(*left_digit) * u32::from(*right_digit);
        }
    }
    for index in (1..result.len()).rev() {
        let carry = result[index] / 10;
        result[index] %= 10;
        result[index - 1] += carry;
    }
    let digits = result
        .into_iter()
        .map(|digit| digit as u8)
        .collect::<Vec<_>>();
    normalized_digits(&digits)
}

fn multiply_small(digits: &[u8], multiplier: u8) -> Vec<u8> {
    let mut output = Vec::with_capacity(digits.len() + 2);
    let mut carry = 0_u16;
    for digit in digits.iter().rev() {
        let product = u16::from(*digit) * u16::from(multiplier) + carry;
        output.push((product % 10) as u8);
        carry = product / 10;
    }
    while carry > 0 {
        output.push((carry % 10) as u8);
        carry /= 10;
    }
    output.reverse();
    normalized_digits(&output)
}

fn increment_digits(digits: &mut Vec<u8>) {
    let mut index = digits.len();
    while index > 0 {
        index -= 1;
        if digits[index] < 9 {
            digits[index] += 1;
            return;
        }
        digits[index] = 0;
    }
    digits.insert(0, 1);
}

fn divide_abs(numerator: &[u8], denominator: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let mut quotient = Vec::with_capacity(numerator.len());
    let mut remainder = vec![0];
    for digit in numerator {
        if remainder.len() == 1 && remainder[0] == 0 {
            remainder[0] = *digit;
        } else {
            remainder.push(*digit);
        }
        let mut quotient_digit = 0_u8;
        while compare_abs(&remainder, denominator) != Ordering::Less {
            remainder = subtract_abs(&remainder, denominator);
            quotient_digit += 1;
        }
        quotient.push(quotient_digit);
    }
    (normalized_digits(&quotient), normalized_digits(&remainder))
}

#[cfg(test)]
mod tests {
    use super::evaluate;

    #[test]
    fn decimal_zero_comparisons_cover_both_signs_and_orders() {
        for expression in [
            "0.1 > 0",
            "0 < 0.1",
            "-0.1 < 0",
            "0 > -0.1",
            "0.000001 > -0",
            "-0 = 0",
            "0e100 = 0e-100",
            "1e-9223372036854775808 > 0",
            "-1e-9223372036854775808 < 0",
            "1e9223372036854775807 > 1",
            "1 > 1e-9223372036854775808",
        ] {
            assert!(evaluate(expression).unwrap(), "{expression}");
        }
        for expression in ["0.1 <= 0", "0 >= 0.1", "-0.1 >= 0", "0 < -0.1"] {
            assert!(!evaluate(expression).unwrap(), "{expression}");
        }
    }

    #[test]
    fn extreme_exponent_arithmetic_returns_errors_without_overflow() {
        for expression in [
            "1e9223372036854775807 + 1e-9223372036854775808",
            "1e-9223372036854775808 - 1e9223372036854775807",
            "1e9223372036854775807 * 1e1",
            "1e-9223372036854775808 / 10",
            "1e9223372036854775807 / 1e-9223372036854775808",
            "999999999999999999999999999999999999999999999999999e9223372036854775807 + 0e9223372036854775807",
        ] {
            assert!(evaluate(expression).is_err(), "{expression}");
        }
        assert!(evaluate("1e100 + 1e100 = 2e100").unwrap());
        assert!(evaluate("1e9223372036854775807 * 10 > 1e9223372036854775807").unwrap());
        assert!(evaluate("1e-9223372036854775808 / 1 = 1e-9223372036854775808").unwrap());
        assert!(evaluate("1e9223372036854775807 / 1 = 1e9223372036854775807").unwrap());
        assert!(evaluate("1e-100 / 10 = 1e-101").unwrap());
        assert!(evaluate("True or 1e-9223372036854775808 / 10").unwrap());
    }

    #[test]
    fn boolean_operators_short_circuit_runtime_errors() {
        for expression in [
            "False and 1 / 0",
            "0 != 0 and 10 / 0 > 0.1",
            "False and missing(1/0)",
        ] {
            assert!(!evaluate(expression).unwrap(), "{expression}");
        }
        for expression in ["True or 1 / 0", "True or missing(1/0)", "1 or 0 ** -1"] {
            assert!(evaluate(expression).unwrap(), "{expression}");
        }
        assert!(evaluate("True and 1 / 0").is_err());
        assert!(evaluate("False or 1 / 0").is_err());
        // Function arguments remain eager; only language boolean operators short circuit.
        assert!(evaluate("AND(False, 1 / 0)").is_err());
    }

    #[test]
    fn short_circuit_chains_and_nested_branches_restore_evaluation() {
        for expression in [
            "False and 1/0 and 2/0 or True",
            "True or 1/0 or 2/0",
            "(False and (True or 1/0)) or (True and not False)",
            "False and 1/0 or 2 == 2",
            "(True or 1/0) and 2 == 2",
        ] {
            assert!(evaluate(expression).unwrap(), "{expression}");
        }
        assert!(!evaluate("3 < 2 < 1/0").unwrap());
        assert!(!evaluate("True and False and 1/0").unwrap());
        assert!(evaluate("(False and 1/0) or 1/0").is_err());
    }

    #[test]
    fn skipped_branches_still_require_valid_syntax() {
        for expression in [
            "False and",
            "True or (1 +)",
            "False and (1/0",
            "True or f(,)",
            "True or (1 2)",
            "False and @",
        ] {
            assert!(evaluate(expression).is_err(), "{expression}");
        }
    }

    #[test]
    fn power_exponents_are_bounded_before_decimal_expansion() {
        for expression in [
            "2 ** 1e100001",
            "2 ** 1e9223372036854775807",
            "2 ** 1e-9223372036854775807",
            "2 ** 10001",
            "2 ** -10001",
        ] {
            assert!(evaluate(expression).is_err(), "{expression}");
        }
        assert!(evaluate("1 ** 10000 = 1").unwrap());
        assert!(evaluate("1 ** -10000 = 1").unwrap());
        assert!(evaluate("2 ** 2 = 4").unwrap());
        assert!(evaluate("2 ** 2.0 = 4").unwrap());
        assert!(evaluate("2 ** 2.5").is_err());
        assert!(evaluate("True or 2 ** 1e100001").unwrap());
        assert!(evaluate("-2 ** 2 = -4").unwrap());
        assert!(evaluate("(-2) ** 2 = 4").unwrap());
    }

    #[test]
    fn decimal_balance_and_comparison_aliases_match_python_expression_semantics() {
        assert!(!evaluate("9697998.71 <> 8875290.77 + 822707.94").unwrap());
        assert!(evaluate("9697998.71 = 8875290.77 + 822707.94").unwrap());
        assert!(evaluate("100 < 0.055 * 2000").unwrap());
    }

    #[test]
    fn boolean_helpers_and_python_truthiness_are_supported() {
        assert!(evaluate("AND(1, OR(0, 2))").unwrap());
        assert!(!evaluate("NOT(2 < 3)").unwrap());
        assert!(evaluate("True = 1").unwrap());
        assert!(evaluate("0.1 + 0.2 = 0.3").unwrap());
    }

    #[test]
    fn decimal_context_rounding_uses_fifty_digits_and_half_even() {
        let value = evaluate("1 / 3 * 3 = 1").unwrap();
        assert!(!value);
        assert!(evaluate("1.005 + 0 = 1.005").unwrap());
    }
}
