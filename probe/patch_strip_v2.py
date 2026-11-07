import io

p = 'audit-core/src/excel_open_repair.rs'
s = io.open(p, encoding='utf-8').read()

q = chr(39)
start = s.index('/// 剥离公式中指向本地已存在表的')
end = s.index('fn quote_bare_refs(')
new_block = r'''/// 表名的"编码无关"正则片段：每个非 ASCII 字符同时匹配字面、十进制 NCR 与
/// 十六进制 NCR。openpyxl 保存时把中文写成 `&#38598;` 实体，WPS/Excel 生成
/// 的文件是字面 UTF-8；两种都必须命中，否则经 openpyxl 处理过的副本会漏修。
fn name_flex_regex(name: &str) -> String {
    let mut parts = String::new();
    for ch in name.chars() {
        if (ch as u32) < 128 {
            parts.push_str(&regex::escape(&ch.to_string()));
        } else {
            let code = ch as u32;
            parts.push_str(&format!(
                "(?:{}|&#{code};|&#[xX]{:x};|&#[xX]{:X};)",
                regex::escape(&ch.to_string()),
                code,
                code
            ));
        }
    }
    parts
}

/// 剥离公式里的外部工作簿引用前缀（业务口径：副本只允许工作簿内引用）。
///
/// 覆盖四种形态（与 Python `_strip_external_prefixes` 逐条一致）：
/// ① `'C:\路径\[簿.xlsx]表名'!` ② `'[簿.xlsx]表名'!`
/// ③ `[簿.xlsx]表名!` ④ `[1]表名!`（或 `[1]'表名'!`）
///
/// 索引式与路径式**一律无条件剥离**（不看本地是否已有同名表）：本地有同名表
/// 则公式改指本地表；本地没有则公式自然成 `#REF!`，按审核异常计入——这优于
/// 整簿被 Excel 拒绝打开。`local_sheets` 仅用于日志判定缺失表名。
///
/// 返回 (新文本, 剥离次数)。
fn strip_external_refs(text: &str, local_sheets: &[String], stripped: &mut usize) -> String {
    let mut result = text.to_owned();
    let mut total = 0usize;

    // ① ② 路径式 + 引号：'...\[簿.xlsx]表名'! → '表名'!
    if let Ok(re) = Regex::new(r"'[^']*\[[^\]/]+\.(?:xlsx|xlsm|xlsb|xls)\]((?:[^']|'')*)'!") {
        let replaced = re.replace_all(&result, |caps: &Captures| {
            format!("'{}'!", &caps[1])
        });
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // ③ 路径式 + 裸名：[簿.xlsx]表名! → 表名!
    if let Ok(re) = Regex::new(r"\[[^\]/]+\.(?:xlsx|xlsm|xlsb|xls)\]([^\s!'\[\]<>()]+)!") {
        let replaced = re.replace_all(&result, |caps: &Captures| {
            format!("{}!", &caps[1])
        });
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // ④ 索引式 + 引号：[1]'表名'! → '表名'!
    if let Ok(re) = Regex::new(r"\[\d+\]'((?:[^']|'')*)'!") {
        let replaced = re.replace_all(&result, |caps: &Captures| {
            format!("'{}'!", &caps[1])
        });
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // ④ 索引式 + 裸名：[1]表名! → 表名!
    if let Ok(re) = Regex::new(r"\[\d+\]([^\s!'\[\]<>()]+)!") {
        let replaced = re.replace_all(&result, |caps: &Captures| {
            format!("{}!", &caps[1])
        });
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // 索引式 + NCR 编码表名：[1]&#38598;...! → 表名!
    for name in local_sheets {
        let flex = name_flex_regex(name);
        let escaped = name_flex_regex(&name.replace('\'', "''"));
        if let Ok(re) = Regex::new(&format!(r"\[\d+\]'{escaped}'!")) {
            let replacement = format!("'{}'!", name.replace('\'', "''"));
            let replaced = re.replace_all(&result, replacement.as_str());
            if replaced != result {
                total += re.find_iter(&result).count();
                result = replaced.into_owned();
            }
        }
        if let Ok(re) = Regex::new(&format!(r"\[\d+\]{flex}!")) {
            let replacement = format!("{name}!");
            let replaced = re.replace_all(&result, replacement.as_str());
            if replaced != result {
                total += re.find_iter(&result).count();
                result = replaced.into_owned();
            }
        }
    }

    *stripped += total;
    result
}

/// 是否仍存在任何形态的外部引用（索引式或路径式）。用于决定 externalLink
/// 定义能否安全清除——**检测与剥离必须同源**，只查 `[N]` 会漏判路径式，
/// 误删外链定义后原本存活的引用变成新的悬空引用。
fn has_any_external_ref(text: &str) -> bool {
    if Regex::new(r"\[\d+\]")
        .ok()
        .is_some_and(|re| re.is_match(text))
    {
        return true;
    }
    Regex::new(r"\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls)\]")
        .ok()
        .is_some_and(|re| re.is_match(text))
}

/// 剥离后仍指向"本地不存在表"的引用名（用于运行日志提示，不阻断）。
/// 只识别合法表引用：`'表名'!` 或中文/字母/数字/下划线/空格组成的 `表名!`。
fn dangling_after_strip(formula_text: &str, local_sheets: &[String]) -> Vec<String> {
    let mut missing: Vec<String> = Vec::new();
    let Ok(re) = Regex::new(r"(?:'((?:[^']|'')*)'|([A-Za-z_\u{4e00}-\u{9fff}][\w\u{4e00}-\u{9fff} .\u{3000}]*))!")
    else {
        return missing;
    };
    for caps in re.captures_iter(formula_text) {
        let name = caps
            .get(1)
            .or_else(|| caps.get(2))
            .map(|m| m.as_str().replace("''", "'"))
            .unwrap_or_default();
        let name = name.trim();
        if name.is_empty() || local_sheets.iter().any(|sheet| sheet == name) {
            continue;
        }
        // 排除单元格地址（如 A1、AB12）
        if Regex::new(r"^[A-Za-z]{1,3}\d+$")
            .ok()
            .is_some_and(|re| re.is_match(name))
        {
            continue;
        }
        if !missing.iter().any(|existing| existing == name) {
            missing.push(name.to_owned());
        }
    }
    missing
}

'''
s = s[:start] + new_block + s[end:]
io.open(p, 'w', encoding='utf-8', newline='').write(s)
print('strip rewritten')
