//! 常见问题（帮助与更新 → 常见问题）。
//!
//! 内容沿用旧 audit `docs/常见问题.json` 的 23 条问答（用户确认「之前旧 audit
//! 的常见问题 你也可以放进 AuditTools 里」）。S1S2 发行版取 `answerS1S2`，
//! 全量版取 `answer`；没有 S1S2 专用答案时回退到完整答案。

use serde_json::Value;

const FAQ_JSON: &str = include_str!("faq_data.json");

/// 返回 {items:[{category,question,answer}]}，按发行档位选择答案文本。
pub fn items(profile: &str) -> Value {
    let parsed: Value = serde_json::from_str(FAQ_JSON).unwrap_or(Value::Array(Vec::new()));
    let s1s2 = profile == "s1s2";
    let items = parsed
        .as_array()
        .map(Vec::as_slice)
        .unwrap_or_default()
        .iter()
        .filter_map(|entry| {
            let category = entry.get("category")?.as_str()?.to_owned();
            let question = entry.get("question")?.as_str()?.to_owned();
            let full = entry.get("answer").and_then(Value::as_str);
            let s1s2_answer = entry.get("answerS1S2").and_then(Value::as_str);
            let answer = if s1s2 {
                s1s2_answer.or(full)
            } else {
                full.or(s1s2_answer)
            }?
            .to_owned();
            Some(serde_json::json!({
                "category": category,
                "question": question,
                "answer": answer,
            }))
        })
        .collect::<Vec<_>>();
    serde_json::json!({"items": items})
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bundled_faq_keeps_every_question_for_both_editions() {
        for profile in ["s1s2", "s1s2s3"] {
            let value = items(profile);
            let items = value["items"].as_array().expect("items array");
            assert_eq!(items.len(), 23, "{profile} 应包含全部 23 条问答");
            for item in items {
                assert!(!item["category"].as_str().unwrap_or_default().is_empty());
                assert!(!item["question"].as_str().unwrap_or_default().is_empty());
                assert!(!item["answer"].as_str().unwrap_or_default().is_empty());
            }
        }
    }

    #[test]
    fn s1s2_edition_prefers_the_shorter_answer() {
        let full = items("s1s2s3");
        let compact = items("s1s2");
        let first = &full["items"][0];
        assert_eq!(first["category"], compact["items"][0]["category"]);
        assert_eq!(first["question"], compact["items"][0]["question"]);
        assert_ne!(
            first["answer"], compact["items"][0]["answer"],
            "首条问答有独立 S1S2 版本，两档答案应不同"
        );
    }
}
