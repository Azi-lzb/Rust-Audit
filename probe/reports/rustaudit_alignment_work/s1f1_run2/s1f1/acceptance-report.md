# Rust-Audit 业务验收报告

- 状态：failed
- 运行编号：20261107_155233_663
- 开始：2026-11-07T15:52:33.662717400+08:00
- 结束：2026-11-07T15:56:21.943646+08:00
- 耗时：228280 ms
- 输出目录：`\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1`

## 业务步骤

| 步骤 | 状态 | 退出码 | 耗时 | 输入未变 | 输出文件数 |
|---|---|---:|---:|---|---:|
| S1-F01 模板审核 | failed | 2 | 228162 ms | true | 130 |

错误原文：

```text
S1-F01 部分完成：7 个待人工重算、0 个部分审核、0 个失败；总用时 226.8 秒；已生成结果 \\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\汇总核查表校验.xlsx
```

关键计数：`{"conditionalIssues":244,"filesTotal":129,"formulaIssues":3345,"matched":true,"skipped":0,"structurePassed":129,"totalIssues":3589}`

## 输入文件哈希

| 输入文件 | SHA-256 未变 |
|---|---|
| `\\?\C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\config-real\外部文件-大集中相关数据.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！10.基础数据-单位贷款202609.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！100.金融基础数据采集（单位存款）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！110.金融基础数据-存量个人存款202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！120.金融基础数据采集（代发工资）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！130.互联网金融业务信息（季度）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！130.互联网金融业务信息（月度）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！140.养老产业贷款信息202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！140.数字经济产业贷款信息202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！140.普惠贷款信息202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！140.科技贷款信息202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！140.绿色贷款信息202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！20.融基础数据-个人贷款202609.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！30.存量专项贷款一批报文202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！40.金融基础数据采集（委托贷款）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！50.金融基础数据-票据业务202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！60.金融基础数据采集（单位贷款置换旧债发生额）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！61.存量专项贷款二批报文202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！70.金融基础数据采集（债券业务）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！80.金融基础数据采集（股权和特定目的载体投资）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\business_templates\！90.金融基础数据采集（同业业务）202608.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-history-with-external.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-manifest-with-external.json` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\TCL科技集团财务有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（代发工资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东博罗农村商业银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（代发工资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东惠东农村商业银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（代发工资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\广东龙门农村商业银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠东惠民村镇银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（代发工资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州农村商业银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-互联网金融业务信息（月度）_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（债券业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（单位存款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（同业业务）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（委托贷款）_2026-08-31_数据核对表.xlsx` | true |
| `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\逐笔统计系统\核查表\惠州博罗长江村镇银行股份有限公司_金融基础数据采集（股权和特定目的载体投资）_2026-08-31_数据核对表.xlsx` | true |

## 生成文件

- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0001-TCL科技集团财务有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0002-TCL科技集团财务有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0003-TCL科技集团财务有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0004-TCL科技集团财务有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0005-TCL科技集团财务有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0006-TCL科技集团财务有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0007-TCL科技集团财务有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0008-TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0009-TCL科技集团财务有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0010-TCL科技集团财务有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0011-TCL科技集团财务有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0012-TCL科技集团财务有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0013-TCL科技集团财务有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0014-TCL科技集团财务有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0015-TCL科技集团财务有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0016-TCL科技集团财务有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0017-TCL科技集团财务有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0018-广东博罗农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0019-广东博罗农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0020-广东博罗农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0021-广东博罗农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0022-广东博罗农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0023-广东博罗农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0024-广东博罗农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0025-广东博罗农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0026-广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0027-广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0028-广东博罗农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0029-广东博罗农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0030-广东博罗农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0031-广东博罗农村商业银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0032-广东博罗农村商业银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0033-广东博罗农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0034-广东博罗农村商业银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0035-广东博罗农村商业银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0036-广东博罗农村商业银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0037-广东惠东农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0038-广东惠东农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0039-广东惠东农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0040-广东惠东农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0041-广东惠东农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0042-广东惠东农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0043-广东惠东农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0044-广东惠东农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0045-广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0046-广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0047-广东惠东农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0048-广东惠东农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0049-广东惠东农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0050-广东惠东农村商业银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0051-广东惠东农村商业银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0052-广东惠东农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0053-广东惠东农村商业银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0054-广东惠东农村商业银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0055-广东惠东农村商业银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0056-广东龙门农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0057-广东龙门农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0058-广东龙门农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0059-广东龙门农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0060-广东龙门农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0061-广东龙门农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0062-广东龙门农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0063-广东龙门农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0064-广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0065-广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0066-广东龙门农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0067-广东龙门农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0068-广东龙门农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0069-广东龙门农村商业银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0070-广东龙门农村商业银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0071-广东龙门农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0072-广东龙门农村商业银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0073-广东龙门农村商业银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0074-广东龙门农村商业银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0075-惠东惠民村镇银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0076-惠东惠民村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0077-惠东惠民村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0078-惠东惠民村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0079-惠东惠民村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0080-惠东惠民村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0081-惠东惠民村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0082-惠东惠民村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0083-惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0084-惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0085-惠东惠民村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0086-惠东惠民村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0087-惠东惠民村镇银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0088-惠东惠民村镇银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0089-惠东惠民村镇银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0090-惠东惠民村镇银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0091-惠东惠民村镇银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0092-惠东惠民村镇银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0093-惠州农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0094-惠州农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0095-惠州农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0096-惠州农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0097-惠州农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0098-惠州农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0099-惠州农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0100-惠州农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0101-惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0102-惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0103-惠州农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0104-惠州农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0105-惠州农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0106-惠州农村商业银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0107-惠州农村商业银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0108-惠州农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0109-惠州农村商业银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0110-惠州农村商业银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0111-惠州农村商业银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0112-惠州博罗长江村镇银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0113-惠州博罗长江村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0114-惠州博罗长江村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0115-惠州博罗长江村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0116-惠州博罗长江村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0117-惠州博罗长江村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0118-惠州博罗长江村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0119-惠州博罗长江村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0120-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0121-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0122-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0123-惠州博罗长江村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0124-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_债券业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0125-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_单位存款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0126-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0127-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_同业业务__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0128-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_委托贷款__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\0129-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_股权和特定目的载体投资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\rustaudit_alignment_work\s1f1_run2\s1f1\steps\S1-F01\S1-F1_审核结果_20261107_155234_115\汇总核查表校验.xlsx`
