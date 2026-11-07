# Rust-Audit 业务验收报告

- 状态：failed
- 运行编号：20261105_233824_508
- 开始：2026-11-05T23:38:24.506991800+08:00
- 结束：2026-11-05T23:41:55.967186500+08:00
- 耗时：211460 ms
- 输出目录：`\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external`

## 业务步骤

| 步骤 | 状态 | 退出码 | 耗时 | 输入未变 | 输出文件数 |
|---|---|---:|---:|---|---:|
| S1-F01 模板审核 | failed | 2 | 211327 ms | true | 95 |

错误原文：

```text
S1-F01 审核完成：公式异常 3224 条、条件格式异常 59 条；总耗时 209.9 秒；结果已保存
```

关键计数：`{"conditionalIssues":59,"filesTotal":129,"formulaIssues":3224,"matched":true,"skipped":0,"structurePassed":129,"totalIssues":3283}`

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

- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0001-TCL科技集团财务有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0002-TCL科技集团财务有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0003-TCL科技集团财务有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0004-TCL科技集团财务有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0005-TCL科技集团财务有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0006-TCL科技集团财务有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0007-TCL科技集团财务有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0008-TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0009-TCL科技集团财务有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0010-TCL科技集团财务有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0011-TCL科技集团财务有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0014-TCL科技集团财务有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0018-广东博罗农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0019-广东博罗农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0020-广东博罗农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0021-广东博罗农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0022-广东博罗农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0023-广东博罗农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0024-广东博罗农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0025-广东博罗农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0026-广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0027-广东博罗农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0028-广东博罗农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0029-广东博罗农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0030-广东博罗农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0033-广东博罗农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0037-广东惠东农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0038-广东惠东农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0039-广东惠东农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0040-广东惠东农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0041-广东惠东农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0042-广东惠东农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0043-广东惠东农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0044-广东惠东农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0045-广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0046-广东惠东农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0047-广东惠东农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0048-广东惠东农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0049-广东惠东农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0052-广东惠东农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0056-广东龙门农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0057-广东龙门农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0058-广东龙门农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0059-广东龙门农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0060-广东龙门农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0061-广东龙门农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0062-广东龙门农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0063-广东龙门农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0064-广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0065-广东龙门农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0066-广东龙门农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0067-广东龙门农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0068-广东龙门农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0071-广东龙门农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0075-惠东惠民村镇银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0076-惠东惠民村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0077-惠东惠民村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0078-惠东惠民村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0079-惠东惠民村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0080-惠东惠民村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0081-惠东惠民村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0082-惠东惠民村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0083-惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0084-惠东惠民村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0085-惠东惠民村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0086-惠东惠民村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0089-惠东惠民村镇银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0093-惠州农村商业银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0094-惠州农村商业银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0095-惠州农村商业银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0096-惠州农村商业银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0097-惠州农村商业银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0098-惠州农村商业银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0099-惠州农村商业银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0100-惠州农村商业银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0101-惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0102-惠州农村商业银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0103-惠州农村商业银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0104-惠州农村商业银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0105-惠州农村商业银行股份有限公司_金融基础数据采集_代发工资__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0108-惠州农村商业银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0112-惠州博罗长江村镇银行股份有限公司_五篇大文章-互联网金融业务信息_月度__B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0113-惠州博罗长江村镇银行股份有限公司_五篇大文章-养老产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0114-惠州博罗长江村镇银行股份有限公司_五篇大文章-数字经济产业贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0115-惠州博罗长江村镇银行股份有限公司_五篇大文章-普惠贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0116-惠州博罗长江村镇银行股份有限公司_五篇大文章-科技贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0117-惠州博罗长江村镇银行股份有限公司_五篇大文章-绿色贷款信息_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0118-惠州博罗长江村镇银行股份有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0119-惠州博罗长江村镇银行股份有限公司_金融基础数据-单位贷款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0120-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款一批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0121-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量专项贷款二批报文_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0122-惠州博罗长江村镇银行股份有限公司_金融基础数据-存量个人存款_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0123-惠州博罗长江村镇银行股份有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\0126-惠州博罗长江村镇银行股份有限公司_金融基础数据采集_单位贷款置换旧债发生额__2026-08-31_数据核对表-recalculated.xlsx`
- `\\?\C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据\S1F1_新旧对比_20261105_231620\rust-with-external\steps\S1-F01\S1-F1_审核结果_20261105_233824_589\汇总核查表校验.xlsx`
