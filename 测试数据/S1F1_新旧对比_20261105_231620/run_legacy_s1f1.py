import json, sys, time, traceback
from pathlib import Path
from dataclasses import is_dataclass, asdict
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
manifest_path=Path(sys.argv[1])
manifest=json.loads(manifest_path.read_text(encoding="utf-8-sig"))
sys.path.insert(0, r"C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core\src")
from base_audit.service import AuditService
from base_audit.workflow.runner import run_dag_flow
results=[]
for item in manifest["batches"]:
    started=time.perf_counter()
    output=Path(item["output_dir"])
    output.mkdir(parents=True, exist_ok=False)
    events=[]
    record={"name":item["name"],"started_epoch":time.time()}
    try:
        service=AuditService(config_path=Path(item["history_path"]), engine_preference="\u81ea\u52a8", conditional_format_evaluator="PYTHON", formula_region_writer="DIRECT_OOXML", external_sheet_writer="DIRECT_OOXML")
        result=run_dag_flow(item["workflow_id"],service=service,template_path=Path(item["template_path"]),input_dir=Path(item["input_dir"]),output_dir=output,period=item["period"],history_path=Path(item["history_path"]),selected_files=[Path(x) for x in item["selected_files"]],external_path=Path(item["external_path"]) if item.get("external_path") else None,recursive=False,on_step=events.append,write_flow_logs=True,strict=True)
        record.update({"status":"completed","result_type":type(result).__name__,"summary":result.summary_text(),"result_fields":vars(result) if hasattr(result,"__dict__") else str(result),"events":events})
    except Exception as e:
        record.update({"status":"failed","error":str(e),"traceback":traceback.format_exc(),"events":events})
    finally:
        record["elapsed_ms"]=round((time.perf_counter()-started)*1000)
    results.append(record)
    manifest_path.with_name("python-results.json").write_text(json.dumps(results,ensure_ascii=False,indent=2,default=str),encoding="utf-8")
    print('BATCH_DONE', len(results), record['status'], record['elapsed_ms'], 'ms', flush=True)
print(json.dumps([{k:r.get(k) for k in ("name","status","elapsed_ms","summary","error")} for r in results],ensure_ascii=False,indent=2,default=str))
if any(r["status"]!="completed" for r in results): sys.exit(2)