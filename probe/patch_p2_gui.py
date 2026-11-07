import io

p = 'audit-core/frontend/web/index.html'
s = io.open(p, encoding='utf-8').read()

# 1) tab button
old = '<button type="button" class="settings-tab-btn" data-tab="updates" onclick="setSettingsTab(\'updates\')">帮助与更新</button>'
new = (old + '\n        <button type="button" class="settings-tab-btn" data-tab="combine" onclick="setSettingsTab(\'combine\')">组合设置</button>')
assert s.count(old) == 1
s = s.replace(old, new)

# 2) setSettingsTab allow combine + dispatch
old = "  if(tab!=='general'&&tab!=='guide'&&tab!=='updates'&&tab!=='advanced')tab='general';"
new = "  if(tab!=='general'&&tab!=='guide'&&tab!=='updates'&&tab!=='combine'&&tab!=='advanced')tab='general';"
assert s.count(old) == 1
s = s.replace(old, new)
old = "  else if(tab==='updates')renderUpdates();"
new = "  else if(tab==='updates')renderUpdates();\n  else if(tab==='combine')renderCombine();"
assert s.count(old) == 1
s = s.replace(old, new)

# 3) renderCombine function (before renderUpdates)
old = "async function renderUpdates(){"
new = """function renderCombine(){
  const pane=document.getElementById('settingsPane');
  /* 组合工作表分组方案在 Rust 后端的 get_dag_editor_data 接口落地前以空态
     呈现；方案的编辑入口仍在“流程编排”的“组合工作表”业务参数中。 */
  pane.innerHTML='<div class="advanced-zone">'
    +'<div class="inspector-empty" style="margin-bottom:12px">组合工作表分组方案列表随“流程编排”数据接口一并提供；当前版本请在“流程编排”页选择“组合工作表”模块编辑方案。组合联合核查表固定按文件名下划线第一段分组，无需方案。</div>'
    +'<div id="combinePane">'+combineSheetsHtml()+'</div>'
    +'</div>';
}

async function renderUpdates(){"""
assert s.count(old) == 1
s = s.replace(old, new)

# 4) updates config subtab gate
old = "  const subTabs='<div class=\"updates-subtabs\">'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='log'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'log\\')\">更新日志</button>'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='faq'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'faq\\')\">常见问题</button>'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='libreoffice'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'libreoffice\\')\">安装 LibreOffice</button>'\n    +'</div>';"
new = "  const configSupported=!!(state&&state.configUpdateSupported);\n  if(!configSupported&&UPDATES_SUB==='config')UPDATES_SUB='log';\n  const subTabs='<div class=\"updates-subtabs\">'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='log'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'log\\')\">更新日志</button>'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='faq'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'faq\\')\">常见问题</button>'\n    +'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='libreoffice'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'libreoffice\\')\">安装 LibreOffice</button>'\n    +(configSupported?'<button type=\"button\" class=\"btn tiny'+(UPDATES_SUB==='config'?' primary':'')+'\" onclick=\"setUpdatesSub(\\'config\\')\">更新配置</button>':'')\n    +'</div>';"
assert s.count(old) == 1, 'subtabs anchor'
s = s.replace(old, new)

# 5) config page branch
old = "    if(UPDATES_SUB==='libreoffice'){\n      renderLibreOfficeInstall(LO_INSTALL);\n      await refreshLibreOfficeInstall(true);\n      return;\n    }"
new = "    if(UPDATES_SUB==='libreoffice'){\n      renderLibreOfficeInstall(LO_INSTALL);\n      await refreshLibreOfficeInstall(true);\n      return;\n    }\n    if(UPDATES_SUB==='config'){\n      const report=await bridge.api.get_config_update();\n      if(!report||report.applicable!==true){\n        body.className='inspector-empty';\n        body.textContent=(report&&report.reason)||'当前环境不适用配置更新。';\n        return;\n      }\n      const entries=report.entries||[];\n      if(!entries.length){\n        body.className='inspector-empty';\n        body.textContent='随包配置与清单一致，没有可更新项。';\n        return;\n      }\n      body.className='';\n      body.innerHTML='<div class=\"guide-flow-body\"><p class=\"module-help-desc\">勾选需要更新的配置；更新前自动备份到配置目录“备份-配置更新-时间戳”，逐文件校验哈希，失败即停止。更新完成后需重启程序生效。</p>'\n        +entries.map(function(entry){\n          const statusLabel={update:'可更新',missing:'本机缺失',same:'一致',unknown:'版本无法确认',local_newer:'本机较新'}[entry.status]||entry.status;\n          const disabled=!entry.available;\n          return '<label class=\"combine-plan-row\" style=\"cursor:'+(disabled?'default':'pointer')+'\"><input type=\"checkbox\" class=\"cfg-update-item\" value=\"'+esc(entry.name)+'\" '+(disabled?'disabled':'')+' '+(entry.status==='update'?'checked':'')+'><span><b>'+esc(entry.title||entry.name)+'</b>　<small>'+esc(statusLabel)+'　随包 v'+esc(entry.bundled_version||'')+(entry.local_version?('　本机 v'+esc(entry.local_version)):'')+'</small></span></label>';\n        }).join('')\n        +'<div style=\"margin-top:12px;display:flex;gap:8px\"><button type=\"button\" class=\"btn primary\" onclick=\"applyConfigUpdate()\">更新所选配置</button></div></div>';\n      return;\n    }"
assert s.count(old) == 1, 'libreoffice branch anchor'
s = s.replace(old, new)

# 6) applyConfigUpdate handler
old = "function libreOfficeInstallVisible(){return currentPage==='settings'&&CFG.tab==='updates'&&UPDATES_SUB==='libreoffice'}"
new = ("function libreOfficeInstallVisible(){return currentPage==='settings'&&CFG.tab==='updates'&&UPDATES_SUB==='libreoffice'}\n"
       "async function applyConfigUpdate(){\n"
       "  const names=[].slice.call(document.querySelectorAll('.cfg-update-item:checked')).map(function(el){return el.value;});\n"
       "  if(!names.length){showToast('请先勾选要更新的配置文件');return;}\n"
       "  if(state&&state.busy)return;\n"
       "  try{\n"
       "    const result=await bridge.api.apply_config_update(names);\n"
       "    if(!result||result.ok!==true){showToast((result&&result.message)||'配置更新失败');return;}\n"
       "    showToast((result.message||'配置已更新')+'；请重启程序使新配置生效。');\n"
       "  }catch(e){showToast(e&&e.message?e.message:String(e));}\n"
       "  await renderUpdates();\n"
       "}")
assert s.count(old) == 1
s = s.replace(old, new)

# 7) evaluator select in advanced
old = "    row('空白格命中计入结果','开启时空白单元格实际触发（含实际显示变红）同样计入 S1-F01 异常，与 Excel/WPS COM 渲染口径一致；关闭后排除空白格命中，只保留有值的触发。',\n      toggle('setConditionalFormatIncludeBlankCells',s.conditionalFormatIncludeBlankCells!==false)),\n  ].join('');"
new = ("    row('空白格命中计入结果','开启时空白单元格实际触发（含实际显示变红）同样计入 S1-F01 异常，与 Excel/WPS COM 渲染口径一致；关闭后排除空白格命中，只保留有值的触发。',\n"
       "      toggle('setConditionalFormatIncludeBlankCells',s.conditionalFormatIncludeBlankCells!==false)),\n"
       "    row('条件格式判定方式','PYTHON=直读 OOXML 规则求值（默认，与源生产一致，不启动 Office）；COM_DISPLAY=以 Excel/WPS 实际显示填充为触发事实（渲染基线，仅 Windows）。COM_EVALUATE（公式引擎逐条求值）尚未实现。',\n"
       "      '<select class=\"gen-select\" id=\"setConditionalFormatEvaluator\" onchange=\"setConditionalFormatEvaluator(this.value)\">'\n"
       "      +['PYTHON','COM_DISPLAY','COM_EVALUATE'].map(function(v){return '<option value=\"'+v+'\"'+(String(s.conditionalFormatEvaluator||'PYTHON')===v?' selected':'')+'>'+v+'</option>';}).join('')\n"
       "      +'</select>'),\n"
       "  ].join('');")
assert s.count(old) == 1, 'evaluator row anchor'
s = s.replace(old, new)

# 8) handler
old = "async function setConditionalFormatIncludeBlankCells(on){"
new = ("async function setConditionalFormatEvaluator(value){\n"
       "  if(state&&state.busy)return;\n"
       "  try{const next=await bridge.api.update({conditionalFormatEvaluator:value});if(next&&typeof next==='object')state=next;}\n"
       "  catch(e){await refresh();showToast(e&&e.message?e.message:String(e));return;}\n"
       "  if(CFG.tab==='advanced')renderAdvanced();\n"
       "}\n"
       "async function setConditionalFormatIncludeBlankCells(on){")
assert s.count(old) == 1
s = s.replace(old, new)

# 9) sync
old = "  syncToggle('setConditionalFormatIncludeBlankCells',s.conditionalFormatIncludeBlankCells!==false);"
new = ("  syncToggle('setConditionalFormatIncludeBlankCells',s.conditionalFormatIncludeBlankCells!==false);\n"
       "  syncSel('setConditionalFormatEvaluator',String(s.conditionalFormatEvaluator||'PYTHON'));")
assert s.count(old) == 1
s = s.replace(old, new)

io.open(p, 'w', encoding='utf-8', newline='').write(s)
print('index.html patched')
