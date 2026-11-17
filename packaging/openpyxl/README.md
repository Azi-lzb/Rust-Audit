# Linux DEB 内置 openpyxl 运行时

`package-linux-deb.sh` 为全量版和 S1/S2 版都要求提供已准备好的隔离运行时，
并把当前源码中的 `openpyxl_worker.py` 一起放入 DEB。运行时缺失或校验失败时，
打包会直接失败，不会生成缺少 openpyxl 的包。BetterOffice 仍编译在 AuditTools 中。

默认运行时来源为：

```text
<product-dir>/libexec/audittools/openpyxl/runtime/
```

也可以使用 `OPENPYXL_RUNTIME_DIR` 指定其他运行时目录；无论使用哪种来源，
目录都必须包含可执行的 `bin/python3`、完整标准库及固定版本的 Python 包：

- Python 版本见 `python-version.txt`。
- openpyxl 和 et-xmlfile 版本见 `requirements.txt`。
- 运行时应可搬迁，面向 x86_64 Linux，且兼容 glibc 2.28。

运行时本身不提交到仓库。应从可信来源提前准备并校验完整性；打包脚本不联网下载，
也不会用开发机上 PATH 中的系统 Python 代替。示例：

```sh
OPENPYXL_RUNTIME_DIR=/path/to/isolated-runtime \
  bash package-linux-deb.sh /path/to/AuditTools /path/to/product 26.3.0.0 ./dist-linux full
```

`--validate-only` 只检查资源是否存在，不执行 Python 或验证运行时兼容性。
正式打包时，`stage-runtime.sh` 会复制运行时并检查解释器及库版本、隔离导入、
worker 启动、符号链接、动态库和 glibc 符号上限。打包后的布局是：

```text
/opt/audittools/libexec/audittools/openpyxl/runtime/bin/python3
/opt/audittools/libexec/audittools/openpyxl/worker/openpyxl_worker.py
```

`stage-runtime.sh` 也可单独用于传统 `/usr/libexec/<product>` 暂存布局：

```sh
bash packaging/openpyxl/stage-runtime.sh audittools /path/to/isolated-runtime /path/to/staging-root
```

运行时使用的 `libcrypt.so.1` 属于目标系统依赖；实际目标机器仍需验证其可用性。
静态依赖检查和构建机上的导入检查不能代替 glibc 2.28 环境的离线安装、worker
操作及真实麒麟/UOS 桌面验收。openpyxl 不负责计算公式；候选工作簿仍须通过
现有写后校验。
