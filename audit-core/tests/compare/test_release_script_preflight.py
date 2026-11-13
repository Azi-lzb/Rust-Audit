"""Resource preflight regressions. Never builds binaries, archives, debs, or installs."""
import os
import json
import hashlib
from pathlib import Path
import shutil
import shlex
import subprocess
import tempfile
import unittest

REPO = Path(__file__).resolve().parents[3]
S12 = ['1.逐笔统计系统_配置.xlsx', '2.报表采集系统_配置.xlsx']
S3 = ['3.0大集中通用配置.xlsx', '3.1大集中执行比较_配置.xlsx',
      '3.2大集中本期数值核对_配置.xlsx', '3.3大集中指标比较拆分_配置.xlsx']
EXTERNAL = '外部文件-大集中相关数据.xlsx'
TEMPLATES = '模板文件_命名区域'


def touch(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b'synthetic preflight placeholder; never packaged')


class Fixtures(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='audit release preflight ')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config = self.root / 'audit-core/config-real'
        for name in S12 + [EXTERNAL]:
            touch(self.config / name)
        (self.config / TEMPLATES).mkdir()
        touch(self.root / 'dist/AuditTools.exe')
        touch(self.root / 'bundle_openpyxl_runtime.py')
        touch(self.root / 'crates/excel-editor/python/openpyxl_worker.py')
        self.write_build_info('s1s2')
        for edition in ['全量', 'S1S2']:
            touch(self.root / f'docs/使用说明-{edition}版.docx')
        touch(self.root / 'product/assets/app-icon.png')

    def write_build_info(self, edition, digest=None):
        exe = self.root / 'dist/AuditTools.exe'
        digest = digest if digest is not None else hashlib.sha256(exe.read_bytes()).hexdigest()
        (self.root / 'dist/AuditTools.build.json').write_text(json.dumps({
            'schema_version': 1, 'edition': edition, 'sha256': digest}), encoding='ascii')

    def no_packages(self):
        self.assertEqual(list(self.root.rglob('*.zip')), [])
        self.assertEqual(list(self.root.rglob('*.deb')), [])


@unittest.skipUnless(os.name == 'nt', 'Windows PowerShell/cmd only')
class WindowsPreflight(Fixtures):
    def run_check(self, edition, build_edition=None, digest=None):
        if (self.root / 'dist/AuditTools.exe').exists():
            self.write_build_info(build_edition or edition, digest)
        script = self.root / 'package-windows.ps1'
        script.write_bytes((REPO / script.name).read_bytes())
        result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive',
                                 '-ExecutionPolicy', 'Bypass', '-File', str(script),
                                 '-Edition', edition, '-ValidateOnly'], capture_output=True)
        self.no_packages()
        return result

    def test_wrong_edition_or_modified_exe_is_rejected(self):
        self.assertNotEqual(self.run_check('s1s2', build_edition='full').returncode, 0)
        self.assertNotEqual(self.run_check('s1s2', digest='0'*64).returncode, 0)

    def test_shortcut_script_is_not_required(self):
        self.assertFalse((self.root / '创建快捷方式.bat').exists())
        result = self.run_check('s1s2')
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_s1s2_does_not_require_s3(self):
        result = self.run_check('s1s2')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(b'configs=2', result.stdout)

    def test_full_requires_all_s3_configs(self):
        self.assertNotEqual(self.run_check('full').returncode, 0)
        for name in S3:
            touch(self.config / name)
        result = self.run_check('full')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(b'configs=6', result.stdout)
        (self.config / S3[-1]).unlink()
        self.assertNotEqual(self.run_check('full').returncode, 0)

    def test_missing_required_configuration_or_external_is_rejected(self):
        for name in [S12[0], EXTERNAL]:
            path = self.config / name
            data = path.read_bytes()
            path.unlink()
            self.assertNotEqual(self.run_check('s1s2').returncode, 0)
            path.write_bytes(data)
        (self.config / TEMPLATES).rmdir()
        self.assertNotEqual(self.run_check('s1s2').returncode, 0)

    def test_missing_runtime_bundler_is_rejected(self):
        (self.root / 'bundle_openpyxl_runtime.py').unlink()
        self.assertNotEqual(self.run_check('s1s2').returncode, 0)

    def test_old_exe_cannot_satisfy_preflight(self):
        (self.root / 'dist/AuditTools.exe').unlink()
        touch(self.root / 'dist/open-data-audit.exe')
        self.assertNotEqual(self.run_check('s1s2').returncode, 0)

    def test_legacy_and_current_entries_forward_edition_and_validate_only(self):
        for name in ['package-windows.bat', 'package.bat', 'package-windows.ps1']:
            (self.root / name).write_bytes((REPO / name).read_bytes())
        result = subprocess.run([os.environ['COMSPEC'], '/d', '/c',
                                 str(self.root / 'package-windows.bat'),
                                 '-Edition', 's1s2', '-ValidateOnly'], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(b'configs=2', result.stdout)
        self.no_packages()

    def test_legacy_bat_delegates_arguments_and_exit_code(self):
        alias = self.root / 'package-windows.bat'
        alias.write_bytes((REPO / alias.name).read_bytes())
        (self.root / 'package.bat').write_bytes(b'@echo off\r\necho PROXY:%*\r\nexit /b 17\r\n')
        result = subprocess.run([os.environ['COMSPEC'], '/d', '/c', str(alias), '-Edition', 's1s2'], capture_output=True)
        self.assertEqual(result.returncode, 17, result.stderr)
        self.assertIn(b'PROXY:-Edition s1s2', result.stdout)
        self.no_packages()


@unittest.skipUnless(os.name == 'posix' and shutil.which('bash'), 'Linux/WSL Bash only')
class LinuxPreflight(Fixtures):
    def run_check(self, edition, config=None):
        # The test invokes only the explicit resource-validation mode.
        env = os.environ.copy()
        env['CONFIG_SRC'] = str(self.config if config is None else config)
        result = subprocess.run(['bash', str(REPO / 'package-linux-deb.sh'), '--validate-only',
                                 str(self.root / 'dist/AuditTools.exe'), str(self.root / 'product'), '0.1.0',
                                 str(self.root / 'output'), edition, str(self.root / 'docs')],
                                env=env, capture_output=True, text=True)
        self.no_packages()
        self.assertFalse((self.root / 'output').exists())
        return result

    def test_s1s2_external_config_source_is_used(self):
        # No config-real exists beside product; only CONFIG_SRC contains the fixtures.
        result = self.run_check('s1s2')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('configs=2', result.stdout)
        self.assertIn(str(self.config), result.stdout)

    def test_full_requires_s3_and_declares_mutual_exclusion(self):
        self.assertNotEqual(self.run_check('full').returncode, 0)
        for name in S3:
            touch(self.config / name)
        full = self.run_check('full')
        slim = self.run_check('s1s2')
        self.assertEqual(full.returncode, 0, full.stderr)
        self.assertEqual(slim.returncode, 0, slim.stderr)
        self.assertIn('Conflicts: audittools-s1s2', full.stdout)
        self.assertIn('Replaces: audittools-s1s2', full.stdout)
        self.assertIn('Conflicts: audittools\n', slim.stdout)
        self.assertIn('Replaces: audittools\n', slim.stdout)
        self.assertIn('configs=6', full.stdout)

    def test_missing_explicit_source_never_falls_back(self):
        # A valid adjacent config directory must not override an invalid CONFIG_SRC.
        shutil.copytree(self.config, self.root / 'config-real')
        self.assertNotEqual(self.run_check('s1s2', self.root / 'does-not-exist').returncode, 0)

    def test_missing_required_config_external_or_template_is_rejected(self):
        for name in S12 + [EXTERNAL]:
            path = self.config / name
            data = path.read_bytes()
            path.unlink()
            self.assertNotEqual(self.run_check('s1s2').returncode, 0)
            path.write_bytes(data)
        (self.config / TEMPLATES).rmdir()
        self.assertNotEqual(self.run_check('s1s2').returncode, 0)

    def test_invalid_edition_is_rejected(self):
        self.assertNotEqual(self.run_check('other').returncode, 0)




@unittest.skipUnless(os.name == 'posix' and shutil.which('sh'), 'POSIX shell only')
class ConfigurationPreservation(unittest.TestCase):
    def test_snapshot_failure_stops_without_changing_active_config(self):
        with tempfile.TemporaryDirectory(prefix='audit config failure ') as temporary:
            root = Path(temporary)
            active = root / 'active'
            active.mkdir()
            (active / 's1.xlsx').write_bytes(b'user edit')
            blocked = root / 'not-a-directory'
            blocked.write_bytes(b'block snapshot destination')
            helper = REPO / 'audit-core/packaging/config-preservation.sh'
            result = subprocess.run(['sh', '-ec', '. "$1"; audit_config_snapshot "$2" "$3"',
                                     'test', str(helper), str(active), str(blocked / 'config-real')],
                                    capture_output=True, text=True)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual((active / 's1.xlsx').read_bytes(), b'user edit')

    def test_generated_deb_maintainer_scripts_preserve_legacy_config(self):
        script = (REPO / 'package-linux-deb.sh').read_text(encoding='utf-8')
        start = script.index('# Embed helpers because preinst')
        end = script.index('chmod 0755 "$PKG_ROOT/DEBIAN/postrm"', start)
        body = script[start:end] + 'chmod 0755 "$PKG_ROOT/DEBIAN/postrm"\n'
        with tempfile.TemporaryDirectory(prefix='audit maintainer scripts ') as temporary:
            root = Path(temporary)
            pkg = root / 'pkg'
            (pkg / 'DEBIAN').mkdir(parents=True)
            env = os.environ.copy()
            tools = root / 'stub-tools'
            tools.mkdir()
            for command in ['update-desktop-database', 'gtk-update-icon-cache']:
                stub = tools / command
                stub.write_text('#!/bin/sh\nexit 0\n')
                stub.chmod(0o755)
            env['PATH'] = str(tools) + os.pathsep + env.get('PATH', os.defpath)
            env['PKG_ROOT'] = str(pkg)
            env['CONFIG_PRESERVATION'] = str(REPO / 'audit-core/packaging/config-preservation.sh')
            # Execute only the actual maintainer-script generator; no package staging/dpkg.
            subprocess.run(['bash', '-eu', '-c', body], env=env, check=True, capture_output=True)
            install = root / 'install'
            state = root / 'state'
            active = install / 'config-real'
            defaults = install / 'config-defaults'
            active.mkdir(parents=True)
            defaults.mkdir()
            (active / 's1.xlsx').write_bytes(b'legacy edited')
            (defaults / 's1.xlsx').write_bytes(b'new default')
            (defaults / 's2.xlsx').write_bytes(b'missing config')
            for phase in ['preinst', 'prerm', 'postinst', 'postrm']:
                path = pkg / 'DEBIAN' / phase
                text = path.read_text().replace('/opt/audittools', shlex.quote(str(install)))
                text = text.replace('/var/lib/audittools', shlex.quote(str(state)))
                path.write_text(text)
                subprocess.run(['sh', '-n', str(path)], check=True, capture_output=True)
            def run(phase, argument):
                result = subprocess.run(['sh', str(pkg / 'DEBIAN' / phase), argument],
                                        env=env, capture_output=True, text=True)
                self.assertEqual(result.returncode, 0, result.stderr)
            run('preinst', 'upgrade')
            shutil.rmtree(active)  # Simulate legacy dpkg-owned files being removed.
            run('postinst', 'configure')
            self.assertEqual((active / 's1.xlsx').read_bytes(), b'legacy edited')
            self.assertEqual((active / 's2.xlsx').read_bytes(), b'missing config')
            (active / 's1.xlsx').write_bytes(b'edited after upgrade')
            run('prerm', 'remove')
            run('postrm', 'remove')
            self.assertEqual((active / 's1.xlsx').read_bytes(), b'edited after upgrade')
            self.assertEqual((state / 'config-real/s1.xlsx').read_bytes(), b'edited after upgrade')

    def test_upgrade_and_edition_switch_keep_edits_and_new_defaults_separate(self):
        with tempfile.TemporaryDirectory(prefix='audit config preserve ') as temporary:
            root = Path(temporary)
            active, saved, defaults = [root / name for name in ['active', 'saved/config-real', 'defaults']]
            active.mkdir()
            (active / '模板').mkdir()
            (active / 's1.xlsx').write_bytes(b'user s1')
            (active / 's3.xlsx').write_bytes(b'user s3')
            (active / '模板/custom.xlsx').write_bytes(b'user template')
            defaults.mkdir()
            (defaults / 's1.xlsx').write_bytes(b'new default')
            (defaults / 's2.xlsx').write_bytes(b'new s2')
            helper = REPO / 'audit-core/packaging/config-preservation.sh'
            def run(operation):
                result = subprocess.run(['sh', '-ec', '. "$1"; shift; ' + operation,
                                         'test', str(helper), str(active), str(saved), str(defaults)],
                                        capture_output=True, text=True)
                self.assertEqual(result.returncode, 0, result.stderr)
            run('audit_config_snapshot "$1" "$2"')
            # Simulate removal of legacy package-owned files before postinst.
            shutil.rmtree(active)
            run('audit_config_restore "$1" "$2" "$3"')
            self.assertEqual((active / 's1.xlsx').read_bytes(), b'user s1')
            self.assertEqual((active / 's3.xlsx').read_bytes(), b'user s3')
            self.assertEqual((active / '模板/custom.xlsx').read_bytes(), b'user template')
            self.assertEqual((active / 's2.xlsx').read_bytes(), b'new s2')
            self.assertEqual((defaults / 's1.xlsx').read_bytes(), b'new default')
            # Edits after the first upgrade must supersede the saved snapshot.
            (active / 's1.xlsx').write_bytes(b'new user edit')
            run('audit_config_snapshot "$1" "$2"')
            (defaults / 's3-new.xlsx').write_bytes(b'full added')
            run('audit_config_restore "$1" "$2" "$3"')
            run('audit_config_restore "$1" "$2" "$3"')
            self.assertEqual((active / 's1.xlsx').read_bytes(), b'new user edit')
            self.assertEqual((active / 's3-new.xlsx').read_bytes(), b'full added')

if __name__ == '__main__':
    unittest.main()
