"""Exercise the production archive command with synthetic files, never Rust or installation."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

REPO = Path(__file__).resolve().parents[3]


@unittest.skipUnless(os.name == 'posix' and shutil.which('dpkg-deb') and shutil.which('ar'),
                     'Linux dpkg-deb/ar required')
class LegacyDebCompression(unittest.TestCase):
    def test_production_command_forces_old_dpkg_compatible_gzip(self):
        commands = [line for line in (REPO / 'package-linux-deb.sh').read_text(encoding='utf-8').splitlines()
                    if line.startswith('dpkg-deb ') and '--build' in line]
        self.assertEqual(len(commands), 1)
        with tempfile.TemporaryDirectory(prefix='audit compressor regression ') as temporary:
            root = Path(temporary)
            pkg = root / 'pkg'
            (pkg / 'DEBIAN').mkdir(parents=True, mode=0o755)
            (pkg / 'DEBIAN/control').write_text(
                'Package: audit-compression-test\nVersion: 1.0\nArchitecture: all\n'
                'Maintainer: Local test <local@example.invalid>\nDescription: Synthetic compression test\n')
            (pkg / 'usr/share/audit-test').mkdir(parents=True)
            (pkg / 'usr/share/audit-test/value.txt').write_text('synthetic data only\n')
            deb = root / 'test.deb'
            env = dict(os.environ, PKG_ROOT=str(pkg), DEB=str(deb))
            result = subprocess.run(['bash', '-e', '-c', commands[0]], env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            members = subprocess.check_output(['ar', 't', str(deb)], text=True).splitlines()
            self.assertEqual(members, ['debian-binary', 'control.tar.gz', 'data.tar.gz'])
            result = subprocess.run(['dpkg-deb', '--contents', str(deb)], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn('value.txt', result.stdout)
            info = [line for line in (REPO / 'package-linux-deb.sh').read_text(encoding='utf-8').splitlines()
                    if line.startswith('dpkg-deb --info ')]
            self.assertEqual(len(info), 1)
            result = subprocess.run(['bash', '-e', '-o', 'pipefail', '-c', info[0]],
                                    env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn('audit-compression-test', result.stdout)


if __name__ == '__main__':
    unittest.main(verbosity=2)
