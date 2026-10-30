"""一键部署考试系统到 Linux 服务器。

用法：
    python deploy.py --dry-run      # 只做本地预检，看要传什么
    python deploy.py                # 预检 → 打包 → 上传 → 备份 → 解包 → 重启 → 探活
    python deploy.py --skip-test    # 跳过本地回归测试

首次运行会在同目录生成 deploy.json，按实际服务器改完再执行。
"""
import argparse
import json
import os
import subprocess
import sys
import tarfile
import time

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE, 'deploy.json')
ARCHIVE = os.path.join(BASE, '_release.tar.gz')

DEFAULT_CONFIG = {
    'host': 'root@8.137.181.130',
    'ssh_port': '31318',
    'ssh_key': '~/.ssh/xiaoshizi.pem',
    'remote_dir': '',
    'remote_hint': 'student_exam_app',
    'app_port': 5000,
    'install_deps': False,
    'pip_cmd': 'pip install -q -r requirements.txt',
    'reload_cmd': 'touch wsgi.py',
    'health_path': '/api/health',
    'keep_backups': 5,
}

EXCLUDE_DIRS = {'uploads', '__pycache__', '.git', '.idea', '.vscode', 'backups', 'release'}
EXCLUDE_SUFFIX = ('.pyc', '.pyo', '.log', '.tmp')
EXCLUDE_NAMES = {'deploy.json', '_release.tar.gz', 'deploy.py'}


def load_config():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        print('已生成配置文件 deploy.json，请核对 host / ssh_key / remote_dir 后重跑')
        return None
    with open(CONFIG_PATH, encoding='utf-8') as f:
        cfg = dict(DEFAULT_CONFIG)
        cfg.update(json.load(f))
    return cfg


def run(cmd, **kwargs):
    return subprocess.run(cmd, **kwargs)


def preflight(skip_test):
    ok = True
    print('== [1/5] 本地预检 ==')
    print('   Python: %s' % sys.version.split()[0])

    graders = [f for f in os.listdir(os.path.join(BASE, 'graders'))
               if f.endswith('.py') and 'grader' in f]
    print('   评分器: %d 个' % len(graders))

    r = run([sys.executable, '-m', 'compileall', '-q', os.path.join(BASE, 'graders')])
    if r.returncode != 0:
        print('   ! 评分器语法检查失败')
        ok = False

    r = run([sys.executable, '-c', 'import flask'], capture_output=True)
    if r.returncode != 0:
        print('   ! 本机未安装 flask，无法本地自测（服务器上仍需安装）')

    questions = 0
    for subject in ('excel', 'word', 'ppt'):
        folder = os.path.join(BASE, 'materials', subject)
        if os.path.isdir(folder):
            questions += sum(1 for n in os.listdir(folder)
                             if os.path.exists(os.path.join(folder, n, '3题干.docx')))
    print('   题库: %d 套' % questions)

    if not skip_test:
        print('   回归测试中...')
        r = run([sys.executable, os.path.join(BASE, 'tools', 'selftest.py')],
                capture_output=True, text=True, encoding='utf-8', errors='replace')
        tail = [ln for ln in (r.stdout or '').splitlines() if ln.startswith('共 ')]
        print('   %s' % (tail[-1] if tail else (r.stdout or '')[-200:]))
        if r.returncode not in (0, 2):
            print('   ! 回归测试执行失败')
            ok = False
    return ok


def build_archive():
    print('== [2/5] 打包 ==')
    if os.path.exists(ARCHIVE):
        os.remove(ARCHIVE)

    def filter_func(info):
        parts = info.name.split('/')
        if any(p in EXCLUDE_DIRS for p in parts):
            return None
        if parts[-1] in EXCLUDE_NAMES or parts[-1].endswith(EXCLUDE_SUFFIX):
            return None
        if parts[-1].startswith('.') and len(parts) > 1:
            return None
        return info

    with tarfile.open(ARCHIVE, 'w:gz') as tar:
        tar.add(BASE, arcname='.', filter=filter_func)
    size = os.path.getsize(ARCHIVE) / 1024.0 / 1024.0
    print('   %s (%.1f MB)' % (os.path.basename(ARCHIVE), size))
    return True


def remote_target(cfg):
    if cfg['remote_dir']:
        return cfg['remote_dir']
    hint = cfg['remote_hint']
    return ('$(dirname "$(find /www/wwwroot /www /opt -maxdepth 5 -name app.py '
            '-path "*%s*" 2>/dev/null | head -1)")' % hint)


def remote_script(cfg):
    target = remote_target(cfg)
    port = cfg['app_port']
    steps = [
        'set -e',
        'DIR=%s' % target,
        r'if [ -z "$DIR" ] || [ "$DIR" = "." ]; then echo "! 服务器上找不到项目目录，请在 deploy.json 填 remote_dir"; exit 3; fi',
        'echo "   远端目录: $DIR"',
        'cd "$DIR"',
        'mkdir -p ../exam_backups',
        'tar czf ../exam_backups/exam_$(date +%Y%m%d_%H%M%S).tar.gz --exclude=uploads --exclude=__pycache__ . 2>/dev/null || true',
        'ls -1t ../exam_backups/exam_*.tar.gz 2>/dev/null | tail -n +%d | xargs -r rm -f' % (cfg['keep_backups'] + 1),
        'tar xzf /tmp/exam_release.tar.gz -C "$DIR"',
        'mkdir -p uploads',
    ]
    if cfg['install_deps']:
        steps.append(cfg['pip_cmd'])
    if cfg['reload_cmd']:
        steps.append(cfg['reload_cmd'])
    steps += [
        'sleep 3',
        'echo -n "   探活: "',
        'curl -s -m 8 -o /dev/null -w "%%{http_code}\\n" http://127.0.0.1:%d%s' % (port, cfg['health_path']),
        'rm -f /tmp/exam_release.tar.gz',
    ]
    return '\n'.join(steps)


def ssh_base(cfg):
    return ['ssh', '-i', os.path.expanduser(cfg['ssh_key']), '-p', str(cfg['ssh_port']),
            '-o', 'BatchMode=yes', '-o', 'StrictHostKeyChecking=accept-new', cfg['host']]


def upload(cfg):
    print('== [3/5] 上传 ==')
    r = run(['scp', '-i', os.path.expanduser(cfg['ssh_key']), '-P', str(cfg['ssh_port']),
             '-o', 'BatchMode=yes', '-o', 'StrictHostKeyChecking=accept-new',
             ARCHIVE, '%s:/tmp/exam_release.tar.gz' % cfg['host']])
    if r.returncode != 0:
        print('   ! 上传失败')
        return False
    print('   已上传 /tmp/exam_release.tar.gz')
    return True


def deploy_remote(cfg):
    print('== [4/5] 远端备份 + 解包 + 重启 ==')
    script = remote_script(cfg).replace('\r\n', '\n')
    r = run(ssh_base(cfg) + ['bash -s'], input=script.encode('utf-8'))
    if r.returncode != 0:
        print('   ! 远端执行失败，退出码 %d' % r.returncode)
        return False
    return True


def verify(cfg):
    print('== [5/5] 探活 ==')
    cmd = 'curl -s -m 8 http://127.0.0.1:%d%s' % (cfg['app_port'], cfg['health_path'])
    r = run(ssh_base(cfg) + [cmd], capture_output=True, text=True, encoding='utf-8', errors='replace')
    body = (r.stdout or '').strip()
    print('   %s' % (body[:300] if body else '(无响应)'))
    try:
        data = json.loads(body)
    except ValueError:
        print('   ! 健康检查返回不是 JSON，请到服务器看日志')
        return False
    if data.get('status') != 'ok':
        print('   ! 服务状态异常')
        return False
    print('   评分器 %s 个，题目 %s 套' % (data.get('graders'), data.get('questions')))
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true', help='只做本地预检与打包，不连服务器')
    parser.add_argument('--skip-test', action='store_true', help='跳过本地回归测试')
    args = parser.parse_args()

    cfg = load_config()
    if cfg is None:
        return 1

    if not preflight(args.skip_test):
        print('预检未通过，已中止')
        return 1
    if not build_archive():
        return 1

    if args.dry_run:
        print('== 打包完成，dry-run 结束 ==')
        print('   配置文件: %s' % CONFIG_PATH)
        print('   远端目录: %s' % (cfg['remote_dir'] or '自动查找 *%s*' % cfg['remote_hint']))
        return 0

    if not upload(cfg) or not deploy_remote(cfg):
        return 1
    time.sleep(1)
    return 0 if verify(cfg) else 1


if __name__ == '__main__':
    sys.exit(main())
