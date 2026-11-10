#!/usr/bin/env python3
"""用 DeepSeek 官方视觉模型（deepseek-v4-flash-vision-exp）审查渲染的 PPT 页面图片。"""
import sys, base64, json, os, urllib.request

BASE = "https://api.deepseek.com"
MODEL = "deepseek-v4-flash-vision-exp"


def _key():
    """Resolve the vision key from env, then the active DSH credential store.

    Both the legacy (~/.dsh) and the current (~/.dsh-linux) homes are checked;
    the previous line scan only matched an unindented key and silently sent an
    empty bearer token when the credential lived under `refs:`.
    """
    k = os.environ.get('DEEPSEEK_API_KEY', '').strip()
    if k:
        return k
    homes = [os.environ.get('DSH_HOME', '').strip(),
             os.path.expanduser('~/.dsh-linux'),
             os.path.expanduser('~/.dsh')]
    for home in homes:
        if not home:
            continue
        cred = os.path.join(home, '.credentials.yaml')
        if not os.path.isfile(cred):
            continue
        try:
            for line in open(cred, encoding='utf-8'):
                stripped = line.strip()
                if stripped.startswith('DEEPSEEK_API_KEY:'):
                    value = stripped.split(':', 1)[1].strip().strip("'\"")
                    if value:
                        return value
        except OSError:
            continue
    return ''


def _mime(path):
    ext = os.path.splitext(path)[1].lower()
    return {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
            'webp': 'image/webp', 'gif': 'image/gif'}.get(ext, 'image/png')


def ask(image_path, question):
    img = base64.b64encode(open(image_path, 'rb').read()).decode()
    mime = _mime(image_path)
    payload = {
        'model': MODEL,
        'messages': [{'role': 'user', 'content': [
            {'type': 'text', 'text': question},
            {'type': 'image_url', 'image_url': {'url': f'data:{mime};base64,' + img}},
        ]}],
    }
    req = urllib.request.Request(BASE + '/chat/completions',
        data=json.dumps(payload).encode(),
        headers={'Authorization': 'Bearer ' + _key(), 'Content-Type': 'application/json'})
    r = urllib.request.urlopen(req, timeout=120)
    d = json.loads(r.read().decode())
    try:
        return d['choices'][0]['message']['content']
    except Exception:
        return json.dumps(d, ensure_ascii=False)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 vision_check.py <图片> [问题]")
        sys.exit(1)
    q = sys.argv[2] if len(sys.argv) > 2 else "请用中文简短点评这张PPT页面的配色、版式、有没有文字溢出/遮挡/留白过多等问题。"
    if not _key():
        print("ERR: 未找到 DEEPSEEK_API_KEY（env、$DSH_HOME、~/.dsh-linux、~/.dsh 均未命中）")
        sys.exit(2)
    try:
        print(ask(sys.argv[1], q))
    except Exception as e:
        print("ERR:", type(e).__name__, e)
        sys.exit(2)
