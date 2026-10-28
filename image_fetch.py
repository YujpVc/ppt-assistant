#!/usr/bin/env python3
"""
图片获取 —— 关键词联网取图，用于 PPT 配图「插入/填充」。
默认走国内（必应国内图片搜索 cn.bing.com，直连、无需 key、国内站点优先）；
国内找不到 → Pexels（走代理，外网）→ loremflickr/picsum 兜底。

代理：读环境变量 PPT_PROXY，默认 http://127.0.0.1:7897（仅用于国外站点 Pexels/loremflickr/picsum）。
国内站点（cn.bing.com + 中文图床）强制直连，不受 shell 代理影响。

用法：
    python3 image_fetch.py <keyword> <out.jpg> [width] [height]
也可作为模块： from image_fetch import fetch
"""
import sys, os, json, re, urllib.request, urllib.parse

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
PROXY = os.environ.get('PPT_PROXY', 'http://127.0.0.1:7897')

# 强制直连（绕过 shell 的 http_proxy，用于国内站点）
_NO_PROXY_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
# 走代理（用于国外站点 Pexels/loremflickr/picsum）
_PROXY_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({'http': PROXY, 'https': PROXY}))


def _open(url, timeout=25, headers=None, via_proxy=False):
    req = urllib.request.Request(url, headers=headers or {})
    opener = _PROXY_OPENER if via_proxy else _NO_PROXY_OPENER
    return opener.open(req, timeout=timeout)


def _pexels_key():
    k = os.environ.get('PEXELS_API_KEY', '').strip()
    if k:
        return k
    if os.path.exists(CONFIG_PATH):
        try:
            return json.load(open(CONFIG_PATH, encoding='utf-8')).get('pexels_key', '').strip()
        except Exception:
            return ''
    return ''


def _has_cjk(s):
    return any('\u4e00' <= c <= '\u9fff' for c in s)


def _is_domestic_host(url):
    """粗略判断图片 URL 是否来自国内站点，用于中文题材优先。"""
    try:
        host = urllib.parse.urlsplit(url).netloc.lower()
    except Exception:
        return False
    if host.endswith(('.cn', '.com.cn', '.net.cn', '.org.cn')):
        return True
    return any(k in host for k in (
        'zhimg.com', 'byteimg.com', 'sinaimg.cn', 'gtimg.cn', 'bjd.com.cn',
        'suning.cn', 'zdmimg.com', 'bbtnews.com.cn', 'sohu.com', '163.com',
        'baidu.com', 'bdimg.com', 'alicdn.com', 'taobao.com', 'jd.com',
    ))


def search_pexels(keyword, key=None, per_page=1, orientation='landscape'):
    """Pexels 语义搜图（外网，走代理）。返回最匹配照片 dict，无结果返回 None。"""
    key = key or _pexels_key()
    if not key:
        return None
    q = urllib.parse.quote(keyword)
    url = f"https://api.pexels.com/v1/search?query={q}&per_page={per_page}&orientation={orientation}"
    with _open(url, timeout=25, headers={'Authorization': key, 'User-Agent': 'Mozilla/5.0'}, via_proxy=True) as r:
        data = json.loads(r.read().decode('utf-8'))
    photos = data.get('photos', [])
    return photos[0] if photos else None


def search_bing_cn(keyword, count=20):
    """必应国内图片搜索（直连、无需 key）。返回完整图 URL 列表，国内站点优先。"""
    q = urllib.parse.quote(keyword)
    url = f"https://cn.bing.com/images/async?q={q}&first=0&count={count}&mmasync=1"
    with _open(url, timeout=25, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer': 'https://cn.bing.com/images/search',
    }, via_proxy=False) as r:
        t = r.read().decode('utf-8', 'ignore')
    urls = re.findall(r'murl&quot;:&quot;(.*?)&quot;', t)
    urls.sort(key=lambda u: 0 if _is_domestic_host(u) else 1)
    return urls


def _save(data, out_path):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(data)
    return out_path


def _download(url, out_path, timeout=30, via_proxy=False):
    with _open(url, timeout=timeout, headers={'User-Agent': 'Mozilla/5.0'}, via_proxy=via_proxy) as r:
        data = r.read()
    if len(data) < 1000:
        raise RuntimeError('下载内容过小')
    return _save(data, out_path)


def fetch_bing_cn(keyword, out_path, count=20):
    """从必应国内图片搜索下载第一张能成功下载的图（国内站点优先、直连）。返回 (path, src_url)。"""
    urls = search_bing_cn(keyword, count)
    for u in urls:
        if u.startswith('//'):
            u = 'https:' + u
        try:
            with _open(u, timeout=20, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                'Referer': 'https://cn.bing.com/images/search',
            }, via_proxy=False) as r:
                data = r.read()
            if len(data) < 1000:
                continue
            _save(data, out_path)
            return out_path, u
        except Exception:
            continue
    raise RuntimeError(f"必应图片搜索下载失败(关键词='{keyword}')")


def search_taobao(keyword, count=15):
    """淘宝商品图搜索（直连）：返回 alicdn 商品图 URL 列表（商品图通常是白底实拍，适合商品题材）。"""
    q = urllib.parse.quote(keyword)
    url = f"https://s.taobao.com/search?q={q}"
    with _open(url, timeout=25, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer': 'https://www.taobao.com/',
    }, via_proxy=False) as r:
        t = r.read().decode('utf-8', 'ignore')
    urls = re.findall(r'//img\.alicdn\.com/[^"\s\\]+', t)
    seen = []
    for u in urls:
        u = 'https:' + u
        if u not in seen:
            seen.append(u)
    return seen[:count]


def fetch_taobao(keyword, out_path, count=15):
    """从淘宝搜索下载第一张能成功下载的商品图。返回 (path, src_url)。"""
    urls = search_taobao(keyword, count)
    for u in urls:
        try:
            with _open(u, timeout=20, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                'Referer': 'https://s.taobao.com/',
            }, via_proxy=False) as r:
                data = r.read()
            if len(data) < 1000:
                continue
            _save(data, out_path)
            return out_path, u
        except Exception:
            continue
    raise RuntimeError(f"淘宝商品图搜索下载失败(关键词='{keyword}')")


_TRANSLATE_DICT = {
    '电动自行车': 'electric bicycle', '电瓶车': 'electric scooter moped',
    '电动车': 'electric scooter', '电动摩托车': 'electric motorcycle',
    '充电桩': 'ev charging station', '充电站': 'charging station',
    '外卖': 'food delivery rider', '头盔': 'helmet', '电池': 'battery',
    '骑行': 'cycling', '城市': 'city street', '交通': 'traffic',
}


def _deepseek_key():
    k = os.environ.get('DEEPSEEK_API_KEY', '').strip()
    if k:
        return k
    cred = os.path.expanduser('~/.dsh/.credentials.yaml')
    if os.path.exists(cred):
        try:
            for line in open(cred, encoding='utf-8'):
                if line.startswith('DEEPSEEK_API_KEY'):
                    return line.split(':', 1)[1].strip()
        except Exception:
            pass
    return ''


def _translate_en(keyword):
    """中文关键词 → 英文（先查内置词表，再走 DeepSeek 翻译），供 Pexels 兜底用。"""
    if not _has_cjk(keyword):
        return keyword
    if keyword in _TRANSLATE_DICT:
        return _TRANSLATE_DICT[keyword]
    dk = _deepseek_key()
    if not dk:
        return keyword
    try:
        payload = {
            'model': 'deepseek-v4-flash',
            'messages': [{'role': 'user', 'content': f'把下面中文关键词翻译成英文，只输出英文（3 个单词以内，不要解释）：{keyword}'}],
        }
        with _open('https://api.deepseek.com/chat/completions', timeout=20,
                   headers={'Authorization': 'Bearer ' + dk, 'Content-Type': 'application/json'},
                   via_proxy=True) as r:
            d = json.loads(r.read().decode('utf-8'))
        en = d['choices'][0]['message']['content'].strip().strip('"')
        return en if en else keyword
    except Exception:
        return keyword


def fetch(keyword, out_path, w=1600, h=900, key=None):
    """下载配图，返回 (path, attribution)。默认走国内（必应国内），找不到再走 Pexels（英文关键词）。"""
    key = key or _pexels_key()
    # 1) 默认：必应国内搜索（直连、国内站点优先，国内题材不违和）
    try:
        path, _ = fetch_bing_cn(keyword, out_path)
        return path, ""
    except Exception:
        pass
    # 2) 淘宝商品图（商品题材更贴切：白底实拍；图片偏小，作补充源）
    try:
        path, _ = fetch_taobao(keyword, out_path)
        return path, ""
    except Exception:
        pass
    # 3) 国内都找不到 → Pexels（外网、走代理；中文关键词先翻成英文，命中更准）
    try:
        en = _translate_en(keyword)
        ph = search_pexels(en, key=key)
        if ph:
            src = ph['src'].get('large2x') or ph['src'].get('large') or ph['src'].get('original')
            _download(src, out_path, via_proxy=True)
            return out_path, ""
    except Exception:
        pass
    # 3) 都不行就报错（不再走 loremflickr/picsum：中文关键词只会拿到随机图）
    raise RuntimeError(f"图片获取失败(关键词='{keyword}')，请换关键词或人工指定图片")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python3 image_fetch.py <keyword> <out.jpg> [width] [height]")
        sys.exit(1)
    kw = sys.argv[1]; out = sys.argv[2]
    w = int(sys.argv[3]) if len(sys.argv) > 3 else 1600
    h = int(sys.argv[4]) if len(sys.argv) > 4 else 900
    try:
        path, attr = fetch(kw, out, w, h)
        print(f"已保存: {path}")
        if attr:
            print(f"署名: {attr}")
    except Exception as e:
        print(f"失败: {e}")
        sys.exit(2)
