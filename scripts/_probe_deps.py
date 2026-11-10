"""临时:调研 bofire 0.5.0 的依赖要求和 breaking changes。"""
import io
import json
import time
import urllib.request
import zipfile


def pypi_json(ver):
    d = json.load(
        urllib.request.urlopen(f"https://pypi.org/pypi/bofire/{ver}/json", timeout=30)
    )
    return d


def wheel(d):
    url = [u["url"] for u in d["urls"] if u["packagetype"] == "bdist_wheel"][0]
    return zipfile.ZipFile(io.BytesIO(urllib.request.urlopen(url, timeout=90).read()))


def retry(fn, n=3):
    for i in range(n):
        try:
            return fn()
        except Exception as e:
            print(f"  retry {i}: {e}")
            time.sleep(4)
    return None


d050 = retry(lambda: pypi_json("0.5.0"))
d041 = retry(lambda: pypi_json("0.4.1"))

if d050 and d041:
    print("=== 0.5.0 requires_dist ===")
    for r in d050["info"].get("requires_dist") or []:
        print("  ", r)
    print()
    print("=== 0.4.1 requires_dist ===")
    for r in d041["info"].get("requires_dist") or []:
        print("  ", r)

# 0.5.0 的 CHANGELOG
z = retry(lambda: wheel(d050))
if z:
    names = [n for n in z.namelist() if "CHANGELOG" in n]
    print("\nCHANGELOG in wheel:", names)
    for n in names:
        txt = z.read(n).decode("utf-8", errors="replace")
        print("\n=== 0.5.0 CHANGELOG (前 120 行) ===")
        print(txt[:6000])
