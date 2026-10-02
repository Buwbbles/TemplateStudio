#!/usr/bin/env python3
"""Template Studio: Etsy listing publisher (Cleaning Business System).
Usage: etsy_publish.py taxonomy | draft <taxonomy_id> | shop | activate <listing_id> | show <listing_id>
Key comes from env ETSY_API_KEY (keystring:secret); OAuth token from .etsy/token.json (auto-refresh).
"""
import json, os, sys, time, uuid, urllib.request, urllib.parse, urllib.error

ROOT = "/home/snow/TemplateStudio"
TOK = f"{ROOT}/.etsy/token.json"
PROD = f"{ROOT}/products/cleaning-business-system"
SHOP = 68377150
API = "https://api.etsy.com/v3/application"
KEY = os.environ["ETSY_API_KEY"]
KEYSTRING = KEY.split(":")[0]


def load_tok():
    t = json.load(open(TOK))
    if time.time() > t.get("obtained_at", 0) + t.get("expires_in", 3600) - 120:
        body = urllib.parse.urlencode({"grant_type": "refresh_token", "client_id": KEYSTRING,
                                       "refresh_token": t["refresh_token"]}).encode()
        r = urllib.request.Request("https://api.etsy.com/v3/public/oauth/token", data=body,
                                   headers={"Content-Type": "application/x-www-form-urlencoded"})
        n = json.load(urllib.request.urlopen(r))
        t.update(n); t["obtained_at"] = int(time.time())
        fd = os.open(TOK, os.O_WRONLY | os.O_TRUNC | os.O_CREAT, 0o600)
        os.write(fd, json.dumps(t).encode()); os.close(fd)
    return t["access_token"]


def call(method, path, data=None, form=None, files=None):
    h = {"x-api-key": KEY, "Authorization": "Bearer " + load_tok()}
    body = None
    if files is not None:
        b = uuid.uuid4().hex; parts = []
        for k, v in (form or {}).items():
            parts.append(f'--{b}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
        for k, (fn, ctype, blob) in files.items():
            parts.append(f'--{b}\r\nContent-Disposition: form-data; name="{k}"; filename="{fn}"\r\nContent-Type: {ctype}\r\n\r\n'.encode() + blob + b"\r\n")
        body = b"".join(parts) + f"--{b}--\r\n".encode()
        h["Content-Type"] = f"multipart/form-data; boundary={b}"
    elif form is not None:
        body = urllib.parse.urlencode(form, doseq=True).encode()
        h["Content-Type"] = "application/x-www-form-urlencoded"
    elif data is not None:
        body = json.dumps(data).encode(); h["Content-Type"] = "application/json"
    r = urllib.request.Request(API + path, data=body, headers=h, method=method)
    try:
        return json.load(urllib.request.urlopen(r))
    except urllib.error.HTTPError as e:
        print("HTTP", e.code, method, path, e.read().decode()[:800]); sys.exit(1)


def taxonomy():
    nodes = call("GET", "/seller-taxonomy/nodes")["results"]
    def walk(ns, path):
        for n in ns:
            p = path + [n["name"]]
            if any(w in n["name"].lower() for w in ("template", "spreadsheet", "planner", "digital")):
                print(n["id"], " > ".join(p))
            walk(n.get("children", []), p)
    walk(nodes, [])


def draft(tax):
    pk = json.load(open(f"{PROD}/listing-package.json"))
    lst = call("POST", f"/shops/{SHOP}/listings", form={
        "quantity": 999, "title": pk["title"], "description": pk["description"], "price": pk["price"],
        "who_made": "i_did", "when_made": "made_to_order", "taxonomy_id": tax, "type": "download",
        "is_supply": "false", "tags": ",".join(pk["tags"])})
    lid = lst["listing_id"]; print("draft listing", lid, lst["state"])
    for i in range(1, 6):
        img = open(f"{PROD}/previews/preview-{i}.png", "rb").read()
        r = call("POST", f"/shops/{SHOP}/listings/{lid}/images",
                 form={"rank": i, "alt_text": pk["captions"][i - 1][:250]},
                 files={"image": (f"cleaning-business-system-{i}.png", "image/png", img)})
        print("image", i, r["listing_image_id"])
    pdf = open(f"{PROD}/delivery/Cleaning-Business-System-Start-Here.pdf", "rb").read()
    r = call("POST", f"/shops/{SHOP}/listings/{lid}/files", form={"name": "Cleaning-Business-System-Start-Here.pdf", "rank": 1},
             files={"file": ("Cleaning-Business-System-Start-Here.pdf", "application/pdf", pdf)})
    print("file", r.get("listing_file_id"), r.get("filename"))
    return lid


def shop():
    r = call("PUT", f"/shops/{SHOP}", form={
        "title": "Tested spreadsheet tools for small business owners",
        "announcement": "Every spreadsheet here is checked before it ships: formulas recalculated, sample data checked, and instructions included. Questions? Message us and we'll help.",
        "digital_sale_message": "Thank you! Open the PDF in your download and click the Make a copy link to get your own editable Google Sheet (you'll need a free Google account). Start on the Start Here tab. If anything doesn't work, message us and we'll fix it."})
    print("shop", r.get("shop_name"), "|", r.get("title"), "| ann:", bool(r.get("announcement")), "| dsm:", bool(r.get("digital_sale_message")))


def show(lid):
    r = call("GET", f"/listings/{lid}?includes=Images")
    f = call("GET", f"/shops/{SHOP}/listings/{lid}/files")
    print(json.dumps({"id": r["listing_id"], "state": r["state"], "title": r["title"], "price": r["price"],
                      "type": r["listing_type"], "tags": len(r["tags"]), "images": len(r.get("images", [])),
                      "files": [x["filename"] for x in f["results"]], "url": r["url"]}, indent=1))


if __name__ == "__main__":
    a = sys.argv[1:]
    if a[0] == "taxonomy": taxonomy()
    elif a[0] == "draft": draft(int(a[1]))
    elif a[0] == "shop": shop()
    elif a[0] == "activate": print(call("PATCH", f"/shops/{SHOP}/listings/{a[1]}", form={"state": "active"})["state"])
    elif a[0] == "show": show(a[1])
