"""Google Slides over MCP, using an existing Google OAuth refresh token.

The token is an `authorized_user` JSON file (client_id, client_secret, refresh_token), the format
google-workspace-mcp writes. The Slides API accepts the `drive` scope, so a token that already
carries drive needs no new consent; the Google Cloud project behind the client must have the
Slides API enabled.

The Slides API cannot read or edit Office files. `convert_office` copies a .pptx into a native
Slides file (Drive converts on copy), which the other tools can then edit; `export_pptx` writes a
native deck back out as .pptx.

Config: GOOGLE_SLIDES_CREDENTIALS = path to the authorized_user JSON.
"""
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

# The mcp package has renamed its high-level server more than once; take whichever this release has.
try:
    from mcp.server.mcpserver import MCPServer as FastMCP
except ImportError:
    try:
        from mcp.server import MCPServer as FastMCP
    except ImportError:
        try:
            from mcp.server.fastmcp import FastMCP
        except ImportError:
            from mcp.server import FastMCP

mcp = FastMCP("google-slides")
SLIDES = "https://slides.googleapis.com/v1/presentations"
DRIVE = "https://www.googleapis.com/drive/v3/files"
_token = {"value": None, "expires": 0.0}


def _access_token():
    if _token["value"] and time.time() < _token["expires"] - 60:
        return _token["value"]
    with open(os.environ["GOOGLE_SLIDES_CREDENTIALS"]) as f:
        c = json.load(f)
    body = urllib.parse.urlencode({
        "client_id": c["client_id"], "client_secret": c["client_secret"],
        "refresh_token": c["refresh_token"], "grant_type": "refresh_token"}).encode()
    r = json.loads(urllib.request.urlopen("https://oauth2.googleapis.com/token", body, timeout=30).read())
    _token["value"], _token["expires"] = r["access_token"], time.time() + int(r.get("expires_in", 3600))
    return _token["value"]


def _call(method, url, payload=None, raw=False):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": "Bearer " + _access_token(), "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            body = r.read()
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read()).get("error", {}).get("message", "")
        except Exception:
            msg = ""
        raise RuntimeError("Google API %d: %s" % (e.code, msg or e.reason)) from None
    return body if raw else (json.loads(body) if body else {})


def _text(element):
    runs = (element.get("shape", {}).get("text", {}) or {}).get("textElements", [])
    return "".join(r.get("textRun", {}).get("content", "") for r in runs).strip()


@mcp.tool()
def get_presentation(presentation_id: str, full: bool = False) -> dict:
    """Read a Slides deck. Default: title, and per slide its objectId plus each text box's
    objectId and text, which is what replace/insert edits need. full=True returns the raw API
    response (large)."""
    p = _call("GET", "%s/%s" % (SLIDES, presentation_id))
    if full:
        return p
    slides = []
    for i, s in enumerate(p.get("slides", []), 1):
        boxes = [{"objectId": e["objectId"], "text": _text(e)}
                 for e in s.get("pageElements", []) if _text(e)]
        slides.append({"index": i, "objectId": s["objectId"], "text": boxes})
    return {"presentationId": p["presentationId"], "title": p.get("title"),
            "slideCount": len(slides), "slides": slides}


@mcp.tool()
def batch_update(presentation_id: str, requests: list) -> dict:
    """Apply Slides API batchUpdate requests, as-is: replaceAllText, insertText, deleteText,
    createSlide, duplicateObject, deleteObject, updateSlidesPosition and the rest. See
    https://developers.google.com/slides/api/reference/rest/v1/presentations/request ."""
    return _call("POST", "%s/%s:batchUpdate" % (SLIDES, presentation_id), {"requests": requests})


@mcp.tool()
def replace_text(presentation_id: str, find: str, replace: str, match_case: bool = True) -> dict:
    """Replace every occurrence of `find` with `replace` across the whole deck. Returns how many
    were changed."""
    r = batch_update(presentation_id, [{"replaceAllText": {
        "containsText": {"text": find, "matchCase": match_case}, "replaceText": replace}}])
    n = sum(x.get("replaceAllText", {}).get("occurrencesChanged", 0) for x in r.get("replies", []))
    return {"occurrencesChanged": n}


@mcp.tool()
def convert_office(file_id: str, name: str = "") -> dict:
    """Copy an Office file (.pptx) into a new native Google Slides file, which the other tools can
    edit. The original is untouched. Returns the new presentation's id."""
    meta = {"mimeType": "application/vnd.google-apps.presentation"}
    if name:
        meta["name"] = name
    r = _call("POST", "%s/%s/copy?supportsAllDrives=true&fields=id,name,webViewLink" % (DRIVE, file_id), meta)
    return {"presentationId": r["id"], "name": r.get("name"), "link": r.get("webViewLink")}


@mcp.tool()
def export_pptx(presentation_id: str, upload_name: str = "") -> dict:
    """Export a native Slides deck as .pptx and upload it to Drive as a new file named
    `upload_name` (default: the deck's title + .pptx). Returns the new file's id and link."""
    data = _call("GET", "%s/%s/export?mimeType=%s" % (DRIVE, presentation_id, urllib.parse.quote(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation")), raw=True)
    title = upload_name or (_call("GET", "%s/%s?fields=name" % (DRIVE, presentation_id)).get("name", "deck") + ".pptx")
    boundary = "slidesmcp%d" % int(time.time())
    body = ("--%s\r\nContent-Type: application/json\r\n\r\n%s\r\n--%s\r\nContent-Type: "
            "application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n"
            % (boundary, json.dumps({"name": title}), boundary)).encode() + data + ("\r\n--%s--" % boundary).encode()
    req = urllib.request.Request("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink",
                                 data=body, method="POST", headers={
                                     "Authorization": "Bearer " + _access_token(),
                                     "Content-Type": "multipart/related; boundary=" + boundary})
    r = json.loads(urllib.request.urlopen(req, timeout=300).read())
    return {"fileId": r["id"], "name": r.get("name"), "link": r.get("webViewLink"), "bytes": len(data)}


def main():
    mcp.run()


if __name__ == "__main__":
    main()
