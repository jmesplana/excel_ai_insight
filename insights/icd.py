"""Icd routes and domain operations."""
import logging
from flask import Blueprint, request, jsonify
from llm_provider import (DEFAULT_OPENAI_MODEL, build_client as build_client_for,
    LLMConfigError, get_client_and_model, provider_label, resolve_config)
from openai import OpenAI, AuthenticationError, APIError, APIConnectionError
from .errors import public_error
import os, re, io, json, zipfile, pickle, tempfile, threading, time
import requests
from concurrent.futures import ThreadPoolExecutor
ICD_MAX_WORKERS = 4
ICD_LLM_MODEL = DEFAULT_OPENAI_MODEL

logger = logging.getLogger(__name__)

bp = Blueprint("icd", __name__)

# =====================================================================
# Medical Translation (ICD-11) — Wave 1: core translation + ICD-10 mapping
# =====================================================================

ICD_RELEASE = "2026-01"
TOKEN_URL = "https://icdaccessmanagement.who.int/connect/token"
MAPPING_ZIP_URL = f"https://icdcdn.who.int/static/releasefiles/{ICD_RELEASE}/mapping.zip"

# Base for the MMS linearization of the configured release.
ICD_MMS_BASE = f"https://id.who.int/icd/release/11/{ICD_RELEASE}/mms"

# Timeouts (seconds).
ICD_HTTP_TIMEOUT = 30
ICD_MAPPING_TIMEOUT = 60

# ---------------------------------------------------------------------
# Token cache
# ---------------------------------------------------------------------
_ICD_TOKEN_CACHE = {}              # (client_id, client_secret) -> {"token", "exp"}
_ICD_TOKEN_LOCK = threading.Lock()


def get_icd_token(client_id, client_secret):
    """Return a (cached) OAuth2 access token for the WHO ICD API.

    Tokens are cached per (client_id, client_secret) and refreshed when within
    60 seconds of expiry. Raises on a failed token exchange (bad credentials).
    """
    key = (client_id, client_secret)
    now = time.time()
    cached = _ICD_TOKEN_CACHE.get(key)
    if cached and (cached["exp"] - 60) > now:
        return cached["token"]

    with _ICD_TOKEN_LOCK:
        cached = _ICD_TOKEN_CACHE.get(key)
        if cached and (cached["exp"] - 60) > now:
            return cached["token"]

        resp = requests.post(
            TOKEN_URL,
            data={
                "client_id": client_id,
                "client_secret": client_secret,
                "scope": "icdapi_access",
                "grant_type": "client_credentials",
            },
            timeout=ICD_HTTP_TIMEOUT,
        )
        resp.raise_for_status()
        payload = resp.json()
        token = payload["access_token"]
        expires_in = payload.get("expires_in", 3600)
        _ICD_TOKEN_CACHE[key] = {"token": token, "exp": now + float(expires_in)}
        return token


# ---------------------------------------------------------------------
# ICD API helpers
# ---------------------------------------------------------------------
def strip_tags(s):
    """Remove any HTML-ish tags (e.g. <em> search highlighting) from a string."""
    return re.sub(r"<[^>]+>", "", s or "")


def _icd_https(uri):
    """Normalize an ICD identifier URI to https to avoid redirects that would
    drop the Authorization header."""
    if uri and uri.startswith("http://"):
        return "https://" + uri[len("http://"):]
    return uri


def _icd_headers(token, lang):
    """Standard headers required on every ICD API call."""
    return {
        "Authorization": f"Bearer {token}",
        "API-Version": "v2",
        "Accept": "application/json",
        "Accept-Language": lang or "en",
    }


# Retry tuning for transient WHO API failures (rate limiting / timeouts).
ICD_MAX_RETRIES = 4
ICD_RETRY_BASE_DELAY = 0.6   # seconds; exponential backoff


def _icd_get(url, token, lang, params=None, timeout=ICD_HTTP_TIMEOUT):
    """GET an ICD API URL with retry + exponential backoff on transient errors.

    Retries on HTTP 429 (rate limit), 5xx, and connection/timeout errors. A 404
    is returned to the caller (some callers treat it as "not found"); other 4xx
    are raised immediately (not transient). Returns the requests.Response."""
    last_exc = None
    for attempt in range(ICD_MAX_RETRIES):
        try:
            resp = requests.get(url, headers=_icd_headers(token, lang),
                                 params=params, timeout=timeout)
            if resp.status_code == 404:
                return resp
            if resp.status_code == 429 or resp.status_code >= 500:
                # Transient: honor Retry-After if present, else back off.
                if attempt < ICD_MAX_RETRIES - 1:
                    retry_after = resp.headers.get("Retry-After")
                    try:
                        delay = float(retry_after) if retry_after else ICD_RETRY_BASE_DELAY * (2 ** attempt)
                    except (TypeError, ValueError):
                        delay = ICD_RETRY_BASE_DELAY * (2 ** attempt)
                    time.sleep(min(delay, 8))
                    continue
            resp.raise_for_status()
            return resp
        except (requests.Timeout, requests.ConnectionError) as e:
            last_exc = e
            if attempt < ICD_MAX_RETRIES - 1:
                time.sleep(ICD_RETRY_BASE_DELAY * (2 ** attempt))
                continue
            raise
    if last_exc:
        raise last_exc
    # Exhausted retries on 429/5xx: raise the last response's status.
    resp.raise_for_status()
    return resp


def _entity_id_from_uri(uri):
    """Return the trailing numeric segment of an ICD entity/linearization URI.

    Postcoordination clusters (``...&...``) and trailing modifiers are ignored;
    the last all-digit path segment is returned (matches the Foundation id used
    in the ICD-10 mapping tables)."""
    if not uri:
        return ""
    base = uri.split("&")[0]
    segments = [s for s in base.rstrip("/").split("/") if s]
    for seg in reversed(segments):
        if seg.isdigit():
            return seg
    return ""


def _icd_stem_uri(uri):
    """Reduce a (possibly postcoordinated) entity URI to its base MMS stem URI so
    it can be fetched. Postcoordination clusters ("...X & ...Y") and trailing
    non-id segments (e.g. "/unspecified") cannot be GET as a single entity."""
    ent_id = _entity_id_from_uri(uri)
    return f"{ICD_MMS_BASE}/{ent_id}" if ent_id else uri


def icd_search(token, q, lang, limit=8):
    """Flexisearch the MMS linearization. Returns up to ``limit`` candidates
    [{code, title, uri, score}] keeping only entries with a non-empty code."""
    if not q or not str(q).strip():
        return []
    url = f"{ICD_MMS_BASE}/search"
    # NOTE: medicalCodingMode is already true by default and controls which
    # properties are searched (titles + synonyms/index terms), so we do NOT pass
    # propertiesToBeSearched — combining the two returns 400 Bad Request.
    params = {
        "q": q,
        "flatResults": "true",
        "useFlexisearch": "true",
        "highlightingEnabled": "false",
    }
    resp = _icd_get(url, token, lang, params=params)
    resp.raise_for_status()
    data = resp.json()
    out = []
    for ent in (data.get("destinationEntities") or []):
        code = ent.get("theCode") or ""
        if not code:
            continue
        out.append({
            "code": code,
            "title": strip_tags(ent.get("title") or ""),
            "uri": _icd_https(ent.get("id") or ""),
            "score": ent.get("score") or 0,
        })
        if len(out) >= limit:
            break
    return out


def icd_entity_title(token, uri, lang):
    """Return the official title (``title.@value``) of an ICD entity in lang.

    The URI is reduced to its stem entity first (postcoordinated clusters can't be
    fetched). Returns "" if the entity can't be resolved (404) rather than raising,
    so an odd/foundation-only URI yields a missing term instead of blanking the row."""
    if not uri:
        return ""
    resp = _icd_get(_icd_https(_icd_stem_uri(uri)), token, lang)
    if resp.status_code == 404:
        return ""
    resp.raise_for_status()
    data = resp.json()
    title = data.get("title")
    if isinstance(title, dict):
        return title.get("@value", "") or ""
    return str(title) if title else ""


def icd_codeinfo(token, code, lang):
    """Resolve an ICD-11 MMS code to its entity (stem) URI. Returns '' if the
    code cannot be resolved."""
    if not code:
        return ""
    url = f"{ICD_MMS_BASE}/codeinfo"
    resp = _icd_get(url, token, lang, params={"code": code})
    if resp.status_code == 404:
        return ""
    resp.raise_for_status()
    data = resp.json()
    return _icd_https(data.get("stemId") or "")


def icd_autocode(token, text, lang):
    """WHO MMS autocode: map a free-text phrase to its single best ICD-11 code.

    Purpose-built for whole free-text diagnoses. Returns a candidate dict
    {code, title, uri, score} (title fetched in ``lang``) or None when there is
    no match (empty theCode)."""
    if not text or not str(text).strip():
        return None
    url = f"{ICD_MMS_BASE}/autocode"
    resp = _icd_get(url, token, lang, params={"searchText": text})
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    data = resp.json()
    code = data.get("theCode") or ""
    if not code:
        return None
    uri = _icd_https(data.get("linearizationURI") or data.get("foundationURI") or "")
    try:
        title = icd_entity_title(token, uri, lang) if uri else (data.get("matchingText") or "")
    except Exception:
        title = data.get("matchingText") or ""
    return {
        "code": code,
        "title": strip_tags(title),
        "uri": uri,
        "score": data.get("matchScore") or 0,
    }


# ---------------------------------------------------------------------
# AI disambiguation
# ---------------------------------------------------------------------
def _score_confidence(score):
    """Map a search score to a confidence bucket."""
    try:
        s = float(score or 0)
    except (TypeError, ValueError):
        s = 0.0
    if s >= 0.7:
        return "high"
    if s >= 0.4:
        return "medium"
    return "low"


def _icd_llm(llm_cfg):
    """Build (client, model) for the ICD helpers from a threaded LLM config.

    Accepts a resolved config dict, or a bare OpenAI API key string for
    backward compatibility. Returns (None, None) when the LLM is unavailable,
    which makes every ICD helper degrade to its non-LLM fallback.
    """
    if not llm_cfg:
        return None, None
    try:
        if isinstance(llm_cfg, str):
            return OpenAI(api_key=llm_cfg), ICD_LLM_MODEL
        from llm_provider import build_client
        return build_client(llm_cfg), llm_cfg.get("model") or ICD_LLM_MODEL
    except Exception:
        return None, None


def ai_pick_candidate(llm_cfg, text, candidates):
    """Pick the best ICD-11 candidate for a free-text term.

    Returns {"index": int (-1 if none), "confidence": "high|medium|low"}.
    On a missing key or any failure, falls back to candidates[0] with a
    score-derived confidence (or index -1 when there are no candidates)."""
    if not candidates:
        return {"index": -1, "confidence": "low"}

    fallback = {"index": 0, "confidence": _score_confidence(candidates[0].get("score"))}

    client, model = _icd_llm(llm_cfg)
    if not client:
        return fallback

    try:
        listing = "\n".join(
            f"{i}: [{c.get('code', '')}] {c.get('title', '')}"
            for i, c in enumerate(candidates)
        )
        prompt = (
            f'A clinician entered the following diagnosis/term:\n"{text}"\n\n'
            f"Here are candidate ICD-11 entities:\n{listing}\n\n"
            "Pick the index of the candidate that best matches the clinical meaning "
            "of the entered term. Account for spelling differences (British vs "
            "American English, e.g. diarrhoea/diarrhea, oedema/edema, anaemia/anemia, "
            "tumour/tumor), abbreviations, synonyms, and minor typos. Prefer the "
            "closest clinically-equivalent candidate even if the wording is not "
            "identical. Only use -1 if NONE of the candidates is clinically related "
            "to the term at all.\n"
            'Respond as JSON: {"index": <int>, "confidence": "high"|"medium"|"low"}'
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a medical coding assistant that maps clinical terms to ICD-11 entities."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        parsed = json.loads(response.choices[0].message.content)
        idx = int(parsed.get("index", 0))
        conf = parsed.get("confidence", "medium")
        if conf not in ("high", "medium", "low"):
            conf = "medium"
        if idx < -1 or idx >= len(candidates):
            return fallback
        return {"index": idx, "confidence": conf}
    except Exception as e:
        logger.warning("Operation failed (%s)", type(e).__name__)
        return fallback


# ---------------------------------------------------------------------
# Free-text query normalization (recall improvement)
# ---------------------------------------------------------------------
# WHO ICD content is written in British English. American spellings, typos and
# loose wording in user data otherwise return zero search hits. We search the
# raw term plus normalized/variant queries and merge the candidates.

# Common American -> British medical spelling fixes (applied as whole-word and
# as substrings for the productive endings).
_AM_TO_BR_WORDS = {
    "diarrhea": "diarrhoea", "edema": "oedema", "anemia": "anaemia",
    "ischemia": "ischaemia", "leukemia": "leukaemia", "septicemia": "septicaemia",
    "bacteremia": "bacteraemia", "hypoxemia": "hypoxaemia", "uremia": "uraemia",
    "hemorrhage": "haemorrhage", "hematoma": "haematoma", "hemophilia": "haemophilia",
    "hematuria": "haematuria", "hemangioma": "haemangioma",
    "tumor": "tumour", "labor": "labour", "behavior": "behaviour",
    "pediatric": "paediatric", "orthopedic": "orthopaedic",
    "esophagus": "oesophagus", "esophageal": "oesophageal",
    "estrogen": "oestrogen", "celiac": "coeliac", "diarrhoea": "diarrhoea",
    "gynecology": "gynaecology", "gynecological": "gynaecological",
    "edematous": "oedematous", "anesthesia": "anaesthesia",
    "hemoglobin": "haemoglobin", "hemolytic": "haemolytic", "fetal": "foetal",
}


def _spelling_variants(text):
    """Return up to a couple of British-spelling variants of an American-spelled
    term, plus generic ending fixes. Cheap, no network/AI. Empty list if no
    substitution applies."""
    if not text:
        return []
    lowered = text.lower()
    variants = set()

    # whole-word substitutions
    swapped = lowered
    for am, br in _AM_TO_BR_WORDS.items():
        if am in swapped:
            swapped = re.sub(rf"\b{re.escape(am)}\b", br, swapped)
    if swapped != lowered:
        variants.add(swapped)

    # productive endings (-emia -> -aemia, -hemo -> -haemo)
    ending = re.sub(r"([a-z])emia\b", r"\1aemia", lowered)
    ending = re.sub(r"\bhemo", "haemo", ending)
    ending = re.sub(r"\bhema", "haema", ending)
    if ending != lowered:
        variants.add(ending)

    variants.discard(lowered)
    return list(variants)


def ai_suggest_icd(llm_cfg, text, source_lang):
    """Use the LLM's medical knowledge to bridge the vocabulary gap between free
    text and ICD terminology (the official title is often worded very differently
    from how a clinician types it, e.g. "snake bite" -> "Toxic effect of venomous
    snakes" / NE83 / T63.0).

    Returns {"queries": [str], "icd11": [str], "icd10": [str], "kind": str} where
    queries are official British-spelling search phrases, icd11/icd10 are the LLM's
    best code guesses, and kind classifies the entry as "diagnosis", "procedure",
    or "other". The CALLER verifies every code against the WHO API before trusting
    it. Returns empty lists / kind "diagnosis" on a missing key or any failure."""
    empty = {"queries": [], "icd11": [], "icd10": [], "kind": "diagnosis"}
    client, model = _icd_llm(llm_cfg)
    if not client or not text:
        return empty
    try:
        prompt = (
            f'A user entered this medical entry (language code "{source_lang}"):\n'
            f'"{text}"\n\n'
            "Help map it to the WHO ICD-11 classification. The official ICD title is "
            "often worded differently from the entered term, so use your medical "
            "knowledge. Provide:\n"
            "1. kind: classify the entry as \"diagnosis\" (a disease/condition/injury), "
            "\"procedure\" (a surgical/medical intervention, e.g. wound management, "
            "reduction/fixation, episiotomy, removal of foreign body), or \"other\". "
            "NOTE: ICD-11 codes diseases, NOT procedures.\n"
            "2. queries: 1-3 official clinical search phrases (British English "
            "spelling, e.g. diarrhoea/oedema/tumour) to look it up.\n"
            "3. icd11: your best guess of the ICD-11 MMS stem code(s) for this "
            "concept (e.g. NE83, 1A40, 8B11), most likely first. [] if unsure.\n"
            "4. icd10: your best guess of the ICD-10 code(s) (e.g. T63.0). [] if unsure.\n"
            "Only give codes you are reasonably confident about; each is verified "
            "against the official API and discarded if invalid.\n"
            'Respond as JSON: {"kind": "...", "queries": [...], "icd11": [...], "icd10": [...]}'
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a medical coding assistant mapping clinical terms to ICD-11 and ICD-10."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        parsed = json.loads(response.choices[0].message.content)

        def _clean(key):
            vals = parsed.get(key) or []
            if not isinstance(vals, list):
                vals = [vals]
            return [str(v).strip() for v in vals if str(v).strip()][:3]

        kind = str(parsed.get("kind", "diagnosis") or "diagnosis").strip().lower()
        if kind not in ("diagnosis", "procedure", "other"):
            kind = "diagnosis"

        return {"queries": _clean("queries"), "icd11": _clean("icd11"),
                "icd10": _clean("icd10"), "kind": kind}
    except Exception as e:
        logger.warning("Operation failed (%s)", type(e).__name__)
        return empty


def resolve_text_anchor(token, text, source_lang, llm_cfg):
    """Resolve a free-text term to a chosen ICD-11 anchor candidate.

    Searches the raw term, cheap British-spelling variants, and (when an OpenAI
    key is present) LLM-normalized queries; merges and de-dupes the candidates,
    then lets ai_pick_candidate choose. Returns (chosen_candidate, confidence, kind)
    where kind is "diagnosis"/"procedure"/"other"; chosen is None if nothing matches.
    """
    suggestions = ai_suggest_icd(llm_cfg, text, source_lang)
    kind = suggestions.get("kind", "diagnosis")

    queries = [text]
    queries.extend(_spelling_variants(text))
    for q in suggestions["queries"]:
        if q.lower() not in {x.lower() for x in queries}:
            queries.append(q)

    merged = []
    seen = set()

    def _add(cand):
        if not cand:
            return
        key = cand.get("uri") or cand.get("code")
        if key and key not in seen:
            seen.add(key)
            merged.append(cand)

    # 1. Verify AI-suggested ICD-11 codes against the API (codeinfo). This bridges
    #    the vocabulary gap (e.g. "snake bite" -> NE83) that text search can't.
    for code in suggestions["icd11"]:
        try:
            uri = icd_codeinfo(token, code, source_lang)
            if uri:
                _add({"code": code,
                      "title": icd_entity_title(token, uri, source_lang),
                      "uri": uri, "score": 0.95})
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)

    # 2. Verify AI-suggested ICD-10 codes via the official mapping table, mapping
    #    each to its ICD-11 anchor entity.
    if suggestions["icd10"]:
        try:
            load_icd_maps()
            for code in suggestions["icd10"]:
                ent = ICD10_TO_ENTITY.get(code)
                if not ent:
                    continue
                icd11_code = ent.get("icd11Code", "")
                # Resolve via codeinfo on the ICD-11 code (the mapping table's
                # numeric id is a foundation id and often 404s as an MMS URL).
                uri = icd_codeinfo(token, icd11_code, source_lang) if icd11_code else ""
                if not uri and ent.get("entityId"):
                    uri = f"{ICD_MMS_BASE}/{ent['entityId']}"
                if uri:
                    _add({"code": icd11_code,
                          "title": icd_entity_title(token, uri, source_lang),
                          "uri": uri, "score": 0.9})
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)

    # 3. WHO autocode (purpose-built free-text -> best code) on raw + best query.
    for q in queries[:2]:
        try:
            _add(icd_autocode(token, q, source_lang))
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)

    # 4. Flexisearch over all queries.
    for q in queries:
        try:
            for c in icd_search(token, q, source_lang):
                _add(c)
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)
        if len(merged) >= 12:
            break

    if not merged:
        return None, None, kind

    merged.sort(key=lambda c: c.get("score") or 0, reverse=True)
    merged = merged[:12]
    pick = ai_pick_candidate(llm_cfg, text, merged)
    idx = pick.get("index", 0)
    if idx is None or idx < 0 or idx >= len(merged):
        return None, None, kind
    return merged[idx], pick.get("confidence", "low"), kind


# ---------------------------------------------------------------------
# ICD-10 mapping tables (downloaded, lazy-loaded, cached)
# ---------------------------------------------------------------------
ENTITY_TO_ICD10 = {}   # numeric entity id -> {"icd10Code", "icd10Title"}
ICD10_MULTI = set()    # icd10 codes that map to MULTIPLE ICD-11 categories
ICD10_TO_ENTITY = {}   # icd10 code -> {"entityId", "icd11Code"}

_ICD_MAPS_LOADED = False
_ICD_MAPS_LOCK = threading.Lock()
_ICD_MAPS_CACHE_FILE = os.path.join(tempfile.gettempdir(), f"icd_maps_{ICD_RELEASE}.pkl")


def _read_zip_text(zf, suffix):
    """Read a tab-separated member of the zip whose name ends with suffix and
    return a list of split rows (BOM-stripped). Returns [] if not present."""
    for name in zf.namelist():
        if name.endswith(suffix):
            raw = zf.read(name).decode("utf-8-sig")
            return [line.split("\t") for line in raw.splitlines() if line.strip()]
    return []


def _find_col(header, *keywords):
    """Return the index of the first header cell containing all keywords
    (case/space-insensitive), else -1."""
    for i, h in enumerate(header):
        hl = h.strip().lower().replace(" ", "").replace("-", "").replace("_", "")
        if all(k in hl for k in keywords):
            return i
    return -1


def _build_icd_maps_from_zip(content):
    """Parse the three mapping files out of the mapping.zip bytes."""
    entity_to_icd10 = {}
    icd10_multi = set()
    icd10_to_entity = {}

    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        # --- foundation_11To10MapToOneCategory.txt ---
        # header: Foundation URI | icd11Code | icd11Chapter | icd11Title |
        #         icd10Code | icd10Title
        rows = _read_zip_text(zf, "foundation_11To10MapToOneCategory.txt")
        if rows:
            header = rows[0]
            uri_i = _find_col(header, "foundation", "uri")
            if uri_i < 0:
                uri_i = 0
            code10_i = _find_col(header, "icd10", "code")
            if code10_i < 0:
                code10_i = 4
            title10_i = _find_col(header, "icd10", "title")
            if title10_i < 0:
                title10_i = 5
            for r in rows[1:]:
                if len(r) <= max(uri_i, code10_i):
                    continue
                ent_id = _entity_id_from_uri(r[uri_i])
                if not ent_id:
                    continue
                icd10_code = (r[code10_i] if len(r) > code10_i else "").strip()
                icd10_title = (r[title10_i] if len(r) > title10_i else "").strip()
                entity_to_icd10[ent_id] = {
                    "icd10Code": icd10_code,
                    "icd10Title": icd10_title,
                }

        # --- 10To11MapToMultipleCategories.txt ---
        # ICD-10 codes that map to MULTIPLE ICD-11 categories (broader than each
        # ICD-11 child).
        rows = _read_zip_text(zf, "10To11MapToMultipleCategories.txt")
        if rows:
            header = rows[0]
            code10_i = _find_col(header, "icd10", "code")
            if code10_i < 0:
                code10_i = 0
            for r in rows[1:]:
                if len(r) <= code10_i:
                    continue
                code = r[code10_i].strip()
                if code:
                    icd10_multi.add(code)

        # --- 10To11MapToOneCategory.txt ---
        # icd10 code -> entity id (+ icd11 code) for inputType=code/icd10.
        rows = _read_zip_text(zf, "10To11MapToOneCategory.txt")
        if rows:
            header = rows[0]
            code10_i = _find_col(header, "icd10", "code")
            if code10_i < 0:
                code10_i = 0
            code11_i = _find_col(header, "icd11", "code")
            uri_i = _find_col(header, "foundation", "uri")
            if uri_i < 0:
                uri_i = _find_col(header, "linearization", "releaseuri")
            if uri_i < 0:
                uri_i = _find_col(header, "releaseuri")
            for r in rows[1:]:
                if len(r) <= code10_i:
                    continue
                code = r[code10_i].strip()
                if not code:
                    continue
                uri = r[uri_i].strip() if (uri_i >= 0 and len(r) > uri_i) else ""
                ent_id = _entity_id_from_uri(uri)
                icd11_code = (r[code11_i].strip()
                              if (code11_i >= 0 and len(r) > code11_i) else "")
                # First mapping wins (one-category table).
                if code not in icd10_to_entity:
                    icd10_to_entity[code] = {
                        "entityId": ent_id,
                        "icd11Code": icd11_code,
                    }

    return entity_to_icd10, icd10_multi, icd10_to_entity


def load_icd_maps():
    """Lazily load the ICD-10 mapping tables.

    Loads from a temp-dir pickle cache if present (survives warm starts);
    otherwise downloads mapping.zip once, parses it, populates the module-level
    dicts, and best-effort writes the pickle cache. Thread-safe and cached for
    the process lifetime. A read-only filesystem (e.g. Vercel) is tolerated by
    catching write errors and keeping the in-memory cache."""
    global ENTITY_TO_ICD10, ICD10_MULTI, ICD10_TO_ENTITY, _ICD_MAPS_LOADED

    if _ICD_MAPS_LOADED:
        return

    with _ICD_MAPS_LOCK:
        if _ICD_MAPS_LOADED:
            return

        # Try the temp-dir cache first.
        try:
            if os.path.exists(_ICD_MAPS_CACHE_FILE):
                with open(_ICD_MAPS_CACHE_FILE, "rb") as f:
                    cached = pickle.load(f)
                ENTITY_TO_ICD10 = cached["entity_to_icd10"]
                ICD10_MULTI = cached["icd10_multi"]
                ICD10_TO_ENTITY = cached["icd10_to_entity"]
                _ICD_MAPS_LOADED = True
                logger.info("Loaded ICD-10 maps from cache file.")
                return
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)

        # Download and parse.
        logger.info("Downloading ICD-10 mapping tables...")
        resp = requests.get(MAPPING_ZIP_URL, timeout=ICD_MAPPING_TIMEOUT)
        resp.raise_for_status()
        e2i, multi, i2e = _build_icd_maps_from_zip(resp.content)

        ENTITY_TO_ICD10 = e2i
        ICD10_MULTI = multi
        ICD10_TO_ENTITY = i2e
        _ICD_MAPS_LOADED = True
        logger.info(
            f"Loaded ICD-10 maps: {len(e2i)} entities, {len(multi)} multi-codes, "
            f"{len(i2e)} icd10->entity."
        )

        # Best-effort write the cache for warm starts.
        try:
            with open(_ICD_MAPS_CACHE_FILE, "wb") as f:
                pickle.dump({
                    "entity_to_icd10": ENTITY_TO_ICD10,
                    "icd10_multi": ICD10_MULTI,
                    "icd10_to_entity": ICD10_TO_ENTITY,
                }, f)
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)


def derive_icd10(entity_id, icd11_code):
    """Derive an ICD-10 mapping + advisory relationship for an ICD-11 entity.

    Returns {"code", "title", "relationship"} where relationship is one of
    same-as / broader-than / narrower-than / no-map."""
    row = ENTITY_TO_ICD10.get(str(entity_id))
    if not row or not row.get("icd10Code"):
        return {"code": "", "title": "", "relationship": "no-map"}

    code = row["icd10Code"]
    if "&" in (icd11_code or ""):
        relationship = "narrower-than"        # postcoordinated ICD-11 cluster
    elif code in ICD10_MULTI:
        relationship = "broader-than"         # ICD-10 broader than ICD-11
    else:
        relationship = "same-as"
    return {"code": code, "title": row.get("icd10Title", ""), "relationship": relationship}


# ---------------------------------------------------------------------
# Per-row translation worker
# ---------------------------------------------------------------------
def _icd_no_match(row_index, note):
    return {
        "rowIndex": row_index,
        "icd11Code": "",
        "sourceTitle": "",
        "entityUri": "",
        "confidence": "none",
        "note": note,
        "outputs": {},
        "source": "",
        "kind": "diagnosis",
    }


def ai_fill_gaps(llm_cfg, text, icd11_code, anchor_title, source_lang,
                 target_lang, need_term, need_icd10):
    """Fill only the cells WHO left empty for an already-resolved ICD-11 code:
    the target-language term and/or the ICD-10 code. Returns
    {"term", "icd10Code", "relationship"} (missing keys blank). Best-effort."""
    client, model = _icd_llm(llm_cfg)
    if not client:
        return {}
    try:
        wants = []
        if need_term:
            wants.append(f'"term": the official ICD-11 term for code {icd11_code} '
                         f'translated into language code "{target_lang}"')
        if need_icd10:
            wants.append('"icd10Code": the best matching ICD-10 code, and '
                         '"relationship": one of same-as|broader-than|narrower-than')
        prompt = (
            f'The WHO ICD-11 code for "{text}" is {icd11_code} ("{anchor_title}"), '
            "but some fields are missing. Using your medical knowledge, provide:\n- "
            + "\n- ".join(wants) + "\n"
            'Respond as JSON: {"term": "", "icd10Code": "", "relationship": ""}'
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a medical coding expert in ICD-10 and ICD-11."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        parsed = json.loads(response.choices[0].message.content)
        return {
            "term": str(parsed.get("term", "") or "").strip(),
            "icd10Code": str(parsed.get("icd10Code", "") or "").strip(),
            "relationship": str(parsed.get("relationship", "") or "").strip(),
        }
    except Exception as e:
        logger.warning("Operation failed (%s)", type(e).__name__)
        return {}


def ai_full_icd(llm_cfg, text, source_lang, target_lang):
    """Last-resort: ask the LLM directly for the ICD codes + translated term when
    the WHO API found nothing. Returns a dict or None. The caller verifies the
    code against WHO and attributes the source accordingly."""
    client, model = _icd_llm(llm_cfg)
    if not client or not text:
        return None
    try:
        prompt = (
            f'The WHO ICD API could not match this medical term:\n"{text}"\n\n'
            "Using your medical knowledge, give the best ICD codes and the official "
            f'term translated into language code "{target_lang}" (use British English '
            "for any English text). Provide your best estimate even if uncertain.\n"
            'Respond as JSON: {"icd11Code": "", "icd11Term": "<term in target language>", '
            '"icd10Code": "", "relationship": "same-as|broader-than|narrower-than|no-map", '
            '"sourceTitle": "<official English term>"}'
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a medical coding expert in ICD-10 and ICD-11."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        parsed = json.loads(response.choices[0].message.content)
        return {
            "icd11Code": str(parsed.get("icd11Code", "") or "").strip(),
            "icd11Term": str(parsed.get("icd11Term", "") or "").strip(),
            "icd10Code": str(parsed.get("icd10Code", "") or "").strip(),
            "relationship": str(parsed.get("relationship", "") or "").strip(),
            "sourceTitle": str(parsed.get("sourceTitle", "") or "").strip(),
        }
    except Exception as e:
        logger.warning("Operation failed (%s)", type(e).__name__)
        return None


def _icd_llm_fallback(token, row_index, text, source_lang, target_lang, targets,
                      llm_cfg, kind="diagnosis"):
    """Build a result from the LLM when WHO has no match. If the LLM's ICD-11 code
    verifies against the WHO API, the result is upgraded to a WHO-sourced answer;
    otherwise it is returned as-is and attributed to the model (auditable)."""
    if not llm_cfg:
        res = _icd_no_match(row_index, "No ICD match found")
        res["kind"] = kind
        return res
    info = ai_full_icd(llm_cfg, text, source_lang, target_lang)
    if not info or not (info.get("icd11Code") or info.get("icd10Code")):
        res = _icd_no_match(row_index, "No ICD match found")
        res["kind"] = kind
        return res

    icd11_code = info.get("icd11Code", "")
    uri = ""
    if icd11_code:
        try:
            uri = icd_codeinfo(token, icd11_code, source_lang)
        except Exception:
            uri = ""

    outputs = {}
    if uri:
        # WHO confirms code existence; matching the input still requires review.
        source = "WHO ICD-11 API"
        confidence = "medium"
        source_title = icd_entity_title(token, uri, source_lang) or info.get("sourceTitle", "")
        entity_id = _entity_id_from_uri(uri)
        if "icd11" in targets:
            outputs["icd11"] = {"code": icd11_code, "codeSource": "WHO", "termSource": "WHO",
                                "term": icd_entity_title(token, uri, target_lang)}
        if "icd10" in targets:
            load_icd_maps()
            d = {**derive_icd10(entity_id, icd11_code), "codeSource": "WHO crosswalk"}
            if not d.get("code") and info.get("icd10Code"):
                source += " (+LLM: ICD-10)"
                d = {"code": info.get("icd10Code", ""), "title": "", "codeSource": "LLM suggestion",
                     "relationship": info.get("relationship", "") or "no-map"}
            outputs["icd10"] = d
        note = ""
    else:
        # Unverified — attribute the answer to the LLM.
        _llm_name = (llm_cfg.get("model") if isinstance(llm_cfg, dict) else None) or ICD_LLM_MODEL
        source = f"LLM ({_llm_name})"
        confidence = "low"
        source_title = info.get("sourceTitle", "")
        if "icd11" in targets:
            outputs["icd11"] = {"code": icd11_code, "codeSource": "LLM suggestion", "termSource": "LLM suggestion", "term": info.get("icd11Term", "")}
        if "icd10" in targets:
            outputs["icd10"] = {"code": info.get("icd10Code", ""), "title": "", "codeSource": "LLM suggestion",
                                "relationship": info.get("relationship", "") or "no-map"}
        note = "Provided by LLM (not found in WHO API)"

    return {
        "rowIndex": row_index,
        "icd11Code": icd11_code,
        "sourceTitle": source_title,
        "entityUri": uri,
        "confidence": confidence,
        "note": note,
        "outputs": outputs,
        "source": source,
        "kind": kind,
    }


def _icd_translate_row(row, token, params):
    """Resolve a single row to an ICD-11 anchor and build the requested target
    outputs. Never raises: any error becomes a no-match-style row."""
    row_index = row.get("rowIndex")
    try:
        text = str(row.get("text") or "").strip()
        if not text:
            return _icd_no_match(row_index, "Empty input")

        input_type = params["inputType"]
        source_system = params["sourceSystem"]
        source_lang = params["sourceLang"]
        target_lang = params["targetLang"]
        targets = params["targets"]
        llm_cfg = params["llmConfig"]

        anchor_code = ""
        anchor_uri = ""
        source_title = ""
        confidence = "high"
        note = ""
        source = "WHO ICD-11 API"
        kind = "diagnosis"

        if input_type == "text":
            chosen, chosen_conf, kind = resolve_text_anchor(
                token, text, source_lang, llm_cfg)
            if not chosen:
                # WHO found nothing — fall back to the LLM (clearly attributed).
                return _icd_llm_fallback(token, row_index, text, source_lang,
                                         target_lang, targets, llm_cfg, kind)
            anchor_code = chosen["code"]
            anchor_uri = chosen["uri"]
            source_title = chosen["title"]
            confidence = chosen_conf or "low"

        elif input_type == "code" and source_system == "mms":
            anchor_uri = icd_codeinfo(token, text, source_lang)
            if not anchor_uri:
                return _icd_no_match(row_index, "ICD-11 code could not be resolved")
            anchor_code = text
            source_title = icd_entity_title(token, anchor_uri, source_lang)

        elif input_type == "code" and source_system == "icd10":
            load_icd_maps()
            ent = ICD10_TO_ENTITY.get(text)
            if not ent:
                return _icd_no_match(row_index, "ICD-10 code not found in mapping table")
            anchor_code = ent.get("icd11Code", "")
            # Resolve via codeinfo on the ICD-11 code; fall back to the table id.
            anchor_uri = icd_codeinfo(token, anchor_code, source_lang) if anchor_code else ""
            if not anchor_uri and ent.get("entityId"):
                anchor_uri = f"{ICD_MMS_BASE}/{ent['entityId']}"
            if not anchor_uri:
                return _icd_no_match(row_index, "ICD-10 code could not be resolved to ICD-11")
            source_title = icd_entity_title(token, anchor_uri, source_lang)

        else:
            return _icd_no_match(row_index, "Unsupported input configuration")

        entity_id = _entity_id_from_uri(anchor_uri)

        outputs = {}
        if "icd11" in targets:
            outputs["icd11"] = {
                "code": anchor_code, "codeSource": "WHO", "termSource": "WHO",
                "term": icd_entity_title(token, anchor_uri, target_lang),
            }
        if "icd10" in targets:
            load_icd_maps()
            outputs["icd10"] = {**derive_icd10(entity_id, anchor_code), "codeSource": "WHO crosswalk"}

        # Gap-fill: WHO resolved a code but left the target-language term and/or the
        # ICD-10 mapping empty (common for postcoordinated codes, e.g. snake bite).
        # Ask the LLM to fill only the missing cells, and note the LLM contribution.
        need_term = "icd11" in targets and not (outputs.get("icd11", {}).get("term"))
        need_icd10 = "icd10" in targets and not (outputs.get("icd10", {}).get("code"))
        if llm_cfg and (need_term or need_icd10):
            gaps = ai_fill_gaps(llm_cfg, text, anchor_code,
                                source_title or text, source_lang, target_lang,
                                need_term, need_icd10)
            filled = []
            if need_term and gaps.get("term"):
                outputs["icd11"]["term"] = gaps["term"]
                outputs["icd11"]["termSource"] = "LLM suggestion"
                filled.append("term")
            if need_icd10 and gaps.get("icd10Code"):
                outputs["icd10"] = {"code": gaps["icd10Code"], "title": "", "codeSource": "LLM suggestion",
                                    "relationship": gaps.get("relationship") or "no-map"}
                filled.append("ICD-10")
            if filled:
                source = f"{source} (+LLM: {', '.join(filled)})"

        return {
            "rowIndex": row_index,
            "icd11Code": anchor_code,
            "sourceTitle": source_title,
            "entityUri": anchor_uri,
            "confidence": confidence,
            "note": note,
            "outputs": outputs,
            "source": source,
            "kind": kind,
        }
    except Exception as e:
        logger.error("Operation failed (%s)", type(e).__name__)
        return _icd_no_match(row_index, f"Error: {public_error(e)[:80]}")


# ---------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------
def _resolve_optional_llm(data):
    """Resolve LLM config for a route where the LLM is an optional assist.

    Returns None instead of raising when nothing is configured, so the caller
    degrades gracefully. Accepts the legacy "openaiApiKey" field.
    """
    payload = dict(data or {})
    if not payload.get('apiKey') and payload.get('openaiApiKey'):
        payload['apiKey'] = payload['openaiApiKey']
    try:
        return resolve_config(payload)
    except LLMConfigError:
        return None


@bp.route('/icd_validate', methods=['POST'])
def icd_validate():
    """Validate WHO ICD API credentials by performing a token exchange."""
    try:
        data = request.json or {}
        client_id = data.get('clientId')
        client_secret = data.get('clientSecret')
        if not client_id or not client_secret:
            return jsonify({"error": "Missing ICD API client credentials."}), 400
        try:
            get_icd_token(client_id, client_secret)
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)
            return jsonify({"error": "Invalid ICD API credentials. Please check your Client ID and Secret."}), 401
        return jsonify({"ok": True})
    except Exception as e:
        logger.error("Operation failed (%s)", type(e).__name__)
        return jsonify({"error": public_error(e)}), 500


# ---------------------------------------------------------------------
@bp.route('/icd_translate_batch', methods=['POST'])
def icd_translate_batch():
    """Translate a batch of rows to ICD-11 (and optionally derived ICD-10).

    Request body:
        {
          "clientId": str, "clientSecret": str,
          # LLM assist (optional): provider config, same shape as the other
          # routes. "openaiApiKey" is still accepted as a legacy alias.
          "provider": "openai" | "azure" | null,
          "apiKey": str | null, "azureEndpoint": str | null,
          "azureDeployment": str | null, "azureApiVersion": str | null,
          "inputType": "text" | "code",
          "sourceSystem": "mms" | "icd10",   # required when inputType == "code"
          "sourceLang": str,
          "targetLang": str,
          "targets": ["icd11", "icd10"],     # subset
          "rows": [{"rowIndex": int, "text": str}]
        }
    """
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data received in request."}), 400

        client_id = data.get('clientId')
        client_secret = data.get('clientSecret')
        if not client_id or not client_secret:
            return jsonify({"error": "Missing ICD API client credentials."}), 400

        input_type = data.get('inputType', 'text')
        if input_type not in ('text', 'code'):
            return jsonify({"error": "inputType must be 'text' or 'code'."}), 400

        source_system = data.get('sourceSystem', 'mms')
        if input_type == 'code' and source_system not in ('mms', 'icd10'):
            return jsonify({"error": "sourceSystem must be 'mms' or 'icd10' for code input."}), 400

        targets = data.get('targets') or ['icd11']
        targets = [t for t in targets if t in ('icd11', 'icd10')]
        if not targets:
            return jsonify({"error": "At least one valid target ('icd11' or 'icd10') is required."}), 400

        rows = data.get('rows') or []
        if not isinstance(rows, list):
            return jsonify({"error": "'rows' must be a list."}), 400

        # Acquire one shared token for the whole batch (also validates creds).
        try:
            token = get_icd_token(client_id, client_secret)
        except Exception as e:
            logger.warning("Operation failed (%s)", type(e).__name__)
            return jsonify({"error": "Invalid ICD API credentials."}), 401

        params = {
            "inputType": input_type,
            "sourceSystem": source_system,
            "sourceLang": data.get('sourceLang') or 'es',
            "targetLang": data.get('targetLang') or 'en',
            "targets": targets,
            # The LLM is an optional assist here: if it is not configured the
            # ICD pipeline still runs on the WHO API alone, so a config error
            # degrades to None rather than failing the whole request.
            "llmConfig": _resolve_optional_llm(data),
        }

        results = []
        error_count = 0

        request_cache = {}
        if rows:
            # Reuse completed lookups within this request only. Source diagnoses
            # must not remain in a process-wide cache across users or runs.
            cfg_key = (params["inputType"], params["sourceSystem"],
                       params["sourceLang"], params["targetLang"],
                       tuple(params["targets"]))

            def run_row(row):
                text_key = str(row.get("text") or "").strip().lower()
                ckey = (cfg_key, text_key) if text_key else None
                if ckey is not None:
                    cached = request_cache.get(ckey)
                    if cached is not None:
                        res = dict(cached)
                        res["rowIndex"] = row.get("rowIndex")
                        return res
                res = _icd_translate_row(row, token, params)
                # Cache only successful resolutions; never cache a "none" result so
                # a transient failure (rate limit/timeout) can still succeed later.
                if ckey is not None and res.get("confidence") != "none":
                    request_cache[ckey] = dict(res)
                return res

            with ThreadPoolExecutor(max_workers=ICD_MAX_WORKERS) as executor:
                for res in executor.map(run_row, rows):
                    results.append(res)
                    if res.get("confidence") == "none":
                        error_count += 1

        return jsonify({"results": results, "errors": error_count})

    except Exception as e:
        logger.error("Operation failed (%s)", type(e).__name__)
        return jsonify({"error": public_error(e)}), 500


