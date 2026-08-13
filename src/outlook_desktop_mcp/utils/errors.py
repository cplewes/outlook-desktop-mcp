"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")

# HRESULTs seen coming out of Outlook COM. The symbolic name is worth more than
# the number when a failure has to be diagnosed later from one line of a log.
MAPI_E_INVALID_ENTRYID = 0x80040107

_HRESULT_NAMES = {
    0x80020003: "DISP_E_MEMBERNOTFOUND",
    0x80020006: "DISP_E_UNKNOWNNAME",
    0x80020009: "DISP_E_EXCEPTION",
    0x80004005: "E_FAIL",
    0x80010001: "RPC_E_CALL_REJECTED",
    0x8001010A: "RPC_E_SERVERCALL_RETRYLATER",
    0x800401E3: "MK_E_UNAVAILABLE",
    0x80080005: "CO_E_SERVER_EXEC_FAILURE",
    MAPI_E_INVALID_ENTRYID: "MAPI_E_INVALID_ENTRYID",
    0x80040108: "MAPI_E_INVALID_OBJECT",
    0x8004010A: "MAPI_E_OBJECT_DELETED",
    0x8004010F: "MAPI_E_NOT_FOUND",
    0x800706BA: "RPC_S_SERVER_UNAVAILABLE",
}

_HEX = set("0123456789abcdefABCDEF")

# Enough to identify a fault, short enough that a provider string cannot turn a
# tool result into a wall of text.
_MAX_DETAIL = 200


def _label(code: int) -> str:
    name = _HRESULT_NAMES.get(code)
    return f"0x{code:08X} {name}" if name else f"0x{code:08X}"


def _clean(value) -> str:
    """One line, bounded. COM strings arrive with stray newlines and padding."""
    text = " ".join(str(value).split())
    if len(text) > _MAX_DETAIL:
        text = text[:_MAX_DETAIL] + "..."
    return text


def check_entry_id(entry_id: str) -> str:
    """Say why entry_id cannot be an Outlook EntryID, or "" if it looks plausible.

    Only catches what is decidable without MAPI. A well-formed but wrong id gets
    caught by Outlook itself; see invalid_entry_id_message.
    """
    if not entry_id:
        return "entry_id is empty"
    if entry_id.strip() != entry_id:
        return "entry_id has leading or trailing whitespace; pass it through verbatim"
    if any(c not in _HEX for c in entry_id):
        return f"entry_id is not hexadecimal ({len(entry_id)} characters)"
    if len(entry_id) % 2:
        return (
            f"entry_id has an odd length ({len(entry_id)} characters); an EntryID is "
            "whole bytes, so at least one character is missing"
        )
    return ""


def is_invalid_entry_id(e: Exception) -> bool:
    """True when Outlook rejected an EntryID as malformed.

    Arrives as DISP_E_EXCEPTION with the real code buried in EXCEPINFO's scode,
    so both have to be checked.
    """
    try:
        import pythoncom
        if not isinstance(e, pythoncom.com_error):
            return False
        hr, _msg, exc, _arg = e.args
        codes = {hr & 0xFFFFFFFF}
        if exc and len(exc) > 5 and exc[5]:
            codes.add(exc[5] & 0xFFFFFFFF)
        return MAPI_E_INVALID_ENTRYID in codes
    except Exception:
        return False


def invalid_entry_id_message(entry_id: str) -> str:
    """What to tell a caller whose EntryID Outlook would not accept.

    Observed cause: an EntryID contains a repeated 32-character block, and a
    caller retyping it rather than copying it skips from the first occurrence to
    the second, silently dropping 48 characters. The id stays hexadecimal and
    even-length, so nothing but MAPI can catch it.
    """
    return (
        f"Outlook rejected this entry_id as malformed (MAPI_E_INVALID_ENTRYID, "
        f"0x{MAPI_E_INVALID_ENTRYID:08X}). The value supplied was {len(entry_id)} "
        "characters. A correctly copied EntryID is accepted even when the item "
        "itself cannot be read, so this means the value was altered or shortened "
        "on the way here rather than that the message is missing. Re-fetch it from "
        "list_emails or search_emails and pass it through verbatim without retyping."
    )


def format_com_error(e: Exception) -> str:
    # Errors we raise deliberately already carry a usable message; wrapping them
    # in "An unexpected error occurred" would bury the part worth reading.
    if isinstance(e, ValueError):
        return str(e)

    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            code = hr & 0xFFFFFFFF

            # When the provider fills in EXCEPINFO, exc is
            # (wcode, source, description, helpfile, helpcontext, scode). The
            # description is the only part that says what actually went wrong;
            # the top-level message for DISP_E_EXCEPTION is always the useless
            # "Exception occurred."
            detail = _clean(exc[2]) if exc and len(exc) > 2 and exc[2] else ""
            source = _clean(exc[1]) if exc and len(exc) > 1 and exc[1] else ""
            scode = exc[5] & 0xFFFFFFFF if exc and len(exc) > 5 and exc[5] else None

            parts = [f"COM Error ({_label(code)}): {_clean(msg)}"]
            if detail:
                parts.append(f"detail: {detail}")
            if source:
                parts.append(f"source: {source}")
            if scode and scode != code:
                parts.append(f"scode: {_label(scode)}")

            message = " | ".join(parts)
            _logger.warning("%s", message)
            return message
    except Exception:
        pass

    # The generic branch used to say only "An unexpected error occurred.", which
    # hid an AttributeError in our own formatting code behind a message that
    # read like a MAPI fault. Name the exception.
    _logger.warning("Unexpected non-COM exception: %s: %s", type(e).__name__, e)
    return f"An unexpected error occurred: {type(e).__name__}: {_clean(e)}"
