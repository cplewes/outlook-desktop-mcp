"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")

# HRESULTs seen coming out of Outlook COM. The symbolic name is worth more than
# the number when a failure has to be diagnosed later from one line of a log.
_HRESULT_NAMES = {
    0x80020003: "DISP_E_MEMBERNOTFOUND",
    0x80020006: "DISP_E_UNKNOWNNAME",
    0x80020009: "DISP_E_EXCEPTION",
    0x80004005: "E_FAIL",
    0x80010001: "RPC_E_CALL_REJECTED",
    0x8001010A: "RPC_E_SERVERCALL_RETRYLATER",
    0x800401E3: "MK_E_UNAVAILABLE",
    0x80080005: "CO_E_SERVER_EXEC_FAILURE",
    0x8004010F: "MAPI_E_NOT_FOUND",
    0x800706BA: "RPC_S_SERVER_UNAVAILABLE",
}

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


def format_com_error(e: Exception) -> str:
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
