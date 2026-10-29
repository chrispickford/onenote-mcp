"""The error type every layer raises; its code prefixes the message."""


class OneNoteError(Exception):
    """Raised when a OneNote COM call fails.

    Carries a machine-readable ``code`` so callers (and the model) can tell a
    transient timeout apart from a hard backend error without parsing prose:

      timeout        OneNote did not respond within the per-operation deadline.
      backend_error  OneNote returned a COM failure (bad id, locked content, …).
      bad_request    The request itself was malformed (e.g. unknown handle).
      config_error   The server's environment configuration is invalid.
      partial_write  A replace wrote the new content, but old page objects it
                     could not delete are still on the page (listed in the
                     message).
      conflict       The page changed since the client read it.
      busy           A previous write to the same page is still running.
    """

    def __init__(self, message: str, code: str = "backend_error") -> None:
        super().__init__(message)
        self.code = code

    def __str__(self) -> str:
        # Prefix the code so it survives transport as plain error text, e.g.
        # "timeout: OneNote did not respond within 20s …".
        return f"{self.code}: {super().__str__()}"
