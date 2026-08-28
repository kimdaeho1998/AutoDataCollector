"""Domain-specific exceptions."""


class ServiceError(RuntimeError):
    """Base error for the collector."""


class AuthenticationError(ServiceError):
    """Raised when login fails."""


class ParseError(ServiceError):
    """Raised when HTML or JSON parsing fails."""


class HttpError(ServiceError):
    """Raised when the HTTP layer returns an unexpected response."""

