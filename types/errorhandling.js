
/**
 * A map of HTTP status codes to error messages.
 * @type {Object.<number, string>}
 */
const errorCodes = {
    400: "Bad Request",
    401: "Unauthorized",
    403: "Forbidden",
    404: "Not Found",
    405: "Method Not Allowed",
    409: "Conflict",
    415: "Unsupported Media Type",
    429: "Too Many Requests",
    500: "Internal Server Error",
    501: "Not Implemented",
    503: "Service Unavailable",
};

/**
 * Throws an error with a specified status and message.
 * @param {number} status - The HTTP status code or error code to associate with the error.
 * @param {string|null} [message=null] - Optional error message. If not provided, a default message is used.
 * @throws {Error} Throws an error object with the given status and message.
 */
module.exports.throwError = (status, message = null) => {
    const errorMessage = message || errorCodes[status] || 'Something went wrong';
    const error = new Error(errorMessage);
    /** @type {Error & {status: number}} */ (error).status = status;
    console.error(error);
    throw error;
};

// 1xx Informational Responses

// 100 Continue
// 101 Switching Protocols
// 102 Processing(WebDAV)
// 2xx Success

// 200 OK
// 201 Created
// 202 Accepted
// 203 Non - Authoritative Information
// 204 No Content
// 205 Reset Content
// 206 Partial Content
// 207 Multi - Status(WebDAV)
// 208 Already Reported(WebDAV)
// 226 IM Used
// 3xx Redirection

// 300 Multiple Choices
// 301 Moved Permanently
// 302 Found
// 303 See Other
// 304 Not Modified
// 305 Use Proxy
// 306(Unused)
// 307 Temporary Redirect
// 308 Permanent Redirect
// 4xx Client Errors

// 400 Bad Request
// 401 Unauthorized
// 402 Payment Required
// 403 Forbidden
// 404 Not Found
// 405 Method Not Allowed
// 406 Not Acceptable
// 407 Proxy Authentication Required
// 408 Request Timeout
// 409 Conflict
// 410 Gone
// 411 Length Required
// 412 Precondition Failed
// 413 Payload Too Large
// 414 URI Too Long
// 415 Unsupported Media Type
// 416 Range Not Satisfiable
// 417 Expectation Failed
// 418 I'm a teapot
// 421 Misdirected Request
// 422 Unprocessable Entity(WebDAV)
// 423 Locked(WebDAV)
// 424 Failed Dependency(WebDAV)
// 425 Too Early
// 426 Upgrade Required
// 428 Precondition Required
// 429 Too Many Requests
// 431 Request Header Fields Too Large
// 451 Unavailable For Legal Reasons
// 5xx Server Errors

// 500 Internal Server Error
// 501 Not Implemented
// 502 Bad Gateway
// 503 Service Unavailable
// 504 Gateway Timeout
// 505 HTTP Version Not Supported
// 506 Variant Also Negotiates
// 507 Insufficient Storage(WebDAV)
// 508 Loop Detected(WebDAV)
// 510 Not Extended
// 511 Network Authentication Required