/** Custom error class that extends Error and includes a status code.
 */
class CustomError extends Error {
    /** Create a custom error.
     * @param {string} message - The error message.
     * @param {number} [status=500] - The HTTP status code (default is 500).
     */
    constructor(message, status = 500) {
        super(); // Do not pass message to the super constructor
        this.message = message; // Set the error message directly
        this.status = status;
        this.name = this.constructor.name; // Set the error name to the class name
        Error.captureStackTrace(this, this.constructor); // Capture the stack trace
    }
    
  
}

module.exports = CustomError;
