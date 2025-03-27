using System;

namespace TwinCATXmlGenerator.Models
{
    /// <summary>
    /// Represents the standard response structure for TwinCAT tree item operations.
    /// </summary>
    public class TreeItemResponse
    {
        /// <summary>
        /// The status of the operation (e.g., "success", "error").
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// A human-readable message providing additional context.
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Optional: Additional data or metadata for the response.
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// Generates a response object for a successful operation.
        /// </summary>
        /// <param name="message">The message to include.</param>
        /// <param name="data">Optional additional data to include.</param>
        /// <returns>A TreeItemResponse object with "success" status.</returns>
        public static TreeItemResponse Success(string message, object data = null)
        {
            return new TreeItemResponse
            {
                Status = "success",
                Message = message,
                Data = data
            };
        }

        /// <summary>
        /// Generates a response object for an unsuccessful operation.
        /// </summary>
        /// <param name="message">The error message to include.</param>
        /// <param name="data">Optional additional data to include.</param>
        /// <returns>A TreeItemResponse object with "error" status.</returns>
        public static TreeItemResponse Error(string message, object data = null)
        {
            return new TreeItemResponse
            {
                Status = "error",
                Message = message,
                Data = data
            };
        }

        /// <summary>
        /// Provides a string representation of the response for debugging/logging purposes.
        /// </summary>
        /// <returns>A string summarizing the response object.</returns>
        public override string ToString()
        {
            return $"Status: {Status}, Message: {Message}, Data: {Data}";
        }
    }
}