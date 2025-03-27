using System;
using System.Collections.Generic;

namespace TwinCATXmlGenerator.Models
{
    /// <summary>
    /// Represents a request to interact with a TwinCAT tree item.
    /// </summary>
    public class TreeItemRequest
    {
        /// <summary>
        /// The parent path of the tree item.
        /// </summary>
        public string ParentPath { get; set; }

        /// <summary>
        /// The name of the tree item.
        /// </summary>
        public string ItemName { get; set; }

        /// <summary>
        /// Attributes to modify or add to the tree item.
        /// </summary>
        public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Validates the TreeItemRequest to ensure required properties are set.
        /// </summary>
        /// <returns>True if valid; otherwise, throws an exception.</returns>
        public bool Validate()
        {
            if (string.IsNullOrEmpty(ParentPath))
            {
                throw new ArgumentException("ParentPath cannot be null or empty.");
            }

            if (string.IsNullOrEmpty(ItemName))
            {
                throw new ArgumentException("ItemName cannot be null or empty.");
            }

            return true;
        }

        /// <summary>
        /// Generates a summary of the TreeItemRequest.
        /// </summary>
        /// <returns>A string summary of the request.</returns>
        public override string ToString()
        {
            return $"ParentPath: {ParentPath}, ItemName: {ItemName}, Attributes: {string.Join(", ", Attributes)}";
        }
    }
}