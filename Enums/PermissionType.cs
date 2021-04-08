namespace ownCloud.Outlook.Enums
{
    /// <summary>
    /// Permission type when share file
    /// </summary>
    public enum PermissionType
    {
        /// <summary>
        /// Read (default for public shares)
        /// </summary>
        Read = 1,

        /// <summary>
        /// Update
        /// </summary>
        Update = 2,

        /// <summary>
        /// Create
        /// </summary>
        Create = 4,

        /// <summary>
        /// Delete
        /// </summary>
        Delete = 8,

        /// <summary>
        /// Read/write
        /// </summary>
        ReadWrite = 15
    }
}