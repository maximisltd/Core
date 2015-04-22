namespace Maximis.Toolkit.Xrm.EntitySerialisation
{
    public class DeserialisationOptions
    {
        /// <summary>
        /// If true, import data that does not map to an existing attribute will be ignored. If false, an Exception will be raised
        /// </summary>
        public bool IgnoreUnknownAttributes { get; set; }

        /// <summary>
        /// Defines a list of Attribute Names which, if an empty string value is supplied, will be explicitly set to null rather than skipped. Useful for updating addresses on existing records where e.g. Address Line 3 might already have a value but be updated with a blank value
        /// </summary>
        public string[] SetToNullIfEmpty { get; set; }
    }
}