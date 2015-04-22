using Maximis.Toolkit.Xrm.EntitySerialisation;
using System.Collections.Generic;

namespace Maximis.Toolkit.Xrm.ImportExport
{
    public enum CheckForExistingMode { DoNotCheck, Id, AllAttributes, AnyAttribute }

    public class ImportOptions : DeserialisationOptions
    {
        private int batchSize = 500;
        private bool continueOnError = true;

        /// <summary>
        /// The number of records to process at a time
        /// </summary>
        public int BatchSize { get { return batchSize; } set { batchSize = value; } }

        /// <summary>
        /// If set to a mode other than "DoNotCheck", an attempt will be made to find an existing record to update. Sometimes necessary but slows down imports.
        /// </summary>
        public CheckForExistingMode CheckForExisting { get; set; }

        /// <summary>
        /// If true, a create or update error will not stop the import from continuing
        /// </summary>
        public bool ContinueOnError { get { return continueOnError; } set { continueOnError = value; } }

        /// <summary>
        /// A list of attributes to use to look for an existing record to update
        /// </summary>
        public List<ExistingMatchAttribute> ExistingMatchAttributes { get; set; }
    }
}