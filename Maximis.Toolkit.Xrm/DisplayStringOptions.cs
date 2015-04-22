namespace Maximis.Toolkit.Xrm
{
    public class DisplayStringOptions
    {
        private string boolFalse = "false";
        private string boolTrue = "true";

        public string BoolFalse { get { return boolFalse; } set { value = boolFalse; } }

        public string BoolTrue { get { return boolTrue; } set { value = boolTrue; } }

        public bool CleanWhiteSpace { get; set; }

        public string DateFormat { get; set; }

        public string LookupFormat { get; set; }

        public string OptionSetFormat { get; set; }
    }
}