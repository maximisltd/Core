using Microsoft.Xrm.Sdk;

namespace Maximis.Toolkit.Xrm.Annotations
{
    public class Annotation
    {
        public AnnotationAttachment Attachment { get; set; }

        public string NoteText { get; set; }

        public EntityReference Regarding { get; set; }

        public string Subject { get; set; }
    }
}