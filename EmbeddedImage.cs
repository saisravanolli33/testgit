namespace MS.Office.SharedUtilities.Email.Models
{
    public class EmbeddedImage
    {
        public string Name { get; set; }
        public byte[] ImageBytes { get; set; }
        public string MimeType { get; set; }

        /// <summary>
        /// Model class for Embedded Images in Html Content.
        /// </summary>
        /// <param name="name">Image file name</param>
        /// <param name="mimeType">For e.g.: "image/jpeg", "image/png, etc."</param>
        /// <param name="imageBytes">Image Byte array</param>
        public EmbeddedImage(string name, string mimeType, byte[] imageBytes)
        {
            this.Name = name;
            this.MimeType = mimeType;
            this.ImageBytes = imageBytes;
        }
    }
}
