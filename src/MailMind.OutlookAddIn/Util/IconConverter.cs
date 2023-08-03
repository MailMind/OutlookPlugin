namespace MailMind.OutlookAddIn.Util
{
    public class IconConverter
    {
        private class AxHostImageConverter : System.Windows.Forms.AxHost
        {
            private AxHostImageConverter() : base(string.Empty) { }
            public static stdole.IPictureDisp GetPictureDispFromImage(System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
        public static stdole.IPictureDisp GetIPictureDispFromImage(System.Drawing.Image image)
        {
            return AxHostImageConverter.GetPictureDispFromImage(image);
        }
    }
}