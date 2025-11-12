namespace TESMEA_TMS.DTOs
{
    public class ComboBoxInfo
    {
        public string Value { get; set; }
        public string Text { get; set; }
        public ComboBoxInfo() { }
        public ComboBoxInfo(string value, string text)
        {
            Value = value;
            Text = text;
        }
    }
}
