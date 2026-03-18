namespace XDoc
{
    public class StatusDocx
    {
        public bool isCreated { get; set; }
    }

    public class StatusDocxWarning : StatusDocx
    {
        public string Warning { get; set; }
    }

    public class StatusDocxError : StatusDocx
    {
        public string Exception { get; set; }
    }
}
