namespace TarantusSAPService.Exception
{
    class SAPObjectException : System.Exception
    {
        private int docEntry;

        public SAPObjectException(string message, int docEntry) : base(message)
        {
            this.docEntry = docEntry;
        }

        public string getFormattedMessage()
        {
            return "Order #" + docEntry + ": " + this.Message;
        }
    }
}
