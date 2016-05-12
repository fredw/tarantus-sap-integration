namespace TarantusSAPService.Exception
{
    class CriticalException : System.Exception
    {
        public CriticalException()
        {
        }

        public CriticalException(string message): base(message)
        {
        }

        public CriticalException(string message, System.Exception inner): base(message, inner)
        {
        }
    }
}
