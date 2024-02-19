
namespace BalanceApp
{
    public static class Global
    {
        public static string PATTERN { get; } = @"\S*\d{2}\S+( M|( ?[а-я]{1,3}))?\b";
        public static string DirectoryToSave { set; get; } = Directory.GetCurrentDirectory() + '\\';
    }
}
