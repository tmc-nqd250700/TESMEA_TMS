namespace TESMEA_TMS.Helpers
{
    public class BusinessException : Exception
    {
        public BusinessException(string message) : base(message) { }

        public static BusinessException UserAccountNotFound() =>
           new BusinessException("Người dùng không tồn tại");
    }
}
