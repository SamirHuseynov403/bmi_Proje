using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI
{
    public static class UserAccess
    {
        private static List<int> AuthorizedUsersExchange = new List<int> { 15, 39,45 };
        private static List<int> ReadOnlyUsersExchange = new List<int> { 45, 74 };
        private static List<int> AuthorizedUsersManagements = new List<int> { 59 };
        public static string GetUserRoleExchang(int userkod)
        {
            if (AuthorizedUsersExchange.Contains(userkod))
            {
                return "FullAccess";
            }
            if (ReadOnlyUsersExchange.Contains(userkod))
            {
                return "ReadOnly";
            }
            return "NoAccess";
        }
        public static string GetUserManagmenet(int managmenetkod)
        {
            if (AuthorizedUsersManagements.Contains(managmenetkod))
            {
                return "Management";
            }
            else
            {
                return "NoAccess";
            }
        }
    }
}
