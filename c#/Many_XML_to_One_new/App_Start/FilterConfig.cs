using System.Web;
using System.Web.Mvc;

namespace Many_XML_to_One_new
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
