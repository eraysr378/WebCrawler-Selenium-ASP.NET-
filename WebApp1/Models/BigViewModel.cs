using System.Collections.Generic;

namespace WebApp1.Models
{
    public class BigViewModel
    {
        public List<ExpandoModel> ExpandoModels = new List<ExpandoModel>();
        public string SortName { get; set; }
    }
}
