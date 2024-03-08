using System;
using System.Collections.Generic;

namespace WebApp1.Models
{
    [Serializable]
    public class Flow
    {
        public List<SearchedItem> ItemList = new List<SearchedItem>();
        public List<PriceModel> priceModels = new List<PriceModel>();
        public string Name { get; set; }
    }
}
