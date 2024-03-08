using System;
using System.Dynamic;

namespace WebApp1.Models
{
    [Serializable]
    public enum EventTypes
    {
        GoToPage,
        Click,
        SetInputValue,
        ReadInnerElements       
    }
    [Serializable]
    public class SearchedItem
    {
        static public int ItemID = 1;
        public int ItemId { get; set; }
        public string XPath { get; set; }
        public string Text { get; set; }
        public string ReadItemsId {  get; set; }
        public EventTypes EventType { get; set; }
        public bool IsGroup { get; set; }
     
        public SearchedItem()
        {
            ItemId = ItemID;
            ItemID++;
        }



    }
}
