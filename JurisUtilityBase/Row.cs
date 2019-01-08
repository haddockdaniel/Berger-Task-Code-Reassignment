using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class Row
    {
        public Row()
        {
            ID = "";
            matter = "";
            emp  = "";
            tkpr  = "";
            date = DateTime.Today;
            hours  = "";
            oldTask  = "";
            newTask  = "";
            entryStatus = 0;
            desc  = "";
            rowNumber = 0;
            error = "";
        }

        public string ID { get; set; }
        public string matter { get; set; }
        public string emp { get; set; }
        public string tkpr { get; set; }
        public DateTime date { get; set; }
        public string hours { get; set; }
        public string oldTask { get; set; }
        public string newTask { get; set; }
        public int entryStatus { get; set; }
        public string desc { get; set; }
        public int rowNumber { get; set; }
        public string error { get; set; }
    }
}
