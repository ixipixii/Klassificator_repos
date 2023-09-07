﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_for_Pioneer
{
    public class Data
    {
        public string pnr_1 { get; set; }

        public string pnr_2 { get; set;}

        public string guid { get; set;}

        public Element element { get; set; }

        public bool flag { get; set; }
    }
}
