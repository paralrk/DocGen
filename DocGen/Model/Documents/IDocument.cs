﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents
{
    interface IDocument
    {
        void Generate();

        List<IRow> CombineRows();

        bool IsGenerated();
    }
}
