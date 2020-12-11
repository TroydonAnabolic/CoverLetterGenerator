using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using FindAndReplaceHelper.CoverBuilder;
using FindAndReplaceHelper.AutomatedJobApplier;

namespace FindAndReplaceHelper
{
    class Program
    {
        static void Main(string[] args)
        { // TODO: use condition logic to elect a different starting paragraph for varied role types aka admin, support, dev, game dev next week.
            // instantiate objects
            CoverBuilding coverBuilder = new CoverBuilding();
            AutomaticApplier automatedJobApplier = new AutomaticApplier();

            automatedJobApplier.BeginApplicationProcess();

            // create the cover
           // coverBuilder.StartApplication();
        }
    }
}
