using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;

namespace OutlookCom
{
    class Program
    {
        static void Main(string[] args)
        {

            //This call always seems to return a type in a GAC assembly.  How do I make it return a type in the Micrsofot.Office.Interop.Outlook DLL that I have referenced?
            var Instance = Marshal.GetActiveObject("Outlook.Application");

            if (Instance == null)
            {
                Console.WriteLine("NOT GOOD:  Null Instance.  Make sure you're running Outlook");
            } else if (System.IO.Path.GetDirectoryName(Instance.GetType().Assembly.Location.ToLower()) == System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location.ToLower()))
            {
                Console.WriteLine("GOOD!!!:  Using our version of the DLL");
            } else
            {
                Console.WriteLine("NOT GOOD:  Unexpected Path - Most likely GAC.");
            }


        }

    }
}
