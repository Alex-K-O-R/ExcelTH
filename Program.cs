using System;

namespace Main
{
    class Program
    {
       
        static void Main(string[] args)
        {
            Console.WriteLine("This is a demo of ExcelTH project.");
            Console.WriteLine(@"Two type of files will be created in your SYSTEM_DISK:\Reports_example\ directory.");
            Console.WriteLine(@"xl*.xlsx is a table made up with help of NetOffice lib.");
            Console.WriteLine(@"gb*.xlsx is a table made up with help of Gembox lib.");
            for (int i = 0; i < 20; i++) Console.Write("-");
            Console.WriteLine("");
            Console.WriteLine("OPTIONS:");
            Console.WriteLine("(1) Create headers demo using NetOffice lib");
            Console.WriteLine("(2) Create headers demo using Gembox lib");
            Console.WriteLine("(3) Create rows demo using NetOffice lib");
            Console.WriteLine("(4) Create rows demo using Gembox lib");

            var a = Console.ReadLine();

            if (a == "1" || a == "2")
            {
                var generated_data_as_2_dimensional_object_array = TableHandlers.ReportData.LoadDataDemo1();
                if (a == "1") Demo.ExcelDemo.GenerateDemoUsingNetOffice(generated_data_as_2_dimensional_object_array);
                if (a == "2") Demo.GemboxDemo.GenerateDemoUsingGembox(generated_data_as_2_dimensional_object_array);
            }

            if (a == "3" || a == "4")
            {
                var generated_data_as_2_dimensional_object_array = TableHandlers.ReportData.LoadDataDemo1();
                if (a == "3") Demo.ExcelDemo.GenerateDemoUsingNetOffice(generated_data_as_2_dimensional_object_array, 1);
                if (a == "4") Demo.GemboxDemo.GenerateDemoUsingGembox(generated_data_as_2_dimensional_object_array, 1);
            }
            
            Console.ReadKey();
        }

    }
}
